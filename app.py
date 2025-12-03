import os
import threading
import queue
import math
from dataclasses import dataclass
from typing import List, Optional, Tuple, Callable
import ctypes
import functools

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# matplotlib backend for Tk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

try:
    import humanize
    HUMANIZE = True
except Exception:
    HUMANIZE = False

@dataclass
class ItemSize:
    label: str
    path: str
    size: int
    is_dir: bool


def human_size(n: int) -> str:
    if HUMANIZE:
        return humanize.naturalsize(n, binary=True)
    # Fallback simple
    units = ["B", "KB", "MB", "GB", "TB"]
    i = 0
    v = float(n)
    while v >= 1024 and i < len(units) - 1:
        v /= 1024
        i += 1
    return f"{v:.2f} {units[i]}"


def safe_stat(path: str) -> Optional[os.stat_result]:
    try:
        return os.stat(path, follow_symlinks=False)
    except Exception:
        return None


def compute_dir_size(path: str,
                     file_filter: Optional[Callable[[str, int], bool]] = None,
                     progress_cb: Optional[Callable[[str], None]] = None,
                     cancel_cb: Optional[Callable[[], bool]] = None) -> int:
    total = 0
    for root, dirs, files in os.walk(path, topdown=True, followlinks=False):
        if cancel_cb and cancel_cb():
            break
        # Skip directories we can't access
        safe_dirs = []
        for d in dirs:
            p = os.path.join(root, d)
            s = safe_stat(p)
            if s is not None:
                safe_dirs.append(d)
        dirs[:] = safe_dirs
        for f in files:
            fp = os.path.join(root, f)
            s = safe_stat(fp)
            if s is not None:
                size = s.st_size
                if file_filter is None or file_filter(fp, size):
                    total += size
            if progress_cb:
                try:
                    progress_cb(fp)
                except Exception:
                    pass
    return total


def list_top_level_items(folder: str,
                         file_filter: Optional[Callable[[str, int], bool]] = None,
                         progress_cb: Optional[Callable[[str], None]] = None,
                         cancel_cb: Optional[Callable[[], bool]] = None) -> List[ItemSize]:
    items: List[ItemSize] = []
    try:
        with os.scandir(folder) as it:
            for entry in it:
                if entry.is_symlink():
                    continue
                if entry.is_dir(follow_symlinks=False):
                    size = compute_dir_size(entry.path, file_filter=file_filter, progress_cb=progress_cb, cancel_cb=cancel_cb)
                    items.append(ItemSize(label=entry.name, path=entry.path, size=size, is_dir=True))
                elif entry.is_file(follow_symlinks=False):
                    s = safe_stat(entry.path)
                    size = s.st_size if s else 0
                    if file_filter is None or file_filter(entry.path, size):
                        items.append(ItemSize(label=entry.name, path=entry.path, size=size, is_dir=False))
    except Exception:
        pass
    return items


class App(tk.Tk):
    def __init__(self):
        # Enable per-monitor DPI awareness on Windows to avoid blurriness
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass
        
        # Get system DPI before creating Tk window
        try:
            hdc = ctypes.windll.user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
            ctypes.windll.user32.ReleaseDC(0, hdc)
            self._system_dpi = max(96, min(240, dpi))
        except Exception:
            self._system_dpi = 100
        
        super().__init__()
        
        # Align Tk scaling with actual monitor DPI for crisp rendering
        try:
            tk_scaling = self._system_dpi / 72.0
            self.tk.call('tk', 'scaling', tk_scaling)
        except Exception:
            pass
            
        self.title("File Size Filter - Pie Chart")
        self.geometry("900x600")

        self.selected_folder: Optional[str] = None
        self.scan_thread: Optional[threading.Thread] = None
        self.scan_queue: queue.Queue = queue.Queue()
        self._cancel_flag = False
        self._last_items: List[ItemSize] = []
        
        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=10, pady=10)

        self.folder_var = tk.StringVar()
        folder_entry = ttk.Entry(top, textvariable=self.folder_var)
        folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        browse_btn = ttk.Button(top, text="Browse", command=self.browse_folder)
        browse_btn.pack(side=tk.LEFT, padx=5)

        refresh_btn = ttk.Button(top, text="Scan", command=self.start_scan)
        refresh_btn.pack(side=tk.LEFT, padx=5)

        export_btn = ttk.Button(top, text="Export CSV", command=self.export_csv)
        export_btn.pack(side=tk.LEFT, padx=5)

        controls = ttk.Frame(self)
        controls.pack(fill=tk.X, padx=10)

        ttk.Label(controls, text="Min size:").pack(side=tk.LEFT)
        self.min_size_var = tk.StringVar(value="0")
        min_entry = ttk.Entry(controls, textvariable=self.min_size_var, width=10)
        min_entry.pack(side=tk.LEFT, padx=5)
        self.size_unit_var = tk.StringVar(value="MB")
        unit_cb = ttk.Combobox(controls, textvariable=self.size_unit_var, values=["B", "KB", "MB", "GB"], width=5, state="readonly")
        unit_cb.pack(side=tk.LEFT)

        self.apply_filter_subfolders = tk.BooleanVar(value=True)
        sub_cb = ttk.Checkbutton(controls, text="Include subfolders", variable=self.apply_filter_subfolders)
        sub_cb.pack(side=tk.LEFT, padx=10)

        self.threshold_mode = tk.BooleanVar(value=True)
        th_cb = ttk.Checkbutton(controls, text="Apply threshold inside folders", variable=self.threshold_mode, command=self.update_apply_button)
        th_cb.pack(side=tk.LEFT, padx=10)

        self.apply_btn = ttk.Button(controls, text="Apply Filter", command=self.on_apply_click)
        self.apply_btn.pack(side=tk.LEFT, padx=10)

        self.status_var = tk.StringVar(value="Select a folder to begin.")
        status = ttk.Label(self, textvariable=self.status_var, anchor="w")
        status.pack(fill=tk.X, padx=10, pady=5)

        # Progress bar + cancel (hidden until scanning)
        self.prog_frame = ttk.Frame(self)
        self.prog_frame.pack(fill=tk.X, padx=10)
        self.progress = ttk.Progressbar(self.prog_frame, mode="indeterminate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cancel_btn = ttk.Button(self.prog_frame, text="Cancel", command=self.cancel_scan, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=6)
        self.prog_frame.pack_forget()

        main = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        main.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        self._paned = main

        # Table pane
        left = ttk.Frame(main, width=360)
        try:
            left.pack_propagate(False)
        except Exception:
            pass
        main.add(left, weight=1)
        
        columns = ("name", "type", "size")
        self.tree = ttk.Treeview(left, columns=columns, show="headings")
        
        self._sort_dirs = {"name": True, "type": True, "size": False}
        self._col_titles = {"name": "Name", "type": "Type", "size": "Size"}
        self.tree.heading("name", text=self._col_titles["name"], command=lambda c="name": self.sort_tree(c))
        self.tree.heading("type", text=self._col_titles["type"], command=lambda c="type": self.sort_tree(c))
        self.tree.heading("size", text=self._col_titles["size"], command=lambda c="size": self.sort_tree(c))
        self.tree.column("name", width=220, stretch=False)
        self.tree.column("type", width=70, stretch=False, anchor="center")
        self.tree.column("size", width=90, stretch=False, anchor="e")
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind('<Double-1>', self._on_show_in_explorer_selected)
        self._build_table_context_menu()

        # Chart pane
        right = ttk.Frame(main)
        main.add(right, weight=3)
        
        # Configure right panel to allow canvas to fill it
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)

        # Set figure DPI
        self.figure = Figure(figsize=(5, 4), dpi=self._system_dpi, constrained_layout=False, frameon=False, facecolor='none')
        self.ax = self.figure.add_subplot(111, frame_on=False, facecolor='none')
        
        self.canvas = FigureCanvasTkAgg(self.figure, master=right)
        self.ax.set_axis_off()
        
        # Transparent background setup
        try:
            self.ax.patch.set_visible(False)
            self.ax.set_facecolor('none')
            self.figure.patch.set_visible(False)
            self.figure.set_facecolor('none')
            self.canvas.get_tk_widget().configure(highlightthickness=0, bd=0)
        except Exception:
            pass
            
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Initial sash pos
        def _init_sash():
            try:
                if self._paned.winfo_ismapped():
                    self._paned.sashpos(0, 360)
            except Exception:
                pass
        self.after(150, _init_sash)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)
            self.selected_folder = folder
            self.start_scan()

    def parse_min_size(self) -> int:
        try:
            val = float(self.min_size_var.get().strip())
        except Exception:
            val = 0.0
        unit = self.size_unit_var.get()
        mult = 1
        if unit == "KB":
            mult = 1024
        elif unit == "MB":
            mult = 1024 ** 2
        elif unit == "GB":
            mult = 1024 ** 3
        return int(max(0, val * mult))

    def start_scan(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Invalid folder", "Please select a valid folder.")
            return
        if self.scan_thread and self.scan_thread.is_alive():
            messagebox.showinfo("Scanning", "A scan is already in progress.")
            return
        self.status_var.set("Scanning… This may take a while for large folders.")
        self.tree.delete(*self.tree.get_children())
        self.ax.clear()
        self.canvas.draw()
        self._cancel_flag = False
        self.cancel_btn.config(state=tk.NORMAL)
        self.prog_frame.pack(fill=tk.X, padx=10)
        self.progress.start(10)
        self.scan_thread = threading.Thread(target=self._scan_worker, args=(folder, self.parse_min_size(), self.apply_filter_subfolders.get(), self.threshold_mode.get()), daemon=True)
        self.scan_thread.start()
        self.after(100, self._poll_scan)

    def cancel_scan(self):
        self._cancel_flag = True

    def _scan_worker(self, folder: str, min_size: int, include_subfolders: bool, threshold_mode: bool):
        try:
            file_filter = None
            if threshold_mode:
                def _ff(path: str, size: int) -> bool:
                    return size >= min_size
                file_filter = _ff

            def _progress_cb(p: str):
                if len(p) > 0:
                    p_display = ("…" + p[-57:]) if len(p) > 60 else p
                    self.scan_queue.put(("progress", p_display))

            def _cancel_cb() -> bool:
                return self._cancel_flag

            items = list_top_level_items(folder, file_filter=file_filter, progress_cb=_progress_cb, cancel_cb=_cancel_cb)
            items = [it for it in items if it.size >= min_size]
            items.sort(key=lambda x: x.size, reverse=True)
            self.scan_queue.put(("done", items))
        except Exception as e:
            self.scan_queue.put(("error", str(e)))

    def _poll_scan(self):
        try:
            msg = self.scan_queue.get_nowait()
        except queue.Empty:
            if self.scan_thread and self.scan_thread.is_alive():
                self.after(200, self._poll_scan)
            else:
                self.status_var.set("Ready.")
            return
        kind, payload = msg
        if kind == "error":
            self.status_var.set("Error during scan.")
            messagebox.showerror("Scan error", payload)
            return
        if kind == "progress":
            self.status_var.set(f"Scanning: {payload}")
            self.after(100, self._poll_scan)
            return
        if kind == "done":
            items: List[ItemSize] = payload
            self._populate(items)
            self._last_items = items
            self.status_var.set(f"Found {len(items)} items. Total: {human_size(sum(i.size for i in items))}")
        if not (self.scan_thread and self.scan_thread.is_alive()):
            self.progress.stop()
            self.cancel_btn.config(state=tk.DISABLED)
            self.prog_frame.pack_forget()

    def _populate(self, items: List[ItemSize]):
        self.tree.delete(*self.tree.get_children())
        for it in items:
            self.tree.insert("", tk.END, values=(it.label, "Folder" if it.is_dir else "File", human_size(it.size)))
        self._label_to_path = {it.label: it.path for it in items}
        self._draw_pie(items)

    def _draw_pie(self, items: List[ItemSize]):
        self.ax.clear()
        
        # 1. Expand the Axes to fill the entire Figure canvas (rect [0,0,1,1])
        try:
            self.ax.set_position([0.0, 0.0, 1.0, 1.0])
        except Exception:
            pass

        if not items:
            self.ax.text(0.5, 0.5, "No items", ha="center", va="center")
            self.canvas.draw()
            return

        sizes = [max(0.0001, i.size) for i in items]
        labels = [i.label for i in items]

        # Consolidate small slices
        MAX_SLICES = 12
        if len(sizes) > MAX_SLICES:
            pairs = list(zip(labels, sizes))
            pairs.sort(key=lambda x: x[1], reverse=True)
            head = pairs[:MAX_SLICES-1]
            tail = pairs[MAX_SLICES-1:]
            other = sum(s for _, s in tail)
            labels = [l for l, _ in head] + ["Other"]
            sizes = [s for _, s in head] + [other]

        # Colors
        try:
            import matplotlib.cm as cm
            import numpy as np
            hashes = np.array([abs(hash(lbl)) % 256 for lbl in labels], dtype=float)
            colors = cm.tab20((hashes % 20) / 20.0)
        except Exception:
            colors = None

        # 2. Draw Pie
        # Use radius < 1.0 so anti-aliasing doesn't get clipped at the edge
        r = 0.98
        wedges, texts = self.ax.pie(
            sizes,
            labels=None,
            autopct=None,
            startangle=90,
            colors=colors,
            wedgeprops={"linewidth": 0.5, "edgecolor": "white"},
            radius=r,
            center=(0, 0),
        )

        # Allow wedges to draw outside their bounding box if needed (prevents minor clipping)
        try:
            for w in wedges:
                w.set_clip_on(False)
        except Exception:
            pass
        
        # 3. CRITICAL FIX for clipping/aspect ratio
        # 'equal' aspect ratio ensures the pie is a circle.
        # adjustable='datalim' ensures the circle stays round by changing the data limits
        # instead of shrinking the axes box. This keeps the chart full-frame.
        self.ax.set_aspect('equal', adjustable='datalim')
        self.ax.autoscale(True)
        self.ax.set_axis_off()

        # Interaction setup
        self._wedge_map = {w: lbl for w, lbl in zip(wedges, labels)}
        self._items_by_label = {i.label: i for i in items}
        self._label_to_wedge = {lbl: w for w, lbl in zip(wedges, labels)}

        # Tooltip
        if not hasattr(self, '_tooltip'):
            self._tooltip = tk.Label(self.canvas.get_tk_widget(), bg='lightyellow', fg='black', bd=1, relief='solid')

        def on_move(event):
            if event.inaxes != self.ax:
                self._tooltip.place_forget()
                return
            found = None
            for w in wedges:
                if w.contains_point((event.x, event.y)):
                    found = w
                    break
            
            # Fallback for tiny wedges (angle based)
            if found is None and event.xdata is not None and event.ydata is not None:
                try:
                    ang = math.degrees(math.atan2(event.ydata, event.xdata))
                    if ang < 0:
                        ang += 360.0
                    
                    pct_map = {}
                    for w in wedges:
                        span = abs(getattr(w, 'theta2', 0) - getattr(w, 'theta1', 0))
                        pct_map[w] = (span / 360.0) * 100.0
                    
                    MIN_PCT = 0.5
                    for w in wedges:
                        t1 = getattr(w, 'theta1', 0)
                        t2 = getattr(w, 'theta2', 0)
                        lo, hi = (t1, t2) if t1 <= t2 else (t2, t1)
                        if lo <= ang <= hi and pct_map.get(w, 0) >= MIN_PCT:
                            found = w
                            break
                except Exception:
                    pass

            # Highlight logic
            for w2 in wedges:
                w2.set_alpha(1.0)
            
            if found is not None:
                found.set_alpha(0.6)
                lbl = self._wedge_map.get(found)
                if lbl:
                    it = self._items_by_label.get(lbl)
                    total = float(sum(sizes)) if sizes else 1.0
                    pct = 0.0
                    try:
                         val = sizes[labels.index(lbl)]
                         pct = (val / total) * 100.0
                    except: 
                        pass
                    
                    tip = f"{lbl} — {human_size(it.size) if it else ''} ({pct:.1f}%)"
                    widget = self.canvas.get_tk_widget()
                    try:
                        x = int(widget.winfo_pointerx() - widget.winfo_rootx() + 12)
                        y = int(widget.winfo_pointery() - widget.winfo_rooty() + 12)
                        self._tooltip.config(text=tip)
                        self._tooltip.place(x=x, y=y)
                    except Exception:
                        pass
            else:
                self._tooltip.place_forget()
            self.canvas.draw_idle()

        # Connect events
        try:
            if hasattr(self, '_mpl_cid_hover') and self._mpl_cid_hover:
                self.canvas.mpl_disconnect(self._mpl_cid_hover)
        except Exception:
            pass
        self._mpl_cid_hover = self.canvas.mpl_connect('motion_notify_event', on_move)

        def on_click(event):
            if event.inaxes != self.ax:
                return
            for w in wedges:
                if w.contains_point((event.x, event.y)):
                    lbl = self._wedge_map.get(w)
                    it = self._items_by_label.get(lbl)
                    if it:
                        for iid in self.tree.get_children(""):
                            vals = self.tree.item(iid, "values")
                            if vals and vals[0] == it.label:
                                self.tree.selection_set(iid)
                                self.tree.see(iid)
                                break
                    break

        try:
            if hasattr(self, '_mpl_cid_click') and self._mpl_cid_click:
                self.canvas.mpl_disconnect(self._mpl_cid_click)
        except Exception:
            pass
        self._mpl_cid_click = self.canvas.mpl_connect('button_press_event', on_click)

        self.canvas.draw()

    def _build_table_context_menu(self):
        self._context_menu = tk.Menu(self, tearoff=0)
        self._context_menu.add_command(label="Open", command=self._ctx_open)
        self._context_menu.add_command(label="Show in Explorer", command=self._ctx_show_in_explorer)
        self._context_menu.add_command(label="Copy Path", command=self._ctx_copy_path)

        def _on_right_click(event):
            iid = self.tree.identify_row(event.y)
            if iid:
                self.tree.selection_set(iid)
                try:
                    self._context_menu.tk_popup(event.x_root, event.y_root)
                finally:
                    self._context_menu.grab_release()
        self.tree.bind('<Button-3>', _on_right_click)
        
        # Hover table row -> Highlight pie wedge
        def _on_tree_motion(event):
            iid = self.tree.identify_row(event.y)
            for w2 in self._label_to_wedge.values():
                w2.set_alpha(1.0)
            if iid:
                vals = self.tree.item(iid, 'values')
                if vals:
                    lbl = vals[0]
                    w = self._label_to_wedge.get(lbl)
                    if w:
                        w.set_alpha(0.6)
            self.canvas.draw_idle()
        self.tree.bind('<Motion>', _on_tree_motion)

    def _get_selected_path(self) -> Optional[str]:
        sel = self.tree.selection()
        if not sel:
            return None
        vals = self.tree.item(sel[0], 'values')
        if not vals:
            return None
        label = vals[0]
        if hasattr(self, '_label_to_path'):
            return self._label_to_path.get(label)
        return None

    def _ctx_open(self):
        path = self._get_selected_path()
        if not path:
            return
        try:
            if os.path.isdir(path):
                import subprocess
                subprocess.Popen(['explorer', os.path.normpath(path)])
            else:
                os.startfile(os.path.normpath(path))
        except Exception:
            try:
                os.startfile(os.path.normpath(path))
            except Exception:
                pass

    def _ctx_show_in_explorer(self):
        path = self._get_selected_path()
        if not path:
            return
        try:
            import subprocess
            subprocess.Popen(['explorer', '/select,', os.path.normpath(path)])
        except Exception:
            pass

    def _ctx_copy_path(self):
        path = self._get_selected_path()
        if not path:
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(path)
            self.status_var.set("Path copied to clipboard.")
        except Exception:
            pass

    def update_apply_button(self):
        if self.threshold_mode.get():
            self.apply_btn.config(text="Apply & Rescan")
        else:
            self.apply_btn.config(text="Apply Filter")

    def on_apply_click(self):
        if self.threshold_mode.get():
            self.start_scan()
        else:
            self.apply_filter_without_rescan()
            
    def apply_filter_without_rescan(self):
        # Fallback if just filtering existing results
        if not self._last_items:
            return
        min_s = self.parse_min_size()
        filtered = [it for it in self._last_items if it.size >= min_s]
        self._populate(filtered)

    def _on_show_in_explorer_selected(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], 'values')
        if not vals:
            return
        label = vals[0]
        path = self._label_to_path.get(label)
        if not path:
            return
        try:
            import subprocess
            if os.path.isdir(path):
                subprocess.Popen(['explorer', os.path.normpath(path)])
            else:
                subprocess.Popen(['explorer', '/select,', os.path.normpath(path)])
        except Exception:
            pass

    def sort_tree(self, col: str):
        asc = self._sort_dirs.get(col, True)
        data = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        if col == "size":
            def _to_bytes(txt: str) -> float:
                try:
                    num, unit = txt.split(" ")
                    val = float(num)
                    mults = {"B":1, "KB":1024, "MB":1024**2, "GB":1024**3, "TB":1024**4}
                    return val * mults.get(unit, 1)
                except Exception:
                    return 0.0
            data = [(_to_bytes(v), k) for v, k in data]
        data.sort(reverse=not asc)
        for idx, (_, k) in enumerate(data):
            self.tree.move(k, "", idx)
        arrow = "▲" if asc else "▼"
        for c in ("name", "type", "size"):
            title = self._col_titles[c]
            if c == col:
                self.tree.heading(c, text=f"{title} {arrow}")
            else:
                self.tree.heading(c, text=title)
        self._sort_dirs[col] = not asc

    def export_csv(self):
        if not self._last_items:
            messagebox.showinfo("Export", "No data to export. Run a scan first.")
            return
        default_name = "file_size_filter_export.csv"
        fp = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=default_name,
                                          filetypes=[("CSV", "*.csv")])
        if not fp:
            return
        try:
            import csv
            with open(fp, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Name", "Path", "Type", "SizeBytes", "SizeHuman"])
                for it in self._last_items:
                    writer.writerow([it.label, it.path, "Folder" if it.is_dir else "File", it.size, human_size(it.size)])
            messagebox.showinfo("Export", f"Saved: {fp}")
            try:
                import subprocess
                subprocess.Popen(["explorer", "/select,", fp])
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Export error", str(e))


if __name__ == "__main__":
    app = App()
    app.mainloop()