import os
import threading
import queue
import math
from dataclasses import dataclass
from typing import List, Optional, Tuple, Callable

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import functools

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


def list_subfolder_items(folder: str) -> List[ItemSize]:
    # Return items for every immediate child across all subfolders (first level under folder)
    items: List[ItemSize] = []
    for root, dirs, files in os.walk(folder, topdown=True, followlinks=False):
        # Only take immediate children of current root (for each subfolder level, we collect their children sizes)
        # To keep UI meaningful, we aggregate by each direct child under selected folder; deeper subfolders represented by their total size.
        if os.path.abspath(root) == os.path.abspath(folder):
            # At top level, same as list_top_level_items
            with os.scandir(root) as it:
                for entry in it:
                    if entry.is_symlink():
                        continue
                    if entry.is_dir(follow_symlinks=False):
                        size = compute_dir_size(entry.path)
                        items.append(ItemSize(label=entry.name, path=entry.path, size=size, is_dir=True))
                    elif entry.is_file(follow_symlinks=False):
                        s = safe_stat(entry.path)
                        size = s.st_size if s else 0
                        items.append(ItemSize(label=entry.name, path=entry.path, size=size, is_dir=False))
            break
    return items


class App(tk.Tk):
    def __init__(self):
        super().__init__()
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
        th_cb = ttk.Checkbutton(controls, text="Apply threshold inside folders", variable=self.threshold_mode)
        th_cb.pack(side=tk.LEFT, padx=10)

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
        # Hide progress UI initially
        self.prog_frame.pack_forget()

        main = ttk.Frame(self)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Table
        left = ttk.Frame(main)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        columns = ("name", "type", "size")
        self.tree = ttk.Treeview(left, columns=columns, show="headings")
        # Enable column sort on click
        self._sort_dirs = {"name": True, "type": True, "size": False}
        self._col_titles = {"name": "Name", "type": "Type", "size": "Size"}
        self.tree.heading("name", text=self._col_titles["name"], command=lambda c="name": self.sort_tree(c))
        self.tree.heading("type", text=self._col_titles["type"], command=lambda c="type": self.sort_tree(c))
        self.tree.heading("size", text=self._col_titles["size"], command=lambda c="size": self.sort_tree(c))
        self.tree.column("name", width=250)
        self.tree.column("type", width=80)
        self.tree.column("size", width=120)
        self.tree.pack(fill=tk.BOTH, expand=True)
        # Double-click row to open file/folder in Explorer
        self.tree.bind('<Double-1>', self._on_open_selected)

        # Chart
        right = ttk.Frame(main)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.figure = Figure(figsize=(5, 4), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.ax.set_title("Size distribution")
        self.canvas = FigureCanvasTkAgg(self.figure, master=right)
        self.ax.set_axis_off()
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

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
        self.ax.set_title("Size distribution")
        self.canvas.draw()
        self._cancel_flag = False
        self.cancel_btn.config(state=tk.NORMAL)
        # Show progress UI only during scan
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
            # If include_subfolders is false, we still show folder totals; true means the filter is applied to files when computing sizes (already totals). For simplicity, we treat totals; deeper filtering per-file would complicate meaning of slices.
            # Sort by size desc
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
            # Hide progress UI after scan ends/cancelled
            self.prog_frame.pack_forget()

    def _populate(self, items: List[ItemSize]):
        self.tree.delete(*self.tree.get_children())
        self._tree_items_ids = []
        for it in items:
            iid = self.tree.insert("", tk.END, values=(it.label, "Folder" if it.is_dir else "File", human_size(it.size)))
            self._tree_items_ids.append(iid)
        # Map labels to paths for open action
        self._label_to_path = {it.label: it.path for it in items}
        self._draw_pie(items)

    def _draw_pie(self, items: List[ItemSize]):
        self.ax.clear()
        if not items:
            self.ax.text(0.5, 0.5, "No items", ha="center", va="center")
            self.canvas.draw()
            return
        sizes = [max(0.0001, i.size) for i in items]
        labels = [i.label for i in items]
        # Limit number of slices for readability: group small ones into "Other"
        MAX_SLICES = 12
        if len(sizes) > MAX_SLICES:
            pairs = list(zip(labels, sizes))
            pairs.sort(key=lambda x: x[1], reverse=True)
            head = pairs[:MAX_SLICES-1]
            tail = pairs[MAX_SLICES-1:]
            other = sum(s for _, s in tail)
            labels = [l for l, _ in head] + ["Other"]
            sizes = [s for _, s in head] + [other]
        wedges, texts = self.ax.pie(
            sizes,
            labels=None,
            autopct=None,
            startangle=90,
            wedgeprops={"linewidth": 0.5, "edgecolor": "white"},
        )
        self.ax.axis('equal')
        self.ax.set_axis_off()  # remove axes lines
        # Build legend: Label — size and percent
        total = sum(sizes)
        legend_labels = [f"{lbl} — {human_size(sz)} ({(sz/total*100):.1f}%)" for lbl, sz in zip(labels, sizes)]
        self.ax.legend(wedges, legend_labels, loc="center left", bbox_to_anchor=(1, 0.5))
        # Map wedges to items for hover selection
        self._wedge_map = {w: lbl for w, lbl in zip(wedges, labels)}
        self._items_by_label = {i.label: i for i in items}
        # Simple tooltip using Tk label near mouse pointer
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
            for w in wedges:
                w.set_alpha(1.0)
            if found is not None:
                found.set_alpha(0.6)
                lbl = self._wedge_map.get(found)
                it = self._items_by_label.get(lbl)
                if it:
                    for iid in self.tree.get_children(""):
                        vals = self.tree.item(iid, "values")
                        if vals and vals[0] == it.label:
                            self.tree.selection_set(iid)
                            self.tree.see(iid)
                            break
                    # Show tooltip with name and size near cursor
                    tip = f"{it.label}\n{human_size(it.size)}"
                    # event.x, event.y are in display coords; place tooltip relative to canvas widget
                    widget = self.canvas.get_tk_widget()
                    try:
                        x = int(widget.winfo_pointerx() - widget.winfo_rootx() + 15)
                        y = int(widget.winfo_pointery() - widget.winfo_rooty() + 15)
                        self._tooltip.config(text=tip)
                        self._tooltip.place(x=x, y=y)
                    except Exception:
                        pass
            else:
                self._tooltip.place_forget()
            self.canvas.draw_idle()
        self._mpl_cid = self.canvas.mpl_connect('motion_notify_event', on_move)
        self.canvas.draw()

    def _on_open_selected(self, _event=None):
        # Open selected item in Explorer
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], 'values')
        if not vals:
            return
        label = vals[0]
        path = None
        if hasattr(self, '_label_to_path'):
            path = self._label_to_path.get(label)
        if not path:
            return
        try:
            import subprocess
            if os.path.isdir(path):
                subprocess.Popen(['explorer', path])
            else:
                subprocess.Popen(['explorer', '/select,', path])
        except Exception:
            try:
                os.startfile(path)
            except Exception:
                pass

    def sort_tree(self, col: str):
            # Toggle sort direction
            asc = self._sort_dirs.get(col, True)
            data = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
            # For size, sort by numeric value from displayed text
            if col == "size":
                def _to_bytes(txt: str) -> float:
                    # Parse human size; fallback to 0
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
            # Update arrow on sorted column heading
            arrow = "▲" if asc else "▼"
            for c in ("name", "type", "size"):
                title = self._col_titles[c]
                if c == col:
                    self.tree.heading(c, text=f"{title} {arrow}")
                else:
                    self.tree.heading(c, text=title)
            self._sort_dirs[col] = not asc

    def _on_sort(self, col: str):
        App.sort_tree(self, col)

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
                # Open Explorer to the saved file location and select it
                import subprocess
                subprocess.Popen(["explorer", "/select,", fp])
            except Exception:
                try:
                    os.startfile(os.path.dirname(fp))
                except Exception:
                    pass
        except Exception as e:
            messagebox.showerror("Export error", str(e))


if __name__ == "__main__":
    app = App()
    app.mainloop()
