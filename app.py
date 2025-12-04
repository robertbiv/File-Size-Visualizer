import os
import threading
import queue
import math
from dataclasses import dataclass
from typing import List, Optional, Callable, Dict
import ctypes
import concurrent.futures

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font

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
    units = ["B", "KB", "MB", "GB", "TB"]
    i = 0
    v = float(n)
    while v >= 1024 and i < len(units) - 1:
        v /= 1024
        i += 1
    return f"{v:.2f} {units[i]}"

def compute_dir_size(path: str,
                     file_filter: Optional[Callable[[str, int], bool]] = None,
                     progress_cb: Optional[Callable[[str], None]] = None,
                     cancel_cb: Optional[Callable[[], bool]] = None) -> int:
    total = 0
    try:
        with os.scandir(path) as it:
            for entry in it:
                if cancel_cb and cancel_cb():
                    return total
                try:
                    if entry.is_symlink():
                        continue
                    if entry.is_dir(follow_symlinks=False):
                        total += compute_dir_size(entry.path, file_filter, progress_cb, cancel_cb)
                    elif entry.is_file(follow_symlinks=False):
                        size = entry.stat(follow_symlinks=False).st_size
                        if file_filter is None or file_filter(entry.path, size):
                            total += size
                except (PermissionError, OSError):
                    pass
    except (PermissionError, OSError):
        pass
    if progress_cb: 
        try: progress_cb(path)
        except: pass
    return total

def list_items_parallel(folder: str,
                        file_filter: Optional[Callable[[str, int], bool]] = None,
                        progress_cb: Optional[Callable[[str], None]] = None,
                        cancel_cb: Optional[Callable[[], bool]] = None) -> List[ItemSize]:
    items: List[ItemSize] = []
    dirs_to_scan = []
    
    try:
        with os.scandir(folder) as it:
            for entry in it:
                if entry.is_symlink():
                    continue
                if entry.is_dir(follow_symlinks=False):
                    dirs_to_scan.append(entry)
                elif entry.is_file(follow_symlinks=False):
                    try:
                        size = entry.stat(follow_symlinks=False).st_size
                        if file_filter is None or file_filter(entry.path, size):
                            items.append(ItemSize(label=entry.name, path=entry.path, size=size, is_dir=False))
                    except:
                        pass
    except Exception:
        return []

    with concurrent.futures.ThreadPoolExecutor(max_workers=16) as executor:
        future_to_entry = {}
        for entry in dirs_to_scan:
            if cancel_cb and cancel_cb():
                break
            future = executor.submit(compute_dir_size, entry.path, file_filter, progress_cb, cancel_cb)
            future_to_entry[future] = entry

        for future in concurrent.futures.as_completed(future_to_entry):
            if cancel_cb and cancel_cb():
                break
            entry = future_to_entry[future]
            try:
                size = future.result()
                if file_filter is None or size > 0: 
                    items.append(ItemSize(label=entry.name, path=entry.path, size=size, is_dir=True))
            except Exception:
                items.append(ItemSize(label=entry.name, path=entry.path, size=0, is_dir=True))

    return items

class App(tk.Tk):
    def __init__(self):
        # DPI Support
        try: ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except: pass
        try:
            hdc = ctypes.windll.user32.GetDC(0)
            self._system_dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
            ctypes.windll.user32.ReleaseDC(0, hdc)
        except: self._system_dpi = 100
        
        super().__init__()
        try: self.tk.call('tk', 'scaling', self._system_dpi / 72.0)
        except: pass
            
        self.title("File Size Visualizer - gh: robertbiv")
        self.geometry("1100x700")

        self.selected_folder: Optional[str] = None
        self.scan_thread: Optional[threading.Thread] = None
        self.scan_queue: queue.Queue = queue.Queue()
        self._cancel_flag = False
        
        self._root_items: List[ItemSize] = [] 
        self._iid_to_path: Dict[str, str] = {}
        self._loaded_iids = set()
        self._pie_stack = []  # Stack to track pie chart states when drilling down
        self._current_pie_items = []
        self._pie_stack = []  # Stack to track pie chart states when drilling down
        self._current_pie_items = []
        
        # --- FONT SETUP ---
        self.default_font_name = "Segoe UI"
        self.font_size_var = tk.IntVar(value=11) # Default bigger size

        self._build_ui()
        self.apply_font_size() # Apply initial defaults

    def _build_ui(self):
        # --- TOP CONTROLS ---
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=10, pady=10)

        # Folder selection
        self.folder_var = tk.StringVar()
        ttk.Label(top, text="Folder:").pack(side=tk.LEFT)
        ttk.Entry(top, textvariable=self.folder_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(top, text="Browse", command=self.browse_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(top, text="Scan", command=self.start_root_scan).pack(side=tk.LEFT, padx=2)
        
        # Font Control
        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        ttk.Label(top, text="Font Size:").pack(side=tk.LEFT)
        fs_spin = ttk.Spinbox(top, from_=8, to=24, textvariable=self.font_size_var, width=3, command=self.apply_font_size)
        fs_spin.pack(side=tk.LEFT, padx=2)
        # Bind Return key to spinbox to update on typing
        fs_spin.bind('<Return>', lambda e: self.apply_font_size())
        
        ttk.Button(top, text="Export CSV", command=self.export_csv).pack(side=tk.LEFT, padx=10)

        # --- FILTER CONTROLS ---
        controls = ttk.Frame(self)
        controls.pack(fill=tk.X, padx=10)
        
        ttk.Label(controls, text="Min size:").pack(side=tk.LEFT)
        self.min_size_var = tk.StringVar(value="0")
        ttk.Entry(controls, textvariable=self.min_size_var, width=10).pack(side=tk.LEFT, padx=5)
        self.size_unit_var = tk.StringVar(value="MB")
        ttk.Combobox(controls, textvariable=self.size_unit_var, values=["B", "KB", "MB", "GB"], width=5, state="readonly").pack(side=tk.LEFT)
        
        self.apply_filter_subfolders = tk.BooleanVar(value=True)
        ttk.Checkbutton(controls, text="Include subfolders", variable=self.apply_filter_subfolders).pack(side=tk.LEFT, padx=10)
        
        self.status_var = tk.StringVar(value="Select a folder to begin.")
        ttk.Label(self, textvariable=self.status_var, anchor="w").pack(fill=tk.X, padx=10, pady=5)

        # Progress
        self.prog_frame = ttk.Frame(self)
        self.progress = ttk.Progressbar(self.prog_frame, mode="indeterminate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cancel_btn = ttk.Button(self.prog_frame, text="Cancel", command=self.cancel_scan)
        self.cancel_btn.pack(side=tk.LEFT, padx=6)
        # Initially hidden, will show during scan
        # Don't pack it yet - will be shown when scanning starts

        # --- MAIN PANE ---
        main = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        main.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        self._paned = main

        left = ttk.Frame(main, width=450)
        left.pack_propagate(False)
        main.add(left, weight=1)
        
        columns = ("type", "size")
        self.tree = ttk.Treeview(left, columns=columns)
        
        self.tree.heading("#0", text="Name", command=lambda: self.sort_tree_col("#0"))
        self.tree.heading("type", text="Type", command=lambda: self.sort_tree_col("type"))
        self.tree.heading("size", text="Size", command=lambda: self.sort_tree_col("size"))
        
        self.tree.column("#0", width=300)
        self.tree.column("type", width=80, anchor="center")
        self.tree.column("size", width=100, anchor="e")
        
        sb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.tree.bind('<<TreeviewOpen>>', self.on_tree_open)
        self.tree.bind('<<TreeviewOpen>>', self.on_tree_open)
        self.tree.bind('<<TreeviewOpen>>', self.on_tree_open)
        self.tree.bind('<<TreeviewClose>>', self.on_tree_close)
        self.tree.bind('<Double-1>', self._on_double_click)
        self.tree.bind('<Motion>', self._on_tree_hover)
        self._build_context_menu()

        right = ttk.Frame(main)
        main.add(right, weight=2)
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)

        self.figure = Figure(figsize=(5, 4), dpi=self._system_dpi, frameon=False, facecolor='none')
        self.ax = self.figure.add_subplot(111, frame_on=False, facecolor='none')
        self.canvas = FigureCanvasTkAgg(self.figure, master=right)
        self.ax.set_axis_off()
        try:
            self.ax.patch.set_visible(False)
            self.figure.patch.set_visible(False)
            self.canvas.get_tk_widget().configure(bd=0, highlightthickness=0)
        except: pass
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.after(200, lambda: self._paned.sashpos(0, 450) if self._paned.winfo_ismapped() else None)

    def apply_font_size(self):
        """Updates the application font size globally."""
        try:
            size = self.font_size_var.get()
        except:
            size = 11
        
        # 1. Update Global Tkinter Style
        style = ttk.Style()
        # Use '.' to apply to everything (Labels, Buttons, etc.)
        style.configure(".", font=(self.default_font_name, size))
        style.configure("Treeview.Heading", font=(self.default_font_name, size, "bold"))
        style.configure("Treeview", font=(self.default_font_name, size))
        
        # 2. IMPORTANT: Increase Row Height for Treeview manually
        # Tkinter Treeviews don't auto-resize rows for larger fonts
        row_h = int(size * 2.2) 
        style.configure("Treeview", rowheight=row_h)

        # 3. Redraw pie chart if it exists (to update Matplotlib text size)
        if self._root_items:
            self._draw_pie(self._root_items)

    def browse_folder(self):
        f = filedialog.askdirectory()
        if f:
            self.folder_var.set(f)
            self.start_root_scan()

    def parse_min_size(self) -> int:
        try: val = float(self.min_size_var.get().strip())
        except: val = 0.0
        unit = self.size_unit_var.get()
        mult = {"KB":1024, "MB":1024**2, "GB":1024**3}.get(unit, 1)
        return int(val * mult)

    def start_root_scan(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Invalid folder.")
            return
        
        self.tree.delete(*self.tree.get_children())
        self._iid_to_path.clear()
        self._loaded_iids.clear()
        self.ax.clear()
        self.canvas.draw()
        
        self.status_var.set("Scanning root level...")
        self.prog_frame.pack(fill=tk.X, padx=10, before=self._paned)
        self.progress.start(10)
        self._cancel_flag = False
        
        self.scan_thread = threading.Thread(target=self._scan_thread_func, 
                                            args=(folder, "", True), daemon=True)
        self.scan_thread.start()
        self.after(100, self._poll_queue)

    def on_tree_open(self, event):
        sel = self.tree.selection()
        if not sel: return
        iid = sel[0]
        
        # Close all other open folders at the same level to avoid confusion
        parent = self.tree.parent(iid)
        siblings = self.tree.get_children(parent)
        for sibling in siblings:
            if sibling != iid:
                # Check if this sibling is open
                if self.tree.item(sibling, 'open'):
                    self.tree.item(sibling, open=False)
        
        # Check if this is a folder being expanded
        children = self.tree.get_children(iid)
        if len(children) == 1 and self.tree.item(children[0], "text") == "dummy":
            # Need to load children first
            if iid not in self._loaded_iids:
                self.tree.delete(children[0])
                path = self._iid_to_path.get(iid)
                if path:
                    self.status_var.set(f"Expanding: {os.path.basename(path)}...")
                    t = threading.Thread(target=self._scan_thread_func, 
                                         args=(path, iid, False), daemon=True)
                    t.start()
                    self.after(100, self._poll_queue)
        else:
            # Children already loaded, redraw pie for this folder
            self._redraw_pie_for_folder(iid)

    def on_tree_close(self, event):
        """Restore previous pie chart when folder is collapsed"""
        if self._pie_stack:
            # Pop the current state
            self._pie_stack.pop()
            if self._pie_stack:
                # Restore the previous level
                previous_items = self._pie_stack[-1]
                self._current_pie_items = previous_items
                self._draw_pie(previous_items)
            else:
                # Back to root level
                self._current_pie_items = self._root_items
                self._draw_pie(self._root_items)

    def _redraw_pie_for_folder(self, iid):
        """Redraw pie chart showing only the contents of the expanded folder"""
        # Get all children of this folder
        children = self.tree.get_children(iid)
        if not children:
            return
        
        # Build ItemSize list from children
        folder_items = []
        for child_iid in children:
            name = self.tree.item(child_iid, "text")
            values = self.tree.item(child_iid, "values")
            if values:
                type_str = values[0]
                size_str = values[1]
                path = self._iid_to_path.get(child_iid, "")
                
                # Parse size back from human readable format
                try:
                    size = self._parse_human_size(size_str)
                    is_dir = (type_str == "Folder")
                    folder_items.append(ItemSize(label=name, path=path, size=size, is_dir=is_dir))
                except:
                    pass
        
        if folder_items:
            # Save current state before drilling down
            current = self._current_pie_items if self._current_pie_items else self._root_items
            self._pie_stack.append(current)
            self._current_pie_items = folder_items
            self._draw_pie(folder_items)

    def _parse_human_size(self, size_str: str) -> int:
        """Parse human-readable size back to bytes"""
        try:
            parts = size_str.split()
            if len(parts) == 2:
                num = float(parts[0])
                unit = parts[1]
                multipliers = {"B": 1, "KB": 1024, "KiB": 1024, "MB": 1024**2, "MiB": 1024**2, 
                              "GB": 1024**3, "GiB": 1024**3, "TB": 1024**4, "TiB": 1024**4}
                return int(num * multipliers.get(unit, 1))
        except:
            pass
        return 0

    def _scan_thread_func(self, folder, parent_iid, is_root):
        try:
            min_size = self.parse_min_size()
            def _prog(p):
                if is_root: self.scan_queue.put(("progress", p))
            def _canc():
                return self._cancel_flag

            items = list_items_parallel(folder, progress_cb=_prog, cancel_cb=_canc)
            items = [it for it in items if it.size >= min_size]
            items.sort(key=lambda x: x.size, reverse=True)
            self.scan_queue.put(("done", (parent_iid, items, is_root)))
        except Exception as e:
            self.scan_queue.put(("error", str(e)))

    def _poll_queue(self):
        try:
            while True:
                msg = self.scan_queue.get_nowait()
                kind, payload = msg
                
                if kind == "error":
                    self.status_var.set(f"Error: {payload}")
                    self._stop_prog()
                    return
                elif kind == "progress":
                    short = (payload[-40:]) if len(payload)>40 else payload
                    self.status_var.set(f"Scanning: ...{short}")
                elif kind == "done":
                    parent_iid, items, is_root = payload
                    self._populate_tree(parent_iid, items)
                    
                    if is_root:
                        self._root_items = items
                        self._current_pie_items = items
                        self._draw_pie(items)
                        self.status_var.set(f"Done. Found {len(items)} items.")
                        self._stop_prog()
                    else:
                        self._loaded_iids.add(parent_iid)
                        # After loading folder children, redraw pie for that folder
                        self._redraw_pie_for_folder(parent_iid)
                        self.status_var.set("Ready.")
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _stop_prog(self):
        self.progress.stop()
        self.prog_frame.pack_forget()

    def _populate_tree(self, parent_iid, items: List[ItemSize]):
        for it in items:
            oid = self.tree.insert(parent_iid, tk.END, text=it.label, 
                                   values=("Folder" if it.is_dir else "File", human_size(it.size)),
                                   open=False)
            self._iid_to_path[oid] = it.path
            if it.is_dir:
                self.tree.insert(oid, tk.END, text="dummy")

    def cancel_scan(self):
        self._cancel_flag = True

    def _draw_pie(self, items):
        self.ax.clear()
        try: self.ax.set_position([0,0,1,1])
        except: pass
        
        if not items:
            self.canvas.draw()
            return
            
        sizes = [max(0.1, i.size) for i in items]
        labels = [i.label for i in items]
        
        if len(sizes) > 12:
            pairs = sorted(zip(labels, sizes), key=lambda x:x[1], reverse=True)
            head = pairs[:11]
            tail = pairs[11:]
            other_sz = sum(s for l,s in tail)
            labels = [l for l,s in head] + ["Other"]
            sizes = [s for l,s in head] + [other_sz]

        try:
            import matplotlib.cm as cm
            import numpy as np
            hashes = np.array([abs(hash(l))%256 for l in labels], dtype=float)
            colors = cm.tab20((hashes%20)/20.0)
        except: colors=None
        
        # Pass font size to matplotlib textprops
        curr_fs = 11
        try: curr_fs = self.font_size_var.get()
        except: pass
        
        wedges, _ = self.ax.pie(sizes, labels=None, startangle=90, colors=colors, 
                                radius=0.98, wedgeprops={'ec':'white', 'lw':0.5},
                                textprops={'fontsize': curr_fs})
        
        self.ax.set_aspect('equal', adjustable='datalim')
        self.ax.autoscale(True)
        
        self._wedge_map = dict(zip(wedges, labels))
        self._lbl_to_wedge = dict(zip(labels, wedges))
        
        # Add hover handler for pie chart
        if hasattr(self, '_pie_hover_cid'):
            self.canvas.mpl_disconnect(self._pie_hover_cid)
        self._pie_hover_cid = self.canvas.mpl_connect('motion_notify_event', self._on_pie_hover)
        
        self.canvas.draw()

    def _on_pie_hover(self, event):
        """Highlight tree row when hovering over pie wedge"""
        if not hasattr(self, '_wedge_map') or event.inaxes != self.ax:
            return
        
        # Find which wedge is under the mouse
        hovered_label = None
        for wedge, label in self._wedge_map.items():
            if wedge.contains_point([event.x, event.y]):
                hovered_label = label
                break
        
        # Highlight the corresponding tree row
        if hovered_label:
            # Find the tree item with this label
            for iid in self.tree.get_children():
                if self.tree.item(iid, "text") == hovered_label:
                    self.tree.selection_set(iid)
                    self.tree.see(iid)
                    break
        else:
            # Clear selection when not hovering over any wedge
            self.tree.selection_remove(self.tree.selection())

    def _on_tree_hover(self, event):
        iid = self.tree.identify_row(event.y)
        if hasattr(self, '_lbl_to_wedge'):
            for w in self._lbl_to_wedge.values(): w.set_alpha(1.0)
            if iid:
                # Get the item text (could be nested in a folder)
                txt = self.tree.item(iid, "text")
                
                # First try direct match
                w = self._lbl_to_wedge.get(txt)
                
                # If no direct match, check if this item is under a root folder
                if not w:
                    # Walk up the tree to find the root parent
                    parent = self.tree.parent(iid)
                    while parent:
                        next_parent = self.tree.parent(parent)
                        if not next_parent:  # This is a root item
                            parent_txt = self.tree.item(parent, "text")
                            w = self._lbl_to_wedge.get(parent_txt)
                            break
                        parent = next_parent
                
                if w: w.set_alpha(0.6)
            self.canvas.draw_idle()

    def _on_double_click(self, event):
        sel = self.tree.selection()
        if not sel: return
        path = self._iid_to_path.get(sel[0])
        if path:
            try:
                import subprocess
                subprocess.run(['explorer', '/select,', os.path.normpath(path)])
            except: pass

    def _build_context_menu(self):
        m = tk.Menu(self, tearoff=0)
        m.add_command(label="Open", command=self._ctx_open)
        m.add_command(label="Open in File Explorer", command=self._ctx_open_explorer)
        m.add_command(label="Copy Path", command=self._ctx_copy)
        self._ctx_menu = m
        self.tree.bind('<Button-3>', self._show_ctx)

    def _show_ctx(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            self._ctx_menu.tk_popup(event.x_root, event.y_root)

    def _ctx_open(self):
        sel = self.tree.selection()
        if sel:
            path = self._iid_to_path.get(sel[0])
            if path:
                try: os.startfile(path)
                except: pass

    def _ctx_open_explorer(self):
        sel = self.tree.selection()
        if sel:
            path = self._iid_to_path.get(sel[0])
            if path:
                try:
                    import subprocess
                    subprocess.run(['explorer', '/select,', os.path.normpath(path)])
                except: pass

    def _ctx_copy(self):
        sel = self.tree.selection()
        if sel:
            path = self._iid_to_path.get(sel[0])
            if path:
                self.clipboard_clear()
                self.clipboard_append(path)

    def export_csv(self):
        if not self._root_items: return
        fp = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv")])
        if fp:
            import csv
            with open(fp, 'w', newline='', encoding='utf-8') as f:
                w = csv.writer(f)
                w.writerow(["Name", "Path", "Type", "Size"])
                for i in self._root_items:
                    w.writerow([i.label, i.path, "Folder" if i.is_dir else "File", i.size])
            messagebox.showinfo("Export", "Done")

    def sort_tree_col(self, col):
        l = [(self.tree.set(k, "size") if col=="size" else self.tree.item(k, "text"), k) 
             for k in self.tree.get_children('')]
        if col == "size":
            def parse(s):
                try:
                    num, unit = s.split()
                    return float(num) * {"B":1, "KB":1024, "MB":1024**2, "GB":1024**3}.get(unit,1)
                except: return 0
            l.sort(key=lambda x: parse(x[0]), reverse=True)
        else:
            l.sort(reverse=False)
        for idx, (_, k) in enumerate(l):
            self.tree.move(k, '', idx)

if __name__ == "__main__":
    app = App()
    app.mainloop()