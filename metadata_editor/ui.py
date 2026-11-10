from __future__ import annotations

import datetime as dt
import queue
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from .operations import CATEGORIES, changed_values, editable_fields, selected_fields, write_selected
from .operations import create_output_copy, output_copy_path
from .engine import AUDIO_SUFFIXES, FIELDS, IMAGE_FIELDS, LEGACY_SUFFIXES, MARKDOWN_FIELDS, MARKDOWN_SUFFIXES, MEDIA_FIELDS, OFFICE_SUFFIXES, PDF_FIELDS, VIDEO_SUFFIXES, ZIP_FIELDS, ZIP_SUFFIXES, IMAGE_SUFFIXES, Metadata, clear_embedded_properties, inplace_write_block_reason, load, make_writable, set_filesystem_times, verify_embedded_clear, write_audio, write_image, write_legacy, write_markdown, write_office, write_pdf, write_video, write_zip
from .history import (FileChangedError, append_history, delete_history_records,
                      history_path, read_history, restore_recorded_metadata)
from .updates import (DownloadCancelled, Release, download_installer, is_newer,
                      latest_release, read_state, should_auto_check, write_state)
from .version import APP_VERSION

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:  # Drag and drop is optional for source runs without installed requirements.
    DND_FILES = None
    TkinterDnD = None

SUPPORTED_SUFFIXES = OFFICE_SUFFIXES | LEGACY_SUFFIXES | MARKDOWN_SUFFIXES | IMAGE_SUFFIXES | AUDIO_SUFFIXES | VIDEO_SUFFIXES | ZIP_SUFFIXES | {".pdf"}
_TkBase = TkinterDnD.Tk if TkinterDnD else tk.Tk


class Editor(_TkBase):
    def __init__(self):
        super().__init__()
        self.title("文件源数据编辑")
        self._apply_icon(self)
        self.after_idle(self._fit_initial_window)
        self.meta: Metadata | None = None
        self.vars: dict[str, tk.StringVar] = {}
        self.history_var = tk.BooleanVar(value=True)
        self.output_mode = tk.StringVar(value="original")
        self.output_dir = tk.StringVar()
        self.output_template = tk.StringVar(value="{name}.已处理")
        self.path_var = tk.StringVar()
        self._update_queue: queue.Queue = queue.Queue()
        self._update_busy = False
        self._download_cancel = threading.Event()
        self._style()
        self._layout()
        self.protocol("WM_DELETE_WINDOW", self._close_app)
        self.after(2000, lambda: self.check_updates(manual=False))

    def _apply_icon(self, window):
        root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
        icon = root / "assets" / "app-icon.png"
        if icon.exists():
            self._icon_image = tk.PhotoImage(file=icon)
            window.iconphoto(True, self._icon_image)

    def _fit_initial_window(self):
        """Size the first window to the actual DPI-aware monitor dimensions."""
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = min(1240, max(760, int(screen_width * 0.90)))
        height = min(780, max(500, int(screen_height * 0.82)))
        self.minsize(min(900, max(680, int(screen_width * 0.62))), min(600, max(440, int(screen_height * 0.58))))
        self.geometry(f"{width}x{height}")

    def _style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        bg, surface, border, ink, muted, accent = "#f7f7f7", "#ffffff", "#cccccc", "#111111", "#666666", "#111111"
        disabled = "#999999"
        self.configure(background=bg)
        style.configure("TFrame", background=bg)
        style.configure("Surface.TFrame", background=surface)
        style.configure("TLabel", background=bg, foreground=ink, font=("Microsoft YaHei UI", 9))
        style.configure("Surface.TLabel", background=surface, foreground=ink)
        style.map("TLabel", background=[("disabled", bg)], foreground=[("disabled", disabled)])
        style.map("Surface.TLabel", background=[("disabled", surface)], foreground=[("disabled", disabled)])
        style.configure("Card.TLabelframe", background=surface, bordercolor=border, relief="solid", borderwidth=1)
        style.configure("Card.TLabelframe.Label", background=bg, foreground=ink, font=("Microsoft YaHei UI", 9, "bold"))
        style.configure("Title.TLabel", background=bg, foreground=ink, font=("Microsoft YaHei UI", 13, "bold"))
        style.configure("Muted.TLabel", background=bg, foreground=muted, font=("Microsoft YaHei UI", 9))
        style.configure("Section.TLabel", background=bg, foreground=ink, font=("Microsoft YaHei UI", 9, "bold"))
        style.configure("TButton", font=("Microsoft YaHei UI", 9), padding=(10, 5), background=surface, foreground=ink, bordercolor=border)
        style.map("TButton",
                  background=[("disabled", surface), ("active", "#eeeeee")],
                  foreground=[("disabled", disabled)],
                  bordercolor=[("disabled", border), ("focus", accent)])
        style.configure("Accent.TButton", font=("Microsoft YaHei UI", 9, "bold"), foreground="white", background=accent, padding=(14, 6), bordercolor=accent)
        style.map("Accent.TButton", background=[("disabled", accent), ("active", "#333333")],
                  foreground=[("disabled", disabled)])
        style.configure("Danger.TButton", font=("Microsoft YaHei UI", 9), foreground=ink, background=surface, bordercolor=border, padding=(10, 5))
        style.map("Danger.TButton", background=[("disabled", surface), ("active", "#eeeeee")],
                  foreground=[("disabled", disabled)])
        style.configure("TEntry", padding=(8, 6), foreground=ink, fieldbackground=surface,
                        bordercolor=border, lightcolor=border, darkcolor=border)
        style.map("TEntry", foreground=[("disabled", disabled)],
                  fieldbackground=[("disabled", surface)], background=[("disabled", surface)],
                  bordercolor=[("disabled", border)])
        style.configure("TScrollbar", background="#eeeeee", troughcolor=bg,
                        bordercolor=border, arrowcolor=ink)
        style.map("TScrollbar", background=[("disabled", "#eeeeee")],
                  troughcolor=[("disabled", bg)], arrowcolor=[("disabled", disabled)])
        style.configure("TRadiobutton", background=bg, foreground=ink, font=("Microsoft YaHei UI", 9))
        style.configure("TCheckbutton", background=bg, foreground=ink, font=("Microsoft YaHei UI", 9))
        style.configure("Surface.TRadiobutton", background=surface, foreground=ink)
        style.configure("Surface.TCheckbutton", background=surface, foreground=ink)
        for name, base in (("TRadiobutton", bg), ("TCheckbutton", bg),
                           ("Surface.TRadiobutton", surface), ("Surface.TCheckbutton", surface)):
            style.map(name, background=[("disabled", base)], foreground=[("disabled", disabled)])
        style.configure("Treeview", background=surface, fieldbackground=surface, foreground=ink, rowheight=27, borderwidth=0, font=("Microsoft YaHei UI", 9))
        style.configure("Treeview.Heading", background="#eeeeee", foreground=ink, font=("Microsoft YaHei UI", 9), relief="flat", padding=(6, 6))
        style.map("Treeview", background=[("selected", ink)], foreground=[("selected", surface)])

    def _output_settings(self, parent, history_var, browse_command):
        panel = ttk.Labelframe(parent, text=" 输出设置 ", style="Card.TLabelframe", padding=(10, 8))
        panel.pack(fill="x", pady=(0, 7))

        choices = ttk.Frame(panel, style="Surface.TFrame")
        choices.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 6))
        ttk.Label(choices, text="保存方式", width=12, style="Surface.TLabel").pack(side="left")
        ttk.Radiobutton(choices, text="修改原文件", variable=self.output_mode, value="original", style="Surface.TRadiobutton").pack(side="left", padx=(0, 18))
        ttk.Radiobutton(choices, text="另存副本", variable=self.output_mode, value="copy", style="Surface.TRadiobutton").pack(side="left")
        history_options = ttk.Frame(choices, style="Surface.TFrame")
        history_options.pack(side="right")
        ttk.Label(history_options, text="记录修改历史", style="Surface.TLabel").pack(side="left", padx=(0, 10))
        ttk.Radiobutton(history_options, text="是", variable=history_var, value=True, style="Surface.TRadiobutton").pack(side="left", padx=(0, 10))
        ttk.Radiobutton(history_options, text="否", variable=history_var, value=False, style="Surface.TRadiobutton").pack(side="left")

        name_label = ttk.Label(panel, text="副本命名", width=12, style="Surface.TLabel")
        name_label.grid(row=1, column=0, sticky="w")
        naming = ttk.Frame(panel, style="Surface.TFrame")
        naming.grid(row=1, column=1, columnspan=2, sticky="w")
        name_entry = ttk.Entry(naming, textvariable=self.output_template, width=28)
        name_entry.pack(side="left")
        name_hint = ttk.Label(naming, text="{name} 代表原文件名 · 同名自动加序号", style="Surface.TLabel")
        name_hint.pack(side="left", padx=(12, 0))

        folder_label = ttk.Label(panel, text="副本文件夹", width=12, style="Surface.TLabel")
        folder_label.grid(row=2, column=0, sticky="w", pady=(6, 0))
        folder_entry = ttk.Entry(panel, textvariable=self.output_dir)
        folder_entry.grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=(6, 0))
        browse_button = ttk.Button(panel, text="浏览", command=browse_command)
        browse_button.grid(row=2, column=2, sticky="ew", pady=(6, 0))

        def update_copy_controls(*_):
            copy_mode = self.output_mode.get() == "copy"
            state = "normal" if copy_mode else "disabled"
            for widget in (name_entry, folder_entry, browse_button, name_label, name_hint, folder_label):
                widget.configure(state=state)

        trace_id = self.output_mode.trace_add("write", update_copy_controls)
        def remove_trace(event):
            if event.widget is panel:
                self.output_mode.trace_remove("write", trace_id)
        panel.bind("<Destroy>", remove_trace, add="+")
        update_copy_controls()
        panel.columnconfigure(1, weight=1)
        return panel

    def _layout(self):
        outer = ttk.Frame(self, padding=(14, 10, 14, 10))
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="文件源数据编辑", style="Title.TLabel").pack(anchor="w", pady=(0, 9))
        chooser = ttk.Frame(outer)
        chooser.pack(fill="x", pady=(0, 8))
        ttk.Label(chooser, text="文件", width=6).pack(side="left")
        ttk.Entry(chooser, textvariable=self.path_var).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(chooser, text="选择文件", command=self.choose).pack(side="left")
        ttk.Button(chooser, text="批量处理", command=self.batch_clear).pack(side="left", padx=(8, 0))
        self._output_settings(outer, self.history_var, self._choose_output_dir)
        self.drop_hint = ttk.Label(outer, text="单个文件可拖入此窗口；多个文件请使用批量处理", style="Muted.TLabel")
        self.drop_hint.pack(anchor="w", pady=(0, 5))
        body = ttk.Panedwindow(outer, orient="horizontal")
        left = ttk.Labelframe(body, text=" 文件信息 ", style="Card.TLabelframe", padding=7)
        right = ttk.Labelframe(body, text=" 编辑属性 ", style="Card.TLabelframe", padding=9)
        body.add(left, weight=1); body.add(right, weight=2)
        self.tree = ttk.Treeview(left, columns=("value",), show="tree headings", selectmode="browse")
        self.tree.heading("#0", text="属性"); self.tree.heading("value", text="值")
        self.tree.column("#0", width=145, stretch=False); self.tree.column("value", width=225)
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview); self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")
        self.form_canvas = tk.Canvas(right, highlightthickness=0, background="#ffffff")
        form_scroll = ttk.Scrollbar(right, orient="vertical", command=self.form_canvas.yview)
        self.form = ttk.Frame(self.form_canvas, style="Surface.TFrame", padding=(2, 2, 10, 2))
        self.form.bind("<Configure>", lambda _e: self.form_canvas.configure(scrollregion=self.form_canvas.bbox("all")))
        self.form_canvas.create_window((0, 0), window=self.form, anchor="nw", tags="form")
        self.form_canvas.bind("<Configure>", lambda e: self.form_canvas.itemconfigure("form", width=e.width))
        self.form_canvas.configure(yscrollcommand=form_scroll.set)
        self.form_canvas.pack(side="left", fill="both", expand=True); form_scroll.pack(side="right", fill="y")
        # Pack the non-shrinking footer first.  The body is the only region
        # allowed to lose height while the user makes the window smaller.
        footer = ttk.Frame(outer); footer.pack(fill="x", side="bottom", pady=(10, 0))
        body.pack(fill="both", expand=True)
        status_row = ttk.Frame(footer); status_row.pack(fill="x", pady=(0, 7))
        self.status = ttk.Label(status_row, text="请选择一个文件开始", style="Muted.TLabel"); self.status.pack(side="left")
        action_row = ttk.Frame(footer); action_row.pack(fill="x")
        ttk.Button(action_row, text="查看修改历史", command=self.show_history).pack(side="left")
        self.update_button = ttk.Button(action_row, text="检查更新", command=self.check_updates)
        self.update_button.pack(side="left", padx=(8, 0))
        ttk.Button(action_row, text="清除嵌入属性", style="Danger.TButton", command=self.clear).pack(side="right")
        ttk.Button(action_row, text="保存修改", style="Accent.TButton", command=self.save).pack(side="right", padx=(0, 8))
        # Let the wheel work naturally even when the pointer is above an entry,
        # label, or nested form widget rather than only above the scrollbar.
        self.bind_all("<MouseWheel>", self._mousewheel, add="+")
        self._register_drop_targets(self)

    def _close_app(self):
        self._download_cancel.set()
        self.destroy()

    def check_updates(self, manual: bool = True):
        if self._update_busy or (not manual and not should_auto_check()):
            return
        self._update_busy = True
        self.update_button.configure(state="disabled")

        def worker():
            try:
                self._update_queue.put(("checked", manual, latest_release()))
            except Exception as exc:
                self._update_queue.put(("check_error", manual, str(exc)))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll_updates)

    def _poll_updates(self):
        while True:
            try:
                kind, *data = self._update_queue.get_nowait()
            except queue.Empty:
                break
            if kind == "checked":
                manual, release = data
                self._update_busy = False
                self.update_button.configure(state="normal")
                state = read_state()
                state["last_checked_utc"] = dt.datetime.now(dt.timezone.utc).isoformat()
                try:
                    write_state(state)
                except OSError:
                    pass
                if not is_newer(release):
                    if manual:
                        messagebox.showinfo("检查更新", f"当前版本 {APP_VERSION} 已是最新版本。", parent=self)
                    continue
                if not manual and state.get("dismissed_tag") == release.tag:
                    continue
                if messagebox.askyesno(
                    "发现新版本",
                    f"当前版本：{APP_VERSION}\n最新版本：{release.tag}\n\n"
                    "下载并运行新版安装程序吗？下载完成后将校验 SHA-256。",
                    parent=self,
                ):
                    self._start_update_download(release)
                else:
                    state["dismissed_tag"] = release.tag
                    try:
                        write_state(state)
                    except OSError:
                        pass
            elif kind == "check_error":
                manual, error = data
                self._update_busy = False
                self.update_button.configure(state="normal")
                if manual:
                    messagebox.showwarning("检查更新失败", f"无法连接 GitHub 或读取发布信息：\n{error}", parent=self)
            elif kind == "progress":
                received, total = data
                if self._progress_window.winfo_exists():
                    self._progress_bar.configure(value=received / total * 100)
                    self._progress_label.configure(text=f"已下载 {received / 1048576:.0f} / {total / 1048576:.0f} MB")
            elif kind == "downloaded":
                self._finish_update_download()
                installer = data[0]
                try:
                    subprocess.Popen([str(installer)], cwd=str(installer.parent), close_fds=True)
                except OSError as exc:
                    messagebox.showerror("无法启动安装程序", str(exc), parent=self)
                else:
                    self._close_app()
                    return
            elif kind in {"download_error", "download_cancelled"}:
                self._finish_update_download()
                if kind == "download_error":
                    messagebox.showerror("更新失败", data[0], parent=self)
        if self._update_busy:
            self.after(100, self._poll_updates)

    def _start_update_download(self, release: Release):
        self._update_busy = True
        self._download_cancel.clear()
        self.update_button.configure(state="disabled")
        dialog = tk.Toplevel(self)
        dialog.title(f"下载 {release.tag}")
        dialog.transient(self)
        dialog.resizable(False, False)
        self._apply_icon(dialog)
        self._progress_window = dialog
        panel = ttk.Frame(dialog, padding=20)
        panel.pack(fill="both", expand=True)
        ttk.Label(panel, text="正在下载并校验新版安装程序…").pack(anchor="w")
        self._progress_bar = ttk.Progressbar(panel, maximum=100, length=380)
        self._progress_bar.pack(fill="x", pady=(14, 8))
        self._progress_label = ttk.Label(panel, text="准备下载", style="Muted.TLabel")
        self._progress_label.pack(anchor="w")
        cancel_button = ttk.Button(panel, text="取消", command=self._download_cancel.set)
        cancel_button.pack(anchor="e", pady=(12, 0))
        dialog.protocol("WM_DELETE_WINDOW", self._download_cancel.set)

        def worker():
            try:
                installer = download_installer(
                    release,
                    progress=lambda received, total: self._update_queue.put(("progress", received, total)),
                    cancelled=self._download_cancel.is_set,
                )
                self._update_queue.put(("downloaded", installer))
            except DownloadCancelled:
                self._update_queue.put(("download_cancelled",))
            except Exception as exc:
                self._update_queue.put(("download_error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_update_download(self):
        self._update_busy = False
        self.update_button.configure(state="normal")
        if self._progress_window.winfo_exists():
            self._progress_window.destroy()

    def _register_drop_targets(self, widget):
        if not DND_FILES:
            return
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind("<<Drop>>", self._drop_main)
        for child in widget.winfo_children():
            self._register_drop_targets(child)

    @staticmethod
    def _inside(widget, ancestor):
        while widget is not None:
            if widget == ancestor:
                return True
            parent = widget.winfo_parent()
            widget = widget.nametowidget(parent) if parent else None
        return False

    def _mousewheel(self, event):
        under_pointer = self.winfo_containing(event.x_root, event.y_root)
        steps = -max(1, abs(event.delta) // 120) if event.delta > 0 else max(1, abs(event.delta) // 120)
        if under_pointer and self._inside(under_pointer, self.tree):
            self.tree.yview_scroll(steps, "units")
            return "break"
        if under_pointer and self._inside(under_pointer, self.form_canvas):
            self.form_canvas.yview_scroll(steps, "units")
            return "break"

    def choose(self):
        path = filedialog.askopenfilename(filetypes=[("支持的文件", " ".join(f"*{suffix}" for suffix in sorted(SUPPORTED_SUFFIXES))), ("所有文件", "*.*")])
        if path: self.open(path)

    def _choose_output_dir(self):
        directory = filedialog.askdirectory(parent=self, title="选择副本输出文件夹")
        if directory:
            self.output_dir.set(directory)

    def _drop_main(self, event):
        paths = [Path(path) for path in self.tk.splitlist(event.data)]
        paths = [path for path in paths if path.is_file()]
        if len(paths) == 1:
            if paths[0].suffix.lower() in SUPPORTED_SUFFIXES:
                self.open(str(paths[0]))
            else:
                messagebox.showwarning("不支持的文件", f"暂不支持此文件类型：{paths[0].suffix or '无扩展名'}", parent=self)
        elif len(paths) > 1:
            messagebox.showinfo("请使用批量清除", "主窗口一次打开一个文件。请点击“批量清除”，再将多个文件拖入批量窗口。", parent=self)
        return "break"

    def open(self, path: str):
        try:
            self.meta = load(path); self.path_var.set(path); self._render(); self.status.configure(text=f"已读取：{self.meta.kind} 文档")
        except Exception as e: messagebox.showerror("无法打开", str(e))

    def _render(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for child in self.form.winfo_children(): child.destroy()
        assert self.meta
        fs = self.tree.insert("", "end", text="文件系统时间", open=True)
        self.tree.insert(fs, "end", text="创建时间", values=(self.meta.values.get("filesystem_created", ""),))
        self.tree.insert(fs, "end", text="修改时间", values=(self.meta.values.get("filesystem_modified", ""),))
        props = self.tree.insert("", "end", text=f"{self.meta.kind} 属性", open=True)
        fields = PDF_FIELDS if self.meta.kind == "PDF" else MARKDOWN_FIELDS if self.meta.kind == "Markdown" else IMAGE_FIELDS if self.meta.kind == "Image" else ZIP_FIELDS if self.meta.kind == "Archive" else MEDIA_FIELDS if self.meta.kind in {"Audio", "Video"} else FIELDS
        for key, label in fields: self.tree.insert(props, "end", text=label, values=(self.meta.values.get(key, ""),))
        if self.meta.kind in {"Office", "Legacy Office", "PDF", "Markdown"}:
            custom_node = self.tree.insert("", "end", text="自定义属性", open=True)
            for key, value in self.meta.custom.items(): self.tree.insert(custom_node, "end", text=key, values=(value,))
        self.vars = {}
        self._section("文件系统时间", [("filesystem_created", "创建时间"), ("filesystem_modified", "修改时间")], readonly=False)
        fields = PDF_FIELDS if self.meta.kind == "PDF" else MARKDOWN_FIELDS if self.meta.kind == "Markdown" else IMAGE_FIELDS if self.meta.kind == "Image" else ZIP_FIELDS if self.meta.kind == "Archive" else MEDIA_FIELDS if self.meta.kind in {"Audio", "Video"} else FIELDS
        groups = [("来源与描述", [x for x in fields if x[0] not in {"created", "modified", "lastPrinted", "totalTime"}]), ("文档元数据时间", [x for x in fields if x[0] in {"created", "modified", "lastPrinted", "totalTime"}])]
        for title, fields in groups: self._section(title, fields)
        if self.meta.kind in {"Office", "Legacy Office", "PDF", "Markdown"}:
            box = ttk.Labelframe(self.form, text=" 自定义属性（每行：名称 = 值） ", style="Card.TLabelframe", padding=10); box.pack(fill="x", pady=(0, 12))
            self.custom_text = tk.Text(box, height=5, font=("Consolas", 10), relief="solid", borderwidth=1,
                                       background="#ffffff", foreground="#111111", insertbackground="#111111",
                                       selectbackground="#111111", selectforeground="#ffffff")
            self.custom_text.pack(fill="x"); self.custom_text.insert("1.0", "\n".join(f"{k} = {v}" for k, v in self.meta.custom.items()))
            if self.meta.kind == "PDF":
                ttk.Label(self.form, text="PDF 自定义 Info 键以 / 开头；加密或已签名 PDF 不可修改。", foreground="#666666", background="#ffffff").pack(anchor="w", pady=8)
            elif self.meta.kind == "Markdown":
                ttk.Label(self.form, text="Markdown 属性保存在 YAML Front Matter；多行和嵌套字段会保留，但暂不支持逐项修改。", foreground="#666666", background="#ffffff").pack(anchor="w", pady=8)
        elif self.meta.kind == "Image":
            ttk.Label(self.form, text="图片编辑 EXIF 常用字段；PNG 不含 EXIF 时不会强行写入。", foreground="#666666", background="#ffffff").pack(anchor="w", pady=8)
        elif self.meta.kind == "Archive":
            ttk.Label(self.form, text="压缩包支持 ZIP 注释；清除属性会清空注释，不会删除文件内容。", foreground="#666666", background="#ffffff").pack(anchor="w", pady=8)
        elif self.meta.kind == "Audio":
            ttk.Label(self.form, text="音频标签使用 Mutagen 写入。", foreground="#666666", background="#ffffff").pack(anchor="w", pady=8)
        elif self.meta.kind == "Video":
            ttk.Label(self.form, text="视频标签使用 FFmpeg 无损封装复制写入，处理时间取决于文件大小。", foreground="#666666", background="#ffffff").pack(anchor="w", pady=8)
        self._register_drop_targets(self.form)

    def _section(self, title, fields, readonly=False):
        box = ttk.Labelframe(self.form, text=f" {title} ", style="Card.TLabelframe", padding=10); box.pack(fill="x", pady=(0, 12))
        for row, (key, label) in enumerate(fields):
            ttk.Label(box, text=label, width=18, style="Surface.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
            var = tk.StringVar(value=self.meta.values.get(key, "")); self.vars[key] = var
            ttk.Entry(box, textvariable=var).grid(row=row, column=1, sticky="ew", pady=5)
        box.columnconfigure(1, weight=1)

    def _custom(self):
        result = {}
        for line in self.custom_text.get("1.0", "end").splitlines():
            if not line.strip(): continue
            if "=" not in line: raise ValueError("自定义属性每行必须使用“名称 = 值”格式。")
            key, value = line.split("=", 1); key = key.strip()
            if not key: raise ValueError("自定义属性名称不能为空。")
            result[key] = value.strip()
        return result

    def _system_times(self):
        assert self.meta
        set_filesystem_times(self.meta.path, self.vars["filesystem_created"].get().strip(), self.vars["filesystem_modified"].get().strip())

    def _copy_target(self, source: Metadata, action: str, label: str, reason: str) -> Metadata | None:
        if not messagebox.askyesno(
            "无法直接修改",
            f"无法直接{action}原文件：{reason}。\n\n是否另存为一个副本并继续？原文件不会被修改。",
            parent=self,
        ):
            return None
        target = filedialog.asksaveasfilename(
            title="另存为可修改副本",
            initialdir=str(Path(self.output_dir.get()) if self.output_dir.get().strip() else source.path.parent),
            initialfile=output_copy_path(source.path, self.output_dir.get(), self.output_template.get()).name,
            defaultextension=source.path.suffix,
            filetypes=[("同类型文件", f"*{source.path.suffix}"), ("所有文件", "*.*")],
        )
        if not target:
            return None
        target_path = Path(target)
        if target_path.resolve() == source.path.resolve():
            messagebox.showwarning("请选择新文件名", "副本不能与原文件同名。", parent=self)
            return None
        shutil.copy2(source.path, target_path)
        make_writable(target_path)
        return load(str(target_path))

    def _target_for_write(self, action: str, label: str) -> tuple[Metadata, bool] | None:
        """Return the original metadata or a user-chosen writable copy target."""
        assert self.meta
        if self.output_mode.get() == "copy":
            target = create_output_copy(self.meta.path, self.output_dir.get(), self.output_template.get())
            return load(str(target)), True
        reason = inplace_write_block_reason(self.meta.path)
        if not reason:
            return self.meta, False
        copy = self._copy_target(self.meta, action, label, reason)
        return (copy, True) if copy else None

    def show_history(self):
        window = tk.Toplevel(self)
        window.title("修改历史")
        self._apply_icon(window)
        window.geometry("1060x690")
        window.minsize(720, 480)
        window.transient(self)
        frame = ttk.Frame(window, padding=14)
        frame.pack(fill="both", expand=True)

        heading = ttk.Frame(frame)
        heading.pack(fill="x", pady=(0, 10))
        ttk.Label(heading, text="修改历史", style="Title.TLabel").pack(side="left")
        ttk.Button(heading, text="刷新", command=lambda: refresh()).pack(side="right")
        restore_button = ttk.Button(heading, text="恢复清除前属性", command=lambda: restore_selected())
        restore_button.pack(side="right", padx=(0, 8))
        restore_button.configure(state="disabled")
        delete_selected_button = ttk.Button(heading, text="删除选中记录", command=lambda: delete_selected())
        delete_selected_button.pack(side="right", padx=(0, 8))
        delete_selected_button.configure(state="disabled")
        delete_all_button = ttk.Button(heading, text="清空全部记录", command=lambda: delete_all())
        delete_all_button.pack(side="right", padx=(0, 8))
        delete_all_button.configure(state="disabled")

        filter_row = ttk.Frame(frame)
        filter_row.pack(fill="x", pady=(0, 8))
        ttk.Label(filter_row, text="查找文件", width=10).pack(side="left")
        query = tk.StringVar()
        ttk.Entry(filter_row, textvariable=query).pack(side="left", fill="x", expand=True)
        count_label = ttk.Label(filter_row, style="Muted.TLabel")
        count_label.pack(side="right", padx=(12, 0))

        table_frame = ttk.Frame(frame)
        table_frame.pack(fill="both", expand=True)
        columns = ("time", "operation", "file", "type", "verification", "hash")
        table = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse", height=10)
        for key, label, width in (("time", "时间", 155), ("operation", "操作", 88),
                                  ("file", "结果文件", 260), ("type", "类型", 100),
                                  ("verification", "验证", 100), ("hash", "SHA-256", 170)):
            table.heading(key, text=label)
            table.column(key, width=width, stretch=key in {"file", "hash"})
        table_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=table.yview)
        table.configure(yscrollcommand=table_scroll.set)
        table.pack(side="left", fill="both", expand=True)
        table_scroll.pack(side="right", fill="y")

        detail = ttk.Labelframe(frame, text=" 记录详情 ", style="Card.TLabelframe", padding=9)
        detail.pack(fill="both", expand=True, pady=(10, 0))
        source_label = ttk.Label(detail, text="选择上方记录查看详情", style="Surface.TLabel", wraplength=950)
        source_label.pack(anchor="w")
        result_label = ttk.Label(detail, style="Surface.TLabel", wraplength=950)
        result_label.pack(anchor="w", pady=(4, 0))
        window.bind("<Configure>", lambda event: (
            source_label.configure(wraplength=max(400, event.width - 80)),
            result_label.configure(wraplength=max(400, event.width - 80)),
        ) if event.widget is window else None)
        hash_row = ttk.Frame(detail, style="Surface.TFrame")
        hash_row.pack(fill="x", pady=(4, 8))
        hash_label = ttk.Label(hash_row, style="Surface.TLabel")
        hash_label.pack(side="left", fill="x", expand=True)

        def copy_hash():
            selection = table.selection()
            if not selection:
                return
            index = int(selection[0])
            if index >= len(visible):
                return
            value = visible[index].get("sha256", "")
            if value:
                window.clipboard_clear()
                window.clipboard_append(value)

        ttk.Button(hash_row, text="复制哈希", command=copy_hash).pack(side="right")
        fields_frame = ttk.Frame(detail, style="Surface.TFrame")
        fields_frame.pack(fill="both", expand=True)
        fields = ttk.Treeview(fields_frame, columns=("value",), show="tree headings", height=8)
        fields.heading("#0", text="属性")
        fields.heading("value", text="记录值")
        fields.column("#0", width=210, stretch=False)
        fields.column("value", width=650)
        fields_scroll = ttk.Scrollbar(fields_frame, orient="vertical", command=fields.yview)
        fields.configure(yscrollcommand=fields_scroll.set)
        fields.pack(side="left", fill="both", expand=True)
        fields_scroll.pack(side="right", fill="y")
        note = ttk.Label(frame, text=f"记录文件：{history_path()}  ·  仅清除记录可恢复；校验目标文件的清除后哈希", style="Muted.TLabel")
        note.pack(anchor="w", pady=(8, 0))

        records: list[dict] = []
        visible: list[dict] = []
        skipped = 0
        operation_names = {"save": "保存", "clear": "清除", "batch_save": "批量修改", "batch_clear": "批量清除", "restore": "恢复属性"}
        verification_names = {"read_back": "已读回", "passed": "通过", "residual": "有残留", "unreadable": "无法读回", "time_failed": "时间未写入"}
        field_labels = {key: label for group in (FIELDS, PDF_FIELDS, MARKDOWN_FIELDS, IMAGE_FIELDS, ZIP_FIELDS, MEDIA_FIELDS)
                        for key, label in group}
        field_labels.update({"filesystem_created": "文件创建时间", "filesystem_modified": "文件修改时间"})

        def local_time(value):
            try:
                return dt.datetime.fromisoformat(value).astimezone().strftime("%Y-%m-%d %H:%M:%S")
            except (TypeError, ValueError):
                return str(value or "")

        def show_selected(_event=None):
            for item in fields.get_children():
                fields.delete(item)
            selection = table.selection()
            if not selection:
                restore_button.configure(state="disabled")
                delete_selected_button.configure(state="disabled")
                source_label.configure(text="选择上方记录查看详情")
                result_label.configure(text="")
                hash_label.configure(text="")
                return
            index = int(selection[0])
            if index >= len(visible):
                return
            record = visible[index]
            delete_selected_button.configure(state="normal")
            can_restore = (record.get("operation") in {"clear", "batch_clear"}
                           and isinstance(record.get("metadata_before"), dict) and bool(record.get("sha256")))
            restore_button.configure(state="normal" if can_restore else "disabled")
            source_label.configure(text=f"原文件：{record.get('source_path', record.get('source_name', ''))}")
            result_label.configure(text=f"结果文件：{record.get('result_path', '')}")
            hash_label.configure(text=f"SHA-256：{record.get('sha256', '')}")
            for snapshot_key, phase in (("metadata_before", "操作前（可恢复）"), ("metadata_after", "操作后")):
                snapshot = record.get(snapshot_key)
                if not isinstance(snapshot, dict):
                    fields.insert("", "end", text=phase, values=("旧记录未保存" if snapshot_key == "metadata_before" else "无法读回",))
                    continue
                phase_node = fields.insert("", "end", text=phase, open=snapshot_key == "metadata_before")
                for group, title in (("values", "预设属性"), ("custom", "自定义属性"), ("readonly", "只读属性")):
                    values = snapshot.get(group, {})
                    if not isinstance(values, dict) or not values:
                        continue
                    parent = fields.insert(phase_node, "end", text=title, open=snapshot_key == "metadata_before")
                    for key, value in values.items():
                        label = field_labels.get(key, str(key)) if group == "values" else str(key)
                        fields.insert(parent, "end", text=label, values=(str(value) if value is not None else "",))

        def apply_filter(*_):
            visible[:] = [record for record in reversed(records)
                          if query.get().strip().casefold() in
                          (str(record.get("source_name", "")) + " " + str(record.get("result_name", "")) + " " +
                           str(record.get("source_path", "")) + " " + str(record.get("result_path", ""))).casefold()]
            for item in table.get_children():
                table.delete(item)
            for index, record in enumerate(visible):
                table.insert("", "end", iid=str(index), values=(
                    local_time(record.get("time_utc")),
                    operation_names.get(record.get("operation"), record.get("operation", "")),
                    record.get("result_name", ""), record.get("file_type") or "—",
                    verification_names.get(record.get("verification"), record.get("verification", "")),
                    (str(record.get("sha256", ""))[:16] + "…") if record.get("sha256") else "",
                ))
            count_label.configure(text=f"{len(visible)} 条" + (f"，跳过损坏记录 {skipped} 条" if skipped else ""))
            if visible:
                table.selection_set("0")
                show_selected()
            else:
                show_selected()

        def refresh():
            nonlocal skipped
            try:
                loaded, skipped = read_history()
            except OSError as exc:
                messagebox.showerror("无法读取修改历史", str(exc), parent=window)
                return
            records[:] = loaded
            delete_all_button.configure(state="normal" if loaded or skipped else "disabled")
            apply_filter()

        def delete_selected():
            selection = table.selection()
            if not selection or int(selection[0]) >= len(visible):
                return
            record = visible[int(selection[0])]
            index = next((i for i, item in enumerate(records) if item is record), None)
            if index is None:
                refresh()
                return
            if not messagebox.askyesno(
                "删除选中记录",
                f"确定删除这条修改历史吗？\n{record.get('result_name') or record.get('result_path', '')}"
                "\n\n只删除历史记录，不会改动文件。删除后无法再通过这条记录恢复清除前属性。",
                parent=window,
            ):
                return
            try:
                delete_history_records({index: record})
            except (OSError, ValueError) as exc:
                messagebox.showerror("删除失败", str(exc), parent=window)
            refresh()

        def delete_all():
            if not records and not skipped:
                return
            if not messagebox.askyesno(
                "清空全部记录",
                f"确定清空全部 {len(records)} 条修改历史吗？"
                "\n\n只删除历史记录，不会改动文件。清空后无法再通过这些记录恢复清除前属性。",
                parent=window,
            ):
                return
            try:
                delete_history_records()
            except (OSError, ValueError) as exc:
                messagebox.showerror("清空失败", str(exc), parent=window)
            refresh()

        def restore_selected():
            selection = table.selection()
            if not selection or int(selection[0]) >= len(visible):
                return
            record = visible[int(selection[0])]
            path = Path(str(record.get("result_path", "")))
            if not messagebox.askyesno(
                "确认恢复属性",
                f"将清除前的属性写回：\n{path}\n\n会先核对它是否仍是清除后的那份文件；哈希不符则不会写入。是否继续？",
                parent=window,
            ):
                return
            try:
                before_restore, after, changed = restore_recorded_metadata(record)
            except FileChangedError as exc:
                messagebox.showwarning("文件已变动", str(exc), parent=window)
                return
            except Exception as exc:
                messagebox.showerror("恢复失败", str(exc), parent=window)
                return
            if not changed:
                messagebox.showinfo("无需恢复", "当前文件的可编辑属性和文件时间已与这条记录一致。", parent=window)
                return
            try:
                append_history(path, after, "restore", before=before_restore, verification="passed")
            except Exception as exc:
                messagebox.showwarning("属性已恢复，记录失败", f"属性已写回，但无法追加修改历史：\n{exc}", parent=window)
            if self.meta and self.meta.path.resolve() == path.resolve():
                self.meta = after
                self._render()
            refresh()
            messagebox.showinfo("恢复完成", "已写回记录中的可编辑属性，并重新读取验证。", parent=window)

        table.bind("<<TreeviewSelect>>", show_selected)
        query.trace_add("write", apply_filter)
        refresh()

    def _record_history(self, before: Metadata, result: Metadata | Path, operation: str,
                        verification: str = "read_back", parent=None) -> None:
        if not self.history_var.get():
            return
        try:
            append_history(before.path, result, operation, before=before, verification=verification)
        except Exception as exc:
            messagebox.showwarning("修改已完成，记录失败", f"文件已修改，但无法写入修改历史：\n{exc}", parent=parent or self)

    def save(self):
        if not self.meta: return
        try:
            before = load(str(self.meta.path))
            target = self._target_for_write("保存", "已修改属性")
            if target is None: return
            meta, is_copy = target
            if not is_copy:
                meta = before
            values = {key: var.get().strip() for key, var in self.vars.items()}
            try:
                if meta.kind == "Office": write_office(meta, values, self._custom(), False)
                elif meta.kind == "Legacy Office": write_legacy(meta, values, False, self._custom())
                elif meta.kind == "Markdown": write_markdown(meta, values, custom=self._custom())
                elif meta.kind == "Image": write_image(meta, values)
                elif meta.kind == "Archive": write_zip(meta, values)
                elif meta.kind == "Audio": write_audio(meta, values)
                elif meta.kind == "Video": write_video(meta, values)
                else: write_pdf(meta, values, False, self._custom())
            except PermissionError:
                if is_copy: raise
                copy = self._copy_target(meta, "保存", "已修改属性", "Windows 拒绝替换原文件；它可能仍被其他程序占用")
                if copy is None: return
                meta, is_copy = copy, True
                if meta.kind == "Office": write_office(meta, values, self._custom(), False)
                elif meta.kind == "Legacy Office": write_legacy(meta, values, False, self._custom())
                elif meta.kind == "Markdown": write_markdown(meta, values, custom=self._custom())
                elif meta.kind == "Image": write_image(meta, values)
                elif meta.kind == "Archive": write_zip(meta, values)
                elif meta.kind == "Audio": write_audio(meta, values)
                elif meta.kind == "Video": write_video(meta, values)
                else: write_pdf(meta, values, False, self._custom())
            time_error = None
            try:
                set_filesystem_times(meta.path, values["filesystem_created"], values["filesystem_modified"])
            except Exception as exc:
                time_error = exc
            try:
                verified = load(str(meta.path))
            except Exception as verify_error:
                self._record_history(before, meta.path, "save", "unreadable")
                self.status.configure(text="保存已执行，但重新读取验证失败")
                messagebox.showwarning("保存后无法验证", f"文件已保存，但无法重新读取属性：\n{verify_error}", parent=self)
                return
            self.meta = verified
            self.path_var.set(str(meta.path))
            self._render()
            self._record_history(before, verified, "save", "time_failed" if time_error else "read_back")
            if time_error:
                self.status.configure(text="元数据已保存，但文件系统时间写入失败")
                messagebox.showwarning("部分保存成功", f"嵌入元数据已保存并记录历史，但文件系统时间写入失败：\n{time_error}", parent=self)
            else:
                self.status.configure(text="已另存为副本并保存。" if is_copy else "保存成功，已重新读取验证。")
        except Exception as e: messagebox.showerror("保存失败", str(e))

    def clear(self):
        if not self.meta: return
        try:
            before = load(str(self.meta.path))
        except Exception as exc:
            messagebox.showerror("无法读取清除前属性", str(exc), parent=self)
            return
        preview = [(before.path.name, label, before.values.get(key, ""), "清除")
                   for key, label in editable_fields(before) if before.values.get(key, "")]
        preview += [(before.path.name, f"自定义：{key}", value, "清除") for key, value in before.custom.items()]
        preview += [(before.path.name, "其他嵌入属性", before.values[key], "清除")
                    for key in ("_all_exif", "_all_tags", "_extra_image_metadata") if before.values.get(key)]
        if not self._property_table("清除前预览", preview, confirm=True): return
        try:
            target = self._target_for_write("清除属性", "已清除属性")
            if target is None: return
            meta, is_copy = target
            if not is_copy:
                meta = before
            try:
                clear_embedded_properties(meta)
            except PermissionError:
                if is_copy: raise
                copy = self._copy_target(meta, "清除属性", "已清除属性", "Windows 拒绝替换原文件；它可能仍被其他程序占用")
                if copy is None: return
                meta, is_copy = copy, True
                clear_embedded_properties(meta)
            try:
                verified, remaining = verify_embedded_clear(meta.path)
            except Exception as verify_error:
                self._record_history(before, meta.path, "clear", "unreadable")
                self.status.configure(text="清除已执行，但重新读取验证失败")
                messagebox.showwarning(
                    "清除后无法验证",
                    f"清除操作已执行，但程序无法重新读取文件确认结果：\\n{verify_error}",
                    parent=self,
                )
                return
            self.meta = verified
            self._record_history(before, verified, "clear", "passed" if not remaining else "residual")
            self.path_var.set(str(meta.path))
            self._render()
            changes = changed_values(before, verified, {key for key, _ in editable_fields(before)} | {"__custom__", "__all__"})
            self._property_table("清除后对照", [(meta.path.name, *row) for row in changes])
            if remaining:
                self.status.configure(text=f"清除验证未通过：仍检测到 {len(remaining)} 项")
                messagebox.showwarning(
                    "清除验证未通过",
                    "重新读取后仍检测到以下嵌入属性：\\n\\n"
                    + "\\n".join(remaining[:20])
                    + ("\\n……" if len(remaining) > 20 else ""),
                    parent=self,
                )
            else:
                self.status.configure(
                    text="已另存为副本并清除、验证通过。" if is_copy else "嵌入属性已清除并重新读取验证通过。"
                )
        except Exception as e: messagebox.showerror("清除失败", str(e))

    def _property_table(self, title, rows, confirm=False, parent=None):
        """Show planned or verified changes to embedded properties."""
        owner = parent or self
        window = tk.Toplevel(owner)
        window.title(title)
        window.transient(owner)
        window.geometry("900x460")
        window.grab_set()
        container = ttk.Frame(window, padding=12)
        container.pack(fill="both", expand=True)
        ttk.Label(container, text=title, style="Title.TLabel").pack(anchor="w")
        ttk.Label(container, text="仅显示程序支持的嵌入属性；Windows 文件时间不在此范围内。", style="Muted.TLabel").pack(anchor="w", pady=(2, 8))
        table = ttk.Treeview(container, columns=("file", "field", "before", "after"), show="headings")
        for key, label, width in (("file", "文件", 180), ("field", "属性", 150), ("before", "原值", 270), ("after", "处理后 / 计划", 220)):
            table.heading(key, text=label)
            table.column(key, width=width)
        for row in rows:
            table.insert("", "end", values=row)
        if not rows:
            table.insert("", "end", values=("", "无可显示属性", "", ""))
        bar = ttk.Scrollbar(container, orient="vertical", command=table.yview)
        table.configure(yscrollcommand=bar.set)
        table.pack(side="left", fill="both", expand=True)
        bar.pack(side="left", fill="y")
        actions = ttk.Frame(window, padding=(12, 0, 12, 12))
        actions.pack(fill="x")
        accepted = tk.BooleanVar(value=False)
        if confirm:
            ttk.Button(actions, text="确认执行", style="Accent.TButton", command=lambda: (accepted.set(True), window.destroy())).pack(side="right")
            ttk.Button(actions, text="取消", command=window.destroy).pack(side="right", padx=8)
        else:
            ttk.Button(actions, text="关闭", command=window.destroy).pack(side="right")
        owner.wait_window(window)
        if parent and owner.winfo_exists():
            owner.grab_set()
        return accepted.get() if confirm else None

    def batch_clear(self):
        self._show_batch_dialog([])

    def _show_batch_dialog(self, paths: list[str]):
        initial_paths = list(paths)
        dialog = tk.Toplevel(self)
        dialog.title("批量处理属性")
        self._apply_icon(dialog)
        dialog.transient(self)
        dialog.grab_set()
        dialog.minsize(760, 430)
        dialog.geometry("900x560")
        frame = ttk.Frame(dialog, padding=(16, 14, 16, 8))
        ttk.Label(frame, text="批量处理属性", style="Title.TLabel").pack(anchor="w")
        hint = ttk.Label(frame, text="将文件拖到此窗口，或点击“添加文件”。支持一次加入多个文件。", style="Muted.TLabel")
        hint.pack(anchor="w", pady=(3, 12))
        options = ttk.Labelframe(frame, text=" 操作方式 ", style="Card.TLabelframe", padding=(10, 7))
        options.pack(fill="x", pady=(0, 8))
        mode = tk.StringVar(value="all")
        for label, value in (("清除全部", "all"), ("按类别清除", "selected"), ("批量修改", "modify")):
            ttk.Radiobutton(options, text=label, variable=mode, value=value, style="Surface.TRadiobutton").pack(side="left", padx=(0, 18))
        category_bar = ttk.Labelframe(frame, text=" 选择清除类别 ", style="Card.TLabelframe", padding=(10, 7))
        categories = {name: tk.BooleanVar(value=True) for name in CATEGORIES}
        for name, variable in categories.items():
            ttk.Checkbutton(category_bar, text=name, variable=variable, style="Surface.TCheckbutton").pack(side="left", padx=(0, 10))
        edit_bar = ttk.Labelframe(frame, text=" 统一修改的属性 ", style="Card.TLabelframe", padding=(10, 7))
        edit_fields = {}
        for label, key in (("标题", "title"), ("作者", "author"), ("关键词", "keywords")):
            enabled, value = tk.BooleanVar(value=False), tk.StringVar()
            ttk.Checkbutton(edit_bar, text=label, variable=enabled, style="Surface.TCheckbutton").pack(side="left")
            ttk.Entry(edit_bar, textvariable=value, width=16).pack(side="left", padx=(0, 8))
            edit_fields[key] = (enabled, value)
        filename_title = tk.BooleanVar(value=False)
        ttk.Checkbutton(edit_bar, text="标题取文件名", variable=filename_title, style="Surface.TCheckbutton").pack(side="left")
        def show_mode_options(*_):
            category_bar.pack_forget()
            edit_bar.pack_forget()
            if mode.get() == "selected":
                category_bar.pack(fill="x", pady=(0, 8), after=options)
            elif mode.get() == "modify":
                edit_bar.pack(fill="x", pady=(0, 8), after=options)
        mode.trace_add("write", show_mode_options)
        def choose_batch_folder():
            directory = filedialog.askdirectory(parent=dialog, title="选择副本输出文件夹")
            if directory:
                self.output_dir.set(directory)
        self._output_settings(frame, self.history_var, choose_batch_folder)
        table = ttk.Treeview(frame, columns=("file", "status", "detail"), show="headings")
        table.heading("file", text="文件")
        table.heading("status", text="状态")
        table.heading("detail", text="结果 / 原因")
        table.column("file", width=420, minwidth=250)
        table.column("status", width=110, anchor="center", stretch=False)
        table.column("detail", width=320, minwidth=180)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=table.yview)
        table.configure(yscrollcommand=scrollbar.set)
        table.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        rows = {}
        paths = []

        footer = ttk.Frame(dialog, padding=(16, 7, 16, 12))
        footer.pack(fill="x", side="bottom")
        frame.pack(fill="both", expand=True)
        progress = ttk.Label(footer, text="等待添加文件", style="Muted.TLabel")
        progress.pack(side="left")
        close_button = ttk.Button(footer, text="关闭", command=dialog.destroy)
        close_button.pack(side="right")
        start_button = ttk.Button(footer, text="开始处理", style="Accent.TButton")
        start_button.pack(side="right", padx=(0, 8))
        start_button.configure(state="disabled")
        add_button = ttk.Button(footer, text="添加文件")
        add_button.pack(side="right", padx=(0, 8))

        def add_paths(incoming):
            added = 0
            for item in incoming:
                path = str(Path(item))
                file_path = Path(path)
                if not file_path.is_file() or file_path.suffix.lower() not in SUPPORTED_SUFFIXES or path in rows:
                    continue
                paths.append(path)
                rows[path] = table.insert("", "end", values=(path, "等待", ""))
                added += 1
            count = len(paths)
            hint.configure(text=f"已添加 {count} 个文件。将清除文件内嵌属性，不修改 Windows 文件时间；受保护文件会另存为可写副本。" if count else "将文件拖到此窗口，或点击“添加文件”。支持一次加入多个文件。")
            progress.configure(text=f"已添加 {count} 个文件" if count else "等待添加文件")
            start_button.configure(state="normal" if count else "disabled")
            return added

        def add_dropped(event):
            incoming = self.tk.splitlist(event.data)
            added = add_paths(incoming)
            if not added:
                messagebox.showwarning("没有可添加的文件", "请拖入本程序支持的文件，且文件需存在于本机。", parent=dialog)
            return "break"

        def add_from_picker():
            incoming = filedialog.askopenfilenames(
                parent=dialog,
                title="添加要清除属性的文件",
                filetypes=[("支持的文件", " ".join(f"*{suffix}" for suffix in sorted(SUPPORTED_SUFFIXES))), ("所有文件", "*.*")],
            )
            if incoming:
                add_paths(incoming)

        add_button.configure(command=add_from_picker)
        if DND_FILES:
            def register_batch_drop(widget):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", add_dropped)
                for child in widget.winfo_children():
                    register_batch_drop(child)

            register_batch_drop(dialog)
        add_paths(initial_paths)

        def plan(meta):
            if mode.get() == "all":
                keys = {key for key, _ in editable_fields(meta)}
                updates = {key: "" for key in keys}
                return updates, keys | {"__custom__", "__all__"}
            if mode.get() == "selected":
                chosen = {name for name, var in categories.items() if var.get()}
                keys = {key for key, _ in selected_fields(meta, chosen)}
                return {key: "" for key in keys}, keys
            requested = {key: value.get().strip() for key, (enabled, value) in edit_fields.items() if enabled.get()}
            if filename_title.get():
                requested["title"] = meta.path.stem
            available = {key for key, _ in editable_fields(meta)}
            updates = {}
            for key, value in requested.items():
                target = "creator" if key == "author" and meta.kind in {"Office", "Legacy Office"} else "tags" if key == "keywords" and meta.kind == "Markdown" else key
                if target in available:
                    updates[target] = value
            return updates, set(updates)

        def start_selected():
            if not paths:
                return
            if mode.get() == "selected" and not any(v.get() for v in categories.values()):
                messagebox.showwarning("请选择类别", "至少选择一种要清除的属性类别。", parent=dialog)
                return
            if mode.get() == "modify" and not filename_title.get() and not any(v[0].get() for v in edit_fields.values()):
                messagebox.showwarning("请选择修改项", "勾选至少一个要统一修改的属性。", parent=dialog)
                return
            snapshots = {}
            preview_rows = []
            skipped = 0
            for path in paths:
                try:
                    meta = load(path)
                    updates, keys = plan(meta)
                    if mode.get() != "all" and not updates:
                        skipped += 1
                        table.item(rows[path], values=(path, "跳过", "该格式没有所选属性"))
                        continue
                    snapshots[path] = (meta, updates, keys)
                    labels = dict(editable_fields(meta))
                    for key, value in updates.items():
                        old = meta.values.get(key, "")
                        if old != value:
                            preview_rows.append((meta.path.name, labels.get(key, key), old, value or "清除"))
                    if mode.get() == "all":
                        preview_rows.extend((meta.path.name, f"自定义：{key}", value, "清除") for key, value in meta.custom.items())
                        preview_rows.extend((meta.path.name, "其他嵌入属性", meta.values[key], "清除") for key in ("_all_exif", "_all_tags", "_extra_image_metadata") if meta.values.get(key))
                except Exception as exc:
                    table.item(rows[path], values=(path, "无法读取", str(exc)))
            if not snapshots:
                messagebox.showerror("无法开始", "没有能够读取的文件。", parent=dialog)
                return
            if not self._property_table("批量处理前预览", preview_rows, confirm=True, parent=dialog):
                return
            start_button.configure(state="disabled")
            close_button.configure(state="disabled")
            succeeded = failed = verification_failed = 0
            result_rows = []
            for index, (path, (_preview_before, updates, keys)) in enumerate(snapshots.items(), start=1):
                table.item(rows[path], values=(path, "处理中", ""))
                progress.configure(text=f"正在处理 {index} / {len(snapshots)}")
                dialog.update_idletasks()
                try:
                    before = load(path)
                    meta = before
                    reason = inplace_write_block_reason(meta.path)
                    is_copy = self.output_mode.get() == "copy" or bool(reason)
                    if is_copy:
                        copy_path = create_output_copy(meta.path, self.output_dir.get(), self.output_template.get())
                        meta = load(str(copy_path))
                    def perform(item):
                        if mode.get() == "all":
                            clear_embedded_properties(item)
                        elif updates:
                            write_selected(item, updates)
                    try:
                        perform(meta)
                    except PermissionError:
                        if is_copy:
                            raise
                        copy_path = create_output_copy(meta.path, self.output_dir.get(), self.output_template.get())
                        meta = load(str(copy_path))
                        perform(meta)
                        is_copy = True
                    try:
                        after = load(str(meta.path))
                    except Exception:
                        self._record_history(before, meta.path, "batch_clear" if mode.get() != "modify" else "batch_save", "unreadable", dialog)
                        raise
                    changes = changed_values(before, after, keys)
                    result_rows.extend((meta.path.name, *change) for change in changes)
                    if mode.get() == "all":
                        try:
                            _, residual = verify_embedded_clear(meta.path)
                        except Exception:
                            self._record_history(before, after, "batch_clear", "unreadable", dialog)
                            raise
                    else:
                        residual = [dict(editable_fields(before)).get(key, key) for key in updates if after.values.get(key, "") != updates[key]]
                    self._record_history(before, after, "batch_clear" if mode.get() != "modify" else "batch_save", "passed" if not residual else "residual", dialog)
                    if residual:
                        verification_failed += 1
                        table.item(rows[path], values=(path, "验证未通过", "未达到预期：" + "、".join(residual[:8])))
                    else:
                        succeeded += 1
                        detail = (f"另存副本：{meta.path.name}；" if is_copy else "") + f"已核对 {len(keys)} 项"
                        table.item(rows[path], values=(path, "成功", detail))
                except Exception as exc:
                    failed += 1
                    table.item(rows[path], values=(path, "失败", str(exc)))
                dialog.update_idletasks()
            progress.configure(text=f"完成：成功 {succeeded}，未通过 {verification_failed}，失败 {failed}，跳过 {skipped}")
            close_button.configure(state="normal")
            self._property_table("批量处理后对照", result_rows, parent=dialog)

        start_button.configure(command=start_selected)


def launch():
    Editor().mainloop()
