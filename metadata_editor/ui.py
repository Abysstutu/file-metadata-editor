from __future__ import annotations

import shutil
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from .operations import CATEGORIES, changed_values, editable_fields, selected_fields, write_selected
from .operations import create_output_copy, output_copy_path, restore_from_backup
from .engine import AUDIO_SUFFIXES, FIELDS, IMAGE_FIELDS, LEGACY_SUFFIXES, MARKDOWN_FIELDS, MARKDOWN_SUFFIXES, MEDIA_FIELDS, OFFICE_SUFFIXES, PDF_FIELDS, VIDEO_SUFFIXES, ZIP_FIELDS, ZIP_SUFFIXES, IMAGE_SUFFIXES, Metadata, backup, clear_embedded_properties, inplace_write_block_reason, load, make_writable, set_filesystem_times, suggested_copy_path, verify_embedded_clear, write_audio, write_image, write_legacy, write_markdown, write_office, write_pdf, write_video, write_zip

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
        self.after_idle(self._fit_initial_window)
        self.meta: Metadata | None = None
        self.vars: dict[str, tk.StringVar] = {}
        self.backup_var = tk.BooleanVar(value=True)
        self.output_mode = tk.StringVar(value="original")
        self.output_dir = tk.StringVar()
        self.output_template = tk.StringVar(value="{name}.已处理")
        self.undo_history: list[tuple[Path, Path]] = []
        self.path_var = tk.StringVar()
        self._style()
        self._layout()

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
        style.configure("TFrame", background="#f7f8fa")
        style.configure("Card.TLabelframe", background="#ffffff", bordercolor="#dfe3e8", relief="solid")
        style.configure("Card.TLabelframe.Label", background="#ffffff", foreground="#243447", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Title.TLabel", background="#f7f8fa", foreground="#172b4d", font=("Microsoft YaHei UI", 15, "bold"))
        style.configure("Muted.TLabel", background="#f7f8fa", foreground="#667085", font=("Microsoft YaHei UI", 9))
        style.configure("Accent.TButton", font=("Microsoft YaHei UI", 10, "bold"), foreground="white", background="#2563eb", padding=(16, 7))
        style.map("Accent.TButton", background=[("active", "#1d4ed8")])

    def _layout(self):
        outer = ttk.Frame(self, padding=18)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="文件源数据编辑", style="Title.TLabel").pack(anchor="w")
        ttk.Label(outer, text="本地编辑 · Office 与 PDF · 保存前可自动备份", style="Muted.TLabel").pack(anchor="w", pady=(2, 14))
        chooser = ttk.Frame(outer)
        chooser.pack(fill="x", pady=(0, 12))
        ttk.Label(chooser, text="文件", width=6).pack(side="left")
        ttk.Entry(chooser, textvariable=self.path_var).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(chooser, text="选择文件", command=self.choose).pack(side="left")
        ttk.Button(chooser, text="批量处理", command=self.batch_clear).pack(side="left", padx=(8, 0))
        output_bar = ttk.Frame(outer)
        output_bar.pack(fill="x", pady=(0, 8))
        ttk.Label(output_bar, text="输出：").pack(side="left")
        ttk.Radiobutton(output_bar, text="修改原文件", variable=self.output_mode, value="original").pack(side="left")
        ttk.Radiobutton(output_bar, text="另存副本", variable=self.output_mode, value="copy").pack(side="left", padx=(8, 12))
        ttk.Label(output_bar, text="文件夹").pack(side="left")
        ttk.Entry(output_bar, textvariable=self.output_dir, width=22).pack(side="left", padx=(4, 4))
        ttk.Button(output_bar, text="选择", command=self._choose_output_dir).pack(side="left")
        ttk.Label(output_bar, text="命名").pack(side="left", padx=(12, 0))
        ttk.Entry(output_bar, textvariable=self.output_template, width=20).pack(side="left", padx=(4, 0))
        ttk.Label(output_bar, text="{name}=原文件名", style="Muted.TLabel").pack(side="left", padx=(8, 0))
        self.drop_hint = ttk.Label(outer, text="也可以将单个文件拖到此窗口打开；多个文件请拖入“批量处理”窗口。", style="Muted.TLabel")
        self.drop_hint.pack(anchor="w", pady=(0, 8))
        body = ttk.Panedwindow(outer, orient="horizontal")
        left = ttk.Labelframe(body, text=" 文件信息 ", style="Card.TLabelframe", padding=8)
        right = ttk.Labelframe(body, text=" 编辑属性 ", style="Card.TLabelframe", padding=14)
        body.add(left, weight=1); body.add(right, weight=2)
        self.tree = ttk.Treeview(left, columns=("value",), show="tree headings", selectmode="browse")
        self.tree.heading("#0", text="属性"); self.tree.heading("value", text="值")
        self.tree.column("#0", width=145, stretch=False); self.tree.column("value", width=225)
        scroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview); self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")
        self.form_canvas = tk.Canvas(right, highlightthickness=0, background="#ffffff")
        form_scroll = ttk.Scrollbar(right, orient="vertical", command=self.form_canvas.yview)
        self.form = ttk.Frame(self.form_canvas, padding=(2, 2, 10, 2))
        self.form.bind("<Configure>", lambda _e: self.form_canvas.configure(scrollregion=self.form_canvas.bbox("all")))
        self.form_canvas.create_window((0, 0), window=self.form, anchor="nw", tags="form")
        self.form_canvas.bind("<Configure>", lambda e: self.form_canvas.itemconfigure("form", width=e.width))
        self.form_canvas.configure(yscrollcommand=form_scroll.set)
        self.form_canvas.pack(side="left", fill="both", expand=True); form_scroll.pack(side="right", fill="y")
        # Pack the non-shrinking footer first.  The body is the only region
        # allowed to lose height while the user makes the window smaller.
        footer = ttk.Frame(outer); footer.pack(fill="x", side="bottom", pady=(12, 0))
        body.pack(fill="both", expand=True)
        ttk.Label(footer, text="保存前备份：").pack(side="left")
        ttk.Radiobutton(footer, text="是", variable=self.backup_var, value=True).pack(side="left", padx=(2, 8))
        ttk.Radiobutton(footer, text="否", variable=self.backup_var, value=False).pack(side="left")
        self.status = ttk.Label(footer, text="请选择一个文件开始", style="Muted.TLabel"); self.status.pack(side="left", padx=16)
        ttk.Button(footer, text="清除嵌入属性", command=self.clear).pack(side="right")
        ttk.Button(footer, text="保存修改", style="Accent.TButton", command=self.save).pack(side="right", padx=(0, 8))
        ttk.Button(footer, text="从备份恢复", command=self.restore_current).pack(side="right", padx=(0, 8))
        ttk.Button(footer, text="撤销上次操作", command=self.undo_last).pack(side="right", padx=(0, 8))
        # Let the wheel work naturally even when the pointer is above an entry,
        # label, or nested form widget rather than only above the scrollbar.
        self.bind_all("<MouseWheel>", self._mousewheel, add="+")
        self._register_drop_targets(self)

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
        if self.meta.kind == "Office":
            custom_node = self.tree.insert("", "end", text="自定义属性", open=True)
            for key, value in self.meta.custom.items(): self.tree.insert(custom_node, "end", text=key, values=(value,))
        self.vars = {}
        self._section("文件系统时间", [("filesystem_created", "创建时间"), ("filesystem_modified", "修改时间")], readonly=False)
        fields = PDF_FIELDS if self.meta.kind == "PDF" else MARKDOWN_FIELDS if self.meta.kind == "Markdown" else IMAGE_FIELDS if self.meta.kind == "Image" else ZIP_FIELDS if self.meta.kind == "Archive" else MEDIA_FIELDS if self.meta.kind in {"Audio", "Video"} else FIELDS
        groups = [("来源与描述", [x for x in fields if x[0] not in {"created", "modified", "lastPrinted", "totalTime"}]), ("文档元数据时间", [x for x in fields if x[0] in {"created", "modified", "lastPrinted", "totalTime"}])]
        for title, fields in groups: self._section(title, fields)
        if self.meta.kind == "Office":
            box = ttk.Labelframe(self.form, text=" 自定义属性（每行：名称 = 值） ", style="Card.TLabelframe", padding=10); box.pack(fill="x", pady=(0, 12))
            self.custom_text = tk.Text(box, height=5, font=("Consolas", 10), relief="solid", borderwidth=1)
            self.custom_text.pack(fill="x"); self.custom_text.insert("1.0", "\n".join(f"{k} = {v}" for k, v in self.meta.custom.items()))
        elif self.meta.kind == "Legacy Office":
            ttk.Label(self.form, text="旧版 Office 使用 Apache POI 属性流兼容层。创建内容时间等字段暂不支持单独修改。", foreground="#a16207", background="#ffffff").pack(anchor="w", pady=8)
        elif self.meta.kind == "Markdown":
            ttk.Label(self.form, text="Markdown 元数据使用 YAML Front Matter 保存；没有时保存会自动创建，清除会移除整个 Front Matter。", foreground="#a16207", background="#ffffff").pack(anchor="w", pady=8)
        elif self.meta.kind == "Image":
            ttk.Label(self.form, text="图片编辑 EXIF 常用字段；PNG 不含 EXIF 时不会强行写入。", foreground="#a16207", background="#ffffff").pack(anchor="w", pady=8)
        elif self.meta.kind == "Archive":
            ttk.Label(self.form, text="压缩包支持 ZIP 注释；清除属性会清空注释，不会删除文件内容。", foreground="#a16207", background="#ffffff").pack(anchor="w", pady=8)
        elif self.meta.kind == "Audio":
            ttk.Label(self.form, text="音频标签使用 Mutagen 写入。", foreground="#a16207", background="#ffffff").pack(anchor="w", pady=8)
        elif self.meta.kind == "Video":
            ttk.Label(self.form, text="视频标签使用 FFmpeg 无损封装复制写入，处理时间取决于文件大小。", foreground="#a16207", background="#ffffff").pack(anchor="w", pady=8)
        else:
            ttk.Label(self.form, text="PDF 会重写元数据保存。加密或已签名 PDF 将拒绝修改。", foreground="#a16207", background="#ffffff").pack(anchor="w", pady=8)
        self._register_drop_targets(self.form)

    def _section(self, title, fields, readonly=False):
        box = ttk.Labelframe(self.form, text=f" {title} ", style="Card.TLabelframe", padding=10); box.pack(fill="x", pady=(0, 12))
        for row, (key, label) in enumerate(fields):
            ttk.Label(box, text=label, width=18).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
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

    def undo_last(self):
        if not self.undo_history:
            messagebox.showinfo("没有可撤销操作", "本次运行中没有可用的操作备份；请使用“从备份恢复”选择历史备份。", parent=self)
            return
        targets = list(self.undo_history)
        if not messagebox.askyesno("确认撤销", f"将从备份恢复 {len(targets)} 个文件。恢复前会为当前状态另建备份，是否继续？", parent=self):
            return
        errors = []
        restored = []
        for target, saved in targets:
            try:
                if not saved.is_file():
                    raise FileNotFoundError(f"备份不存在：{saved}")
                backup(target)
                restore_from_backup(target, saved)
                restored.append(target)
            except Exception as exc:
                errors.append(f"{target.name}: {exc}")
        self.undo_history = [(target, saved) for target, saved in targets if target not in restored]
        if self.meta and self.meta.path in restored:
            self.open(str(self.meta.path))
        self.status.configure(text=f"已恢复 {len(restored)} 个文件；失败 {len(errors)} 个")
        if errors:
            messagebox.showwarning("部分文件恢复失败", "\n".join(errors[:12]), parent=self)

    def restore_current(self):
        if not self.meta:
            messagebox.showinfo("请先选择文件", "先打开要恢复的文件。", parent=self)
            return
        target = self.meta.path
        saved = filedialog.askopenfilename(
            parent=self, title="选择此文件的 .metadata-backup 备份",
            initialdir=str(target.parent),
            filetypes=[("备份文件", f"*.metadata-backup{target.suffix}"), ("所有文件", "*.*")],
        )
        if not saved:
            return
        source = Path(saved)
        if not source.name.startswith(f"{target.stem}.") or ".metadata-backup" not in source.stem:
            messagebox.showerror("备份不匹配", "请选择当前文件的备份，文件名应以原文件名开头。", parent=self)
            return
        if not messagebox.askyesno("确认恢复", f"用备份恢复：\n{target}\n\n恢复前会为当前状态另建备份。是否继续？", parent=self):
            return
        try:
            backup(target)
            restore_from_backup(target, source)
            self.open(str(target))
            self.undo_history = []
            self.status.configure(text="已从备份恢复，当前状态另有备份可供再次恢复。")
        except Exception as exc:
            messagebox.showerror("恢复失败", str(exc), parent=self)

    def save(self):
        if not self.meta: return
        try:
            target = self._target_for_write("保存", "已修改属性")
            if target is None: return
            meta, is_copy = target
            saved_backup = backup(meta.path) if self.backup_var.get() else None
            values = {key: var.get().strip() for key, var in self.vars.items()}
            try:
                if meta.kind == "Office": write_office(meta, values, self._custom(), False)
                elif meta.kind == "Legacy Office": write_legacy(meta, values, False)
                elif meta.kind == "Markdown": write_markdown(meta, values)
                elif meta.kind == "Image": write_image(meta, values)
                elif meta.kind == "Archive": write_zip(meta, values)
                elif meta.kind == "Audio": write_audio(meta, values)
                elif meta.kind == "Video": write_video(meta, values)
                else: write_pdf(meta, values, False)
            except PermissionError:
                if is_copy: raise
                copy = self._copy_target(meta, "保存", "已修改属性", "Windows 拒绝替换原文件；它可能仍被其他程序占用")
                if copy is None: return
                meta, is_copy = copy, True
                saved_backup = backup(meta.path) if self.backup_var.get() else None
                if meta.kind == "Office": write_office(meta, values, self._custom(), False)
                elif meta.kind == "Legacy Office": write_legacy(meta, values, False)
                elif meta.kind == "Markdown": write_markdown(meta, values)
                elif meta.kind == "Image": write_image(meta, values)
                elif meta.kind == "Archive": write_zip(meta, values)
                elif meta.kind == "Audio": write_audio(meta, values)
                elif meta.kind == "Video": write_video(meta, values)
                else: write_pdf(meta, values, False)
            set_filesystem_times(meta.path, values["filesystem_created"], values["filesystem_modified"])
            self.undo_history = [(meta.path, saved_backup)] if saved_backup else []
            self.open(str(meta.path))
            self.status.configure(text="已另存为副本并保存。" if is_copy else "保存成功，已重新读取验证。")
        except Exception as e: messagebox.showerror("保存失败", str(e))

    def clear(self):
        if not self.meta: return
        before = self.meta
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
            saved_backup = backup(meta.path) if self.backup_var.get() else None
            try:
                clear_embedded_properties(meta)
            except PermissionError:
                if is_copy: raise
                copy = self._copy_target(meta, "清除属性", "已清除属性", "Windows 拒绝替换原文件；它可能仍被其他程序占用")
                if copy is None: return
                meta, is_copy = copy, True
                saved_backup = backup(meta.path) if self.backup_var.get() else None
                clear_embedded_properties(meta)
            self.undo_history = [(meta.path, saved_backup)] if saved_backup else []
            try:
                verified, remaining = verify_embedded_clear(meta.path)
            except Exception as verify_error:
                self.status.configure(text="清除已执行，但重新读取验证失败")
                messagebox.showwarning(
                    "清除后无法验证",
                    f"清除操作已执行，但程序无法重新读取文件确认结果：\\n{verify_error}",
                    parent=self,
                )
                return
            self.meta = verified
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
        dialog.transient(self)
        dialog.grab_set()
        dialog.minsize(760, 430)
        dialog.geometry("900x560")
        frame = ttk.Frame(dialog, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="批量处理属性", style="Title.TLabel").pack(anchor="w")
        hint = ttk.Label(frame, text="将文件拖到此窗口，或点击“添加文件”。支持一次加入多个文件。", style="Muted.TLabel")
        hint.pack(anchor="w", pady=(3, 12))
        options = ttk.Frame(frame)
        options.pack(fill="x", pady=(0, 8))
        mode = tk.StringVar(value="all")
        for label, value in (("清除全部", "all"), ("按类别清除", "selected"), ("批量修改", "modify")):
            ttk.Radiobutton(options, text=label, variable=mode, value=value).pack(side="left", padx=(0, 14))
        category_bar = ttk.Frame(frame)
        category_bar.pack(fill="x", pady=(0, 8))
        ttk.Label(category_bar, text="清除类别：").pack(side="left")
        categories = {name: tk.BooleanVar(value=True) for name in CATEGORIES}
        for name, variable in categories.items():
            ttk.Checkbutton(category_bar, text=name, variable=variable).pack(side="left", padx=(0, 10))
        edit_bar = ttk.Frame(frame)
        edit_bar.pack(fill="x", pady=(0, 8))
        edit_fields = {}
        for label, key in (("标题", "title"), ("作者", "author"), ("关键词", "keywords")):
            enabled, value = tk.BooleanVar(value=False), tk.StringVar()
            ttk.Checkbutton(edit_bar, text=label, variable=enabled).pack(side="left")
            ttk.Entry(edit_bar, textvariable=value, width=16).pack(side="left", padx=(0, 8))
            edit_fields[key] = (enabled, value)
        filename_title = tk.BooleanVar(value=False)
        ttk.Checkbutton(edit_bar, text="标题取文件名", variable=filename_title).pack(side="left")
        output_bar = ttk.Frame(frame)
        output_bar.pack(fill="x", pady=(0, 8))
        ttk.Label(output_bar, text="输出：").pack(side="left")
        ttk.Radiobutton(output_bar, text="修改原文件", variable=self.output_mode, value="original").pack(side="left")
        ttk.Radiobutton(output_bar, text="另存副本", variable=self.output_mode, value="copy").pack(side="left", padx=(8, 12))
        ttk.Label(output_bar, text="文件夹").pack(side="left")
        ttk.Entry(output_bar, textvariable=self.output_dir, width=20).pack(side="left", padx=4)
        def choose_batch_folder():
            directory = filedialog.askdirectory(parent=dialog, title="选择副本输出文件夹")
            if directory:
                self.output_dir.set(directory)
        ttk.Button(output_bar, text="选择", command=choose_batch_folder).pack(side="left")
        ttk.Label(output_bar, text="命名").pack(side="left", padx=(12, 0))
        ttk.Entry(output_bar, textvariable=self.output_template, width=18).pack(side="left", padx=4)
        ttk.Label(output_bar, text="{name}=原文件名", style="Muted.TLabel").pack(side="left")
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

        footer = ttk.Frame(dialog, padding=(16, 0, 16, 16))
        footer.pack(fill="x")
        batch_backup = tk.BooleanVar(value=self.backup_var.get())
        ttk.Label(footer, text="保存前备份：").pack(side="left")
        ttk.Radiobutton(footer, text="是", variable=batch_backup, value=True).pack(side="left", padx=(2, 8))
        ttk.Radiobutton(footer, text="否", variable=batch_backup, value=False).pack(side="left")
        progress = ttk.Label(footer, text="等待添加文件", style="Muted.TLabel")
        progress.pack(side="left", padx=16)
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

        def start():
            if not paths:
                return
            if not messagebox.askyesno("确认批量清除", f"确认清除这 {len(paths)} 个文件的所有内嵌属性吗？\n此操作将无法撤销。", parent=dialog):
                return
            start_button.configure(state="disabled")
            close_button.configure(state="disabled")
            succeeded = failed = verification_failed = 0
            for index, path in enumerate(paths, start=1):
                table.item(rows[path], values=(path, "处理中", ""))
                progress.configure(text=f"正在处理 {index} / {len(paths)}")
                dialog.update_idletasks()
                try:
                    meta = load(path)
                    reason = inplace_write_block_reason(meta.path)
                    is_copy = bool(reason)
                    if is_copy:
                        copy_path = suggested_copy_path(meta.path)
                        shutil.copy2(meta.path, copy_path)
                        make_writable(copy_path)
                        meta = load(str(copy_path))
                    if batch_backup.get() and not is_copy:
                        backup(meta.path)
                    try:
                        clear_embedded_properties(meta)
                    except PermissionError:
                        if is_copy:
                            raise
                        copy_path = suggested_copy_path(meta.path)
                        shutil.copy2(meta.path, copy_path)
                        make_writable(copy_path)
                        meta = load(str(copy_path))
                        clear_embedded_properties(meta)
                        is_copy = True
                    detail = f"已另存为：{meta.path.name}；" if is_copy else ""
                    try:
                        _, remaining = verify_embedded_clear(meta.path)
                    except Exception as verify_error:
                        verification_failed += 1
                        table.item(
                            rows[path],
                            values=(path, "无法验证", f"清除已执行，但重新读取失败：{verify_error}"),
                        )
                    else:
                        if remaining:
                            verification_failed += 1
                            table.item(
                                rows[path],
                                values=(path, "验证未通过", detail + "残留：" + "、".join(remaining[:8])),
                            )
                        else:
                            succeeded += 1
                            table.item(
                                rows[path],
                                values=(path, "成功", detail + "清除后重新读取验证通过"),
                            )
                except Exception as exc:
                    failed += 1
                    table.item(rows[path], values=(path, "失败", str(exc)))
                dialog.update_idletasks()
            progress.configure(text=f"完成：验证通过 {succeeded}，验证未通过 {verification_failed}，处理失败 {failed}")
            close_button.configure(state="normal")
            messagebox.showinfo(
                "批量清除完成",
                f"验证通过：{succeeded}\n验证未通过：{verification_failed}\n处理失败：{failed}\n详细结果见列表。",
                parent=dialog,
            )

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
            operation_backups = []
            for index, (path, (before, updates, keys)) in enumerate(snapshots.items(), start=1):
                table.item(rows[path], values=(path, "处理中", ""))
                progress.configure(text=f"正在处理 {index} / {len(snapshots)}")
                dialog.update_idletasks()
                try:
                    meta = before
                    reason = inplace_write_block_reason(meta.path)
                    is_copy = self.output_mode.get() == "copy" or bool(reason)
                    if is_copy:
                        copy_path = create_output_copy(meta.path, self.output_dir.get(), self.output_template.get())
                        meta = load(str(copy_path))
                    saved_backup = backup(meta.path) if batch_backup.get() else None
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
                        saved_backup = backup(meta.path) if batch_backup.get() else None
                        perform(meta)
                        is_copy = True
                    if saved_backup:
                        operation_backups.append((meta.path, saved_backup))
                    after = load(str(meta.path))
                    changes = changed_values(before, after, keys)
                    result_rows.extend((meta.path.name, *change) for change in changes)
                    if mode.get() == "all":
                        _, residual = verify_embedded_clear(meta.path)
                    else:
                        residual = [dict(editable_fields(before)).get(key, key) for key in updates if after.values.get(key, "") != updates[key]]
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
            self.undo_history = operation_backups
            close_button.configure(state="normal")
            self._property_table("批量处理后对照", result_rows, parent=dialog)

        start_button.configure(command=start_selected)


def launch():
    Editor().mainloop()
