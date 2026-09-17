from __future__ import annotations

import shutil
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from .engine import FIELDS, PDF_FIELDS, Metadata, backup, clear_embedded_properties, inplace_write_block_reason, load, make_writable, set_filesystem_times, suggested_copy_path, write_office, write_pdf


class Editor(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("文件源数据编辑")
        self.after_idle(self._fit_initial_window)
        self.meta: Metadata | None = None
        self.vars: dict[str, tk.StringVar] = {}
        self.backup_var = tk.BooleanVar(value=True)
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
        ttk.Button(chooser, text="批量清除", command=self.batch_clear).pack(side="left", padx=(8, 0))
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
        # Let the wheel work naturally even when the pointer is above an entry,
        # label, or nested form widget rather than only above the scrollbar.
        self.bind_all("<MouseWheel>", self._mousewheel, add="+")

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
        path = filedialog.askopenfilename(filetypes=[("支持的文件", "*.docx *.xlsx *.pptx *.docm *.xlsm *.pptm *.dotx *.xltx *.potx *.pdf"), ("所有文件", "*.*")])
        if path: self.open(path)

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
        for key, label in (PDF_FIELDS if self.meta.kind == "PDF" else FIELDS): self.tree.insert(props, "end", text=label, values=(self.meta.values.get(key, ""),))
        if self.meta.kind == "Office":
            custom_node = self.tree.insert("", "end", text="自定义属性", open=True)
            for key, value in self.meta.custom.items(): self.tree.insert(custom_node, "end", text=key, values=(value,))
        self.vars = {}
        self._section("文件系统时间", [("filesystem_created", "创建时间"), ("filesystem_modified", "修改时间")], readonly=False)
        groups = [("来源与描述", [x for x in (PDF_FIELDS if self.meta.kind == "PDF" else FIELDS) if x[0] not in {"created", "modified", "lastPrinted", "totalTime"}]), ("文档元数据时间", [x for x in (PDF_FIELDS if self.meta.kind == "PDF" else FIELDS) if x[0] in {"created", "modified", "lastPrinted", "totalTime"}])]
        for title, fields in groups: self._section(title, fields)
        if self.meta.kind == "Office":
            box = ttk.Labelframe(self.form, text=" 自定义属性（每行：名称 = 值） ", style="Card.TLabelframe", padding=10); box.pack(fill="x", pady=(0, 12))
            self.custom_text = tk.Text(box, height=5, font=("Consolas", 10), relief="solid", borderwidth=1)
            self.custom_text.pack(fill="x"); self.custom_text.insert("1.0", "\n".join(f"{k} = {v}" for k, v in self.meta.custom.items()))
        else:
            ttk.Label(self.form, text="PDF 会重写元数据保存。加密或已签名 PDF 将拒绝修改。", foreground="#a16207", background="#ffffff").pack(anchor="w", pady=8)

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
            initialdir=str(source.path.parent),
            initialfile=suggested_copy_path(source.path, label).name,
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
        reason = inplace_write_block_reason(self.meta.path)
        if not reason:
            return self.meta, False
        copy = self._copy_target(self.meta, action, label, reason)
        return (copy, True) if copy else None

    def save(self):
        if not self.meta: return
        try:
            target = self._target_for_write("保存", "已修改属性")
            if target is None: return
            meta, is_copy = target
            if self.backup_var.get() and not is_copy: backup(meta.path)
            values = {key: var.get().strip() for key, var in self.vars.items()}
            try:
                if meta.kind == "Office": write_office(meta, values, self._custom(), False)
                else: write_pdf(meta, values, False)
            except PermissionError:
                if is_copy: raise
                copy = self._copy_target(meta, "保存", "已修改属性", "Windows 拒绝替换原文件；它可能仍被其他程序占用")
                if copy is None: return
                meta, is_copy = copy, True
                if meta.kind == "Office": write_office(meta, values, self._custom(), False)
                else: write_pdf(meta, values, False)
            set_filesystem_times(meta.path, values["filesystem_created"], values["filesystem_modified"])
            self.open(str(meta.path))
            self.status.configure(text="已另存为副本并保存。" if is_copy else "保存成功，已重新读取验证。")
        except Exception as e: messagebox.showerror("保存失败", str(e))

    def clear(self):
        if not self.meta: return
        if not messagebox.askyesno("确认清除", "将清空文档内嵌属性和自定义属性；不会更改 Windows 文件系统时间。是否继续？"): return
        try:
            target = self._target_for_write("清除属性", "已清除属性")
            if target is None: return
            meta, is_copy = target
            if self.backup_var.get() and not is_copy: backup(meta.path)
            try:
                clear_embedded_properties(meta)
            except PermissionError:
                if is_copy: raise
                copy = self._copy_target(meta, "清除属性", "已清除属性", "Windows 拒绝替换原文件；它可能仍被其他程序占用")
                if copy is None: return
                meta, is_copy = copy, True
                clear_embedded_properties(meta)
            self.open(str(meta.path))
            self.status.configure(text="已另存为副本并清除属性。" if is_copy else "嵌入属性已清除。")
        except Exception as e: messagebox.showerror("清除失败", str(e))

    def batch_clear(self):
        paths = filedialog.askopenfilenames(
            title="选择要清除属性的文件",
            filetypes=[("支持的文件", "*.docx *.xlsx *.pptx *.docm *.xlsm *.pptm *.dotx *.xltx *.potx *.dotm *.xltm *.potm *.ppsx *.ppsm *.pdf")],
        )
        if not paths:
            return
        self._show_batch_dialog(list(dict.fromkeys(paths)))

    def _show_batch_dialog(self, paths: list[str]):
        dialog = tk.Toplevel(self)
        dialog.title("批量清除嵌入属性")
        dialog.transient(self)
        dialog.grab_set()
        dialog.minsize(760, 430)
        dialog.geometry("900x560")
        frame = ttk.Frame(dialog, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="批量清除嵌入属性", style="Title.TLabel").pack(anchor="w")
        ttk.Label(frame, text=f"已选择 {len(paths)} 个文件。将清除 Office/PDF 的内嵌属性，不修改 Windows 文件时间。", style="Muted.TLabel").pack(anchor="w", pady=(3, 12))
        table = ttk.Treeview(frame, columns=("file", "status", "detail"), show="headings")
        table.heading("file", text="文件")
        table.heading("status", text="状态")
        table.heading("detail", text="结果 / 原因")
        table.column("file", width=420, minwidth=250)
        table.column("status", width=90, anchor="center", stretch=False)
        table.column("detail", width=320, minwidth=180)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=table.yview)
        table.configure(yscrollcommand=scrollbar.set)
        table.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        rows = {path: table.insert("", "end", values=(path, "等待", "")) for path in paths}

        footer = ttk.Frame(dialog, padding=(16, 0, 16, 16))
        footer.pack(fill="x")
        batch_backup = tk.BooleanVar(value=self.backup_var.get())
        ttk.Label(footer, text="保存前备份：").pack(side="left")
        ttk.Radiobutton(footer, text="是", variable=batch_backup, value=True).pack(side="left", padx=(2, 8))
        ttk.Radiobutton(footer, text="否", variable=batch_backup, value=False).pack(side="left")
        progress = ttk.Label(footer, text="等待开始", style="Muted.TLabel")
        progress.pack(side="left", padx=16)
        close_button = ttk.Button(footer, text="关闭", command=dialog.destroy)
        close_button.pack(side="right")
        start_button = ttk.Button(footer, text="开始清除", style="Accent.TButton")
        start_button.pack(side="right", padx=(0, 8))

        def start():
            if not messagebox.askyesno("确认批量清除", f"确认清除这 {len(paths)} 个文件的所有内嵌属性吗？\n此操作将无法撤销。", parent=dialog):
                return
            start_button.configure(state="disabled")
            close_button.configure(state="disabled")
            succeeded = failed = 0
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
                    succeeded += 1
                    detail = f"已另存为：{meta.path.name}" if is_copy else "属性已清除"
                    table.item(rows[path], values=(path, "成功", detail))
                except Exception as exc:
                    failed += 1
                    table.item(rows[path], values=(path, "失败", str(exc)))
                dialog.update_idletasks()
            progress.configure(text=f"处理完成：成功 {succeeded}，失败 {failed}")
            close_button.configure(state="normal")
            messagebox.showinfo("批量清除完成", f"成功：{succeeded}\n失败：{failed}\n详细结果见列表。", parent=dialog)

        start_button.configure(command=start)


def launch():
    Editor().mainloop()
