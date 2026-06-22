from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .processor import (
    CorrectionRule,
    WorkbookInfo,
    apply_corrections_to_folder,
    copy_rules_to_all_points,
    discover_workbooks,
    export_a0_summary,
    load_rules,
    save_rules,
)


OP_LABELS = {
    "": "不修改",
    "add": "加 +",
    "sub": "减 -",
    "mul": "乘 ×",
    "div": "除 ÷",
    "formula": "公式",
    "replace": "替换",
}
LABEL_TO_OP = {label: op for op, label in OP_LABELS.items()}

COLORS = {
    "sidebar_bg": "#1E3A5F",
    "sidebar_hover": "#2A4A78",
    "sidebar_active": "#2E5A8E",
    "sidebar_text": "#E8EAF6",
    "sidebar_section": "#90A4AE",
    "accent": "#1565C0",
    "accent_hover": "#0D47A1",
    "accent_light": "#E3F2FD",
    "success": "#2E7D32",
    "bg": "#F5F7FA",
    "content_bg": "#FFFFFF",
    "border": "#E0E0E0",
    "row_alt": "#F8F9FA",
    "text_primary": "#212121",
    "text_secondary": "#757575",
    "white": "#FFFFFF",
    "warn_bg": "#FFF8E1",
    "warn_border": "#FFE082",
    "warn_text": "#5D4037",
    "table_header": "#1565C0",
}


class CorrectionApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("测斜 Excel 修正工具")
        self.geometry("1080x720")
        self.minsize(900, 600)
        self.configure(bg=COLORS["bg"])

        self.folder: Path | None = None
        self.workbooks: list[WorkbookInfo] = []
        self.rules: dict[str, dict[str, CorrectionRule]] = {}
        self.current_point = tk.StringVar()
        self.status = tk.StringVar(value="请选择包含原始 Excel 的文件夹。")
        self.rule_widgets: dict[tuple[str, str], tuple[tk.StringVar, tk.StringVar]] = {}

        self._setup_styles()
        self._build_shell()
        self._show_step_1()

    def _setup_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TScrollbar", background=COLORS["border"], troughcolor=COLORS["bg"], relief="flat", borderwidth=0)
        style.configure("TCombobox", fieldbackground=COLORS["content_bg"], background=COLORS["content_bg"], foreground=COLORS["text_primary"], relief="flat")
        style.configure("TEntry", fieldbackground=COLORS["content_bg"], foreground=COLORS["text_primary"], relief="flat")
        style.configure("TSeparator", background=COLORS["border"])

    def _build_shell(self) -> None:
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # ── Sidebar ──────────────────────────────────────────────
        sidebar = tk.Frame(self, bg=COLORS["sidebar_bg"], width=200)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        sidebar.columnconfigure(0, weight=1)

        logo_frame = tk.Frame(sidebar, bg=COLORS["sidebar_bg"])
        logo_frame.grid(row=0, column=0, sticky="ew", padx=18, pady=(28, 6))
        tk.Label(logo_frame, text="测斜", font=("Microsoft YaHei UI", 22, "bold"), bg=COLORS["sidebar_bg"], fg=COLORS["white"]).pack(anchor="w")
        tk.Label(logo_frame, text="Excel 修正工具", font=("Microsoft YaHei UI", 10), bg=COLORS["sidebar_bg"], fg=COLORS["sidebar_section"]).pack(anchor="w")

        tk.Frame(sidebar, bg=COLORS["sidebar_active"], height=1).grid(row=1, column=0, sticky="ew", padx=18, pady=(10, 14))

        tk.Label(sidebar, text="工 作 流 程", font=("Microsoft YaHei UI", 8), bg=COLORS["sidebar_bg"], fg=COLORS["sidebar_section"]).grid(row=2, column=0, sticky="w", padx=22, pady=(0, 4))

        nav_frame = tk.Frame(sidebar, bg=COLORS["sidebar_bg"])
        nav_frame.grid(row=3, column=0, sticky="ew", padx=8)
        nav_frame.columnconfigure(0, weight=1)
        self._nav_btn(nav_frame, 0, "①  选择文件夹", self._show_step_1)
        self._nav_btn(nav_frame, 1, "②  设置修正", self._show_step_2)
        self._nav_btn(nav_frame, 2, "③  生成文件", self._show_step_3)

        tk.Frame(sidebar, bg=COLORS["sidebar_active"], height=1).grid(row=4, column=0, sticky="ew", padx=18, pady=12)

        tk.Label(sidebar, text="工 具", font=("Microsoft YaHei UI", 8), bg=COLORS["sidebar_bg"], fg=COLORS["sidebar_section"]).grid(row=5, column=0, sticky="w", padx=22, pady=(0, 4))

        tools_frame = tk.Frame(sidebar, bg=COLORS["sidebar_bg"])
        tools_frame.grid(row=6, column=0, sticky="ew", padx=8)
        tools_frame.columnconfigure(0, weight=1)
        self._nav_btn(tools_frame, 0, "  保存参数", self._save_rules_dialog)
        self._nav_btn(tools_frame, 1, "  加载参数", self._load_rules_dialog)
        self._nav_btn(tools_frame, 2, "  汇总 A0", self._export_a0_summary_dialog)

        sidebar.rowconfigure(7, weight=1)
        tk.Label(sidebar, text="v1.0", font=("Microsoft YaHei UI", 8), bg=COLORS["sidebar_bg"], fg=COLORS["sidebar_section"]).grid(row=8, column=0, sticky="sw", padx=18, pady=12)

        # ── Main area ────────────────────────────────────────────
        main = tk.Frame(self, bg=COLORS["bg"])
        main.grid(row=0, column=1, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=1)

        self.content = tk.Frame(main, bg=COLORS["bg"])
        self.content.grid(row=0, column=0, sticky="nsew")
        self.content.columnconfigure(0, weight=1)
        self.content.rowconfigure(0, weight=1)

        tk.Frame(main, bg=COLORS["border"], height=1).grid(row=1, column=0, sticky="ew")

        status_bar = tk.Frame(main, bg=COLORS["bg"])
        status_bar.grid(row=2, column=0, sticky="ew")
        tk.Label(status_bar, textvariable=self.status, font=("Microsoft YaHei UI", 9), bg=COLORS["bg"], fg=COLORS["text_secondary"]).pack(side="left", padx=16, pady=7)

    def _nav_btn(self, parent: tk.Frame, row: int, text: str, command) -> tk.Button:
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=("Microsoft YaHei UI", 10),
            bg=COLORS["sidebar_bg"],
            fg=COLORS["sidebar_text"],
            activebackground=COLORS["sidebar_hover"],
            activeforeground=COLORS["white"],
            relief="flat",
            bd=0,
            anchor="w",
            padx=14,
            pady=9,
            cursor="hand2",
        )
        btn.grid(row=row, column=0, sticky="ew", pady=1)
        btn.bind("<Enter>", lambda e: btn.configure(bg=COLORS["sidebar_hover"]))
        btn.bind("<Leave>", lambda e: btn.configure(bg=COLORS["sidebar_bg"]))
        return btn

    def _primary_btn(self, parent: tk.Widget, text: str, command) -> tk.Button:
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=("Microsoft YaHei UI", 10, "bold"),
            bg=COLORS["accent"],
            fg=COLORS["white"],
            activebackground=COLORS["accent_hover"],
            activeforeground=COLORS["white"],
            relief="flat",
            bd=0,
            padx=20,
            pady=8,
            cursor="hand2",
        )
        btn.bind("<Enter>", lambda e: btn.configure(bg=COLORS["accent_hover"]))
        btn.bind("<Leave>", lambda e: btn.configure(bg=COLORS["accent"]))
        return btn

    def _secondary_btn(self, parent: tk.Widget, text: str, command) -> tk.Button:
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=("Microsoft YaHei UI", 10),
            bg=COLORS["content_bg"],
            fg=COLORS["accent"],
            activebackground=COLORS["accent_light"],
            activeforeground=COLORS["accent"],
            relief="solid",
            bd=1,
            padx=12,
            pady=6,
            cursor="hand2",
            highlightbackground=COLORS["accent"],
            highlightthickness=1,
        )
        btn.bind("<Enter>", lambda e: btn.configure(bg=COLORS["accent_light"]))
        btn.bind("<Leave>", lambda e: btn.configure(bg=COLORS["content_bg"]))
        return btn

    def _success_btn(self, parent: tk.Widget, text: str, command) -> tk.Button:
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            font=("Microsoft YaHei UI", 11, "bold"),
            bg=COLORS["success"],
            fg=COLORS["white"],
            activebackground="#1B5E20",
            activeforeground=COLORS["white"],
            relief="flat",
            bd=0,
            padx=24,
            pady=10,
            cursor="hand2",
        )
        btn.bind("<Enter>", lambda e: btn.configure(bg="#1B5E20"))
        btn.bind("<Leave>", lambda e: btn.configure(bg=COLORS["success"]))
        return btn

    def _page_header(self, title: str) -> None:
        header = tk.Frame(self.content, bg=COLORS["content_bg"], height=60)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        tk.Frame(header, bg=COLORS["accent"], width=5).pack(side="left", fill="y")
        tk.Label(header, text=title, font=("Microsoft YaHei UI", 16, "bold"), bg=COLORS["content_bg"], fg=COLORS["text_primary"]).pack(side="left", padx=20)
        tk.Frame(self.content, bg=COLORS["border"], height=1).grid(row=1, column=0, sticky="ew")

    def _clear_content(self) -> None:
        for child in self.content.winfo_children():
            child.destroy()
        self.content.columnconfigure(0, weight=1)
        self.content.rowconfigure(0, weight=1)

    def _show_step_1(self) -> None:
        self._clear_content()
        self._page_header("Step 1  选择原始 Excel 文件夹")

        body = tk.Frame(self.content, bg=COLORS["bg"])
        body.grid(row=2, column=0, sticky="nsew", padx=24, pady=20)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(4, weight=1)
        self.content.rowconfigure(2, weight=1)

        tk.Label(
            body,
            text="程序只读取该文件夹下的 .xlsx 文件，不会修改原始文件；多 sheet 工作簿会把每个 sheet 当作一个点号处理。",
            font=("Microsoft YaHei UI", 10),
            bg=COLORS["bg"],
            fg=COLORS["text_secondary"],
            wraplength=620,
            justify="left",
        ).grid(row=0, column=0, sticky="w", pady=(0, 14))

        self._primary_btn(body, "选择文件夹", self._choose_folder).grid(row=1, column=0, sticky="w")

        self.folder_label = tk.Label(body, text=self._folder_text(), font=("Microsoft YaHei UI", 10), bg=COLORS["bg"], fg=COLORS["text_secondary"])
        self.folder_label.grid(row=2, column=0, sticky="w", pady=(14, 6))

        tk.Frame(body, bg=COLORS["border"], height=1).grid(row=3, column=0, sticky="ew", pady=(0, 8))

        # File list card
        list_card = tk.Frame(body, bg=COLORS["content_bg"], highlightbackground=COLORS["border"], highlightthickness=1)
        list_card.grid(row=4, column=0, sticky="nsew")
        list_card.columnconfigure(0, weight=1)
        list_card.rowconfigure(1, weight=1)

        list_hdr = tk.Frame(list_card, bg=COLORS["accent_light"])
        list_hdr.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Label(list_hdr, text="已识别的 Excel 文件", font=("Microsoft YaHei UI", 9, "bold"), bg=COLORS["accent_light"], fg=COLORS["accent"]).pack(side="left", padx=12, pady=6)

        self.files_box = tk.Listbox(
            list_card,
            height=10,
            font=("Microsoft YaHei UI", 10),
            bg=COLORS["content_bg"],
            fg=COLORS["text_primary"],
            selectbackground=COLORS["accent"],
            selectforeground=COLORS["white"],
            bd=0,
            highlightthickness=0,
            activestyle="none",
        )
        sb = ttk.Scrollbar(list_card, orient="vertical", command=self.files_box.yview)
        self.files_box.configure(yscrollcommand=sb.set)
        self.files_box.grid(row=1, column=0, sticky="nsew", padx=(8, 0), pady=8)
        sb.grid(row=1, column=1, sticky="ns", pady=8)

        self._refresh_files_box()

    def _show_step_2(self) -> None:
        self._clear_content()
        if not self.workbooks:
            self.status.set("请先在 Step 1 选择包含 .xlsx 的文件夹。")
            self._show_step_1()
            return

        self._page_header("Step 2  按点号和列设置修正规则")

        body = tk.Frame(self.content, bg=COLORS["bg"])
        body.grid(row=2, column=0, sticky="nsew", padx=24, pady=16)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(3, weight=1)
        self.content.rowconfigure(2, weight=1)

        # Hint banners
        hints = [
            "公式说明：x 表示本列原值，[列名] 表示同一行其他列原值，rand() 表示 0~1 随机数，例如 x + [A0] * 0.1 + rand()。",
            "自动重算：A轴管口/管底起算、B轴管口/管底起算会根据修正后的 A0/A180/B0/B180 自动更新。",
        ]
        for i, hint in enumerate(hints):
            banner = tk.Frame(body, bg=COLORS["warn_bg"], highlightbackground=COLORS["warn_border"], highlightthickness=1)
            banner.grid(row=i, column=0, sticky="ew", pady=(0, 6))
            tk.Label(banner, text=hint, font=("Microsoft YaHei UI", 9), bg=COLORS["warn_bg"], fg=COLORS["warn_text"], wraplength=700, justify="left").pack(anchor="w", padx=12, pady=6)

        # Point selector toolbar
        toolbar = tk.Frame(body, bg=COLORS["content_bg"], highlightbackground=COLORS["border"], highlightthickness=1)
        toolbar.grid(row=2, column=0, sticky="ew", pady=(4, 8))

        point_ids = sorted({info.point_id for info in self.workbooks})
        if self.current_point.get() not in point_ids:
            self.current_point.set(point_ids[0])

        tk.Label(toolbar, text="当前点号：", font=("Microsoft YaHei UI", 10), bg=COLORS["content_bg"], fg=COLORS["text_primary"]).pack(side="left", padx=(14, 4), pady=8)
        point_combo = ttk.Combobox(toolbar, textvariable=self.current_point, values=point_ids, state="readonly", width=20)
        point_combo.pack(side="left", pady=8)
        point_combo.bind("<<ComboboxSelected>>", self._change_point)
        self._secondary_btn(toolbar, "应用到全部点号", self._apply_current_rules_to_all_points).pack(side="left", padx=(12, 0), pady=8)

        # Rules table
        self.rules_holder = tk.Frame(body, bg=COLORS["bg"])
        self.rules_holder.grid(row=3, column=0, sticky="nsew")
        self.rules_holder.columnconfigure(0, weight=1)
        self.rules_holder.rowconfigure(0, weight=1)
        self._render_rule_table()

    def _show_step_3(self) -> None:
        self._clear_content()
        self._page_header("Step 3  生成修正后 Excel")

        body = tk.Frame(self.content, bg=COLORS["bg"])
        body.grid(row=2, column=0, sticky="nsew", padx=24, pady=24)
        body.columnconfigure(0, weight=1)
        self.content.rowconfigure(2, weight=1)

        # Info card
        info_card = tk.Frame(body, bg=COLORS["content_bg"], highlightbackground=COLORS["border"], highlightthickness=1)
        info_card.grid(row=0, column=0, sticky="ew", pady=(0, 20))

        rows = [
            ("原始文件夹", self._folder_text()),
            ("输出位置", "原始文件夹下的 corrected 子文件夹"),
            ("文件命名", "原文件名 + -修正后.xlsx；如已存在会自动追加时间后缀"),
        ]
        for i, (label, value) in enumerate(rows):
            row_frame = tk.Frame(info_card, bg=COLORS["content_bg"])
            row_frame.pack(fill="x", padx=18, pady=(14 if i == 0 else 8, 8 if i < len(rows) - 1 else 16))
            tk.Label(row_frame, text=label, font=("Microsoft YaHei UI", 10, "bold"), bg=COLORS["content_bg"], fg=COLORS["text_secondary"], width=16, anchor="w").pack(side="left")
            tk.Label(row_frame, text=value, font=("Microsoft YaHei UI", 10), bg=COLORS["content_bg"], fg=COLORS["text_primary"]).pack(side="left")
            if i < len(rows) - 1:
                tk.Frame(info_card, bg=COLORS["border"], height=1).pack(fill="x", padx=18)

        self._success_btn(body, "生成修正后 Excel", self._generate_files).grid(row=1, column=0, sticky="w")

    def _choose_folder(self) -> None:
        selected = filedialog.askdirectory(title="选择包含原始 Excel 的文件夹")
        if not selected:
            return
        try:
            self.folder = Path(selected)
            self.workbooks = discover_workbooks(self.folder)
            self._ensure_rules_for_workbooks()
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))
            return

        self.status.set(f"已识别 {len(self.workbooks)} 个 Excel 文件。")
        self.folder_label.configure(text=self._folder_text())
        self._refresh_files_box()

    def _refresh_files_box(self) -> None:
        if not hasattr(self, "files_box"):
            return
        self.files_box.delete(0, tk.END)
        for info in self.workbooks:
            self.files_box.insert(tk.END, f"  {info.point_id}  |  {info.path.name}  |  {len(info.headers)} 列")

    def _render_rule_table(self) -> None:
        for child in self.rules_holder.winfo_children():
            child.destroy()
        point_id = self.current_point.get()
        headers = self._headers_for_point(point_id)

        canvas = tk.Canvas(self.rules_holder, highlightthickness=0, bg=COLORS["bg"])
        scrollbar = ttk.Scrollbar(self.rules_holder, orient="vertical", command=canvas.yview)
        table = tk.Frame(canvas, bg=COLORS["content_bg"])
        table.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=table, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Table header row
        hdr = tk.Frame(table, bg=COLORS["table_header"])
        hdr.pack(fill="x")
        for col_text, col_width in [("列  名", 22), ("操  作", 12), ("值 / 公式", 28)]:
            tk.Label(hdr, text=col_text, font=("Microsoft YaHei UI", 9, "bold"), bg=COLORS["table_header"], fg=COLORS["white"], width=col_width, anchor="w").pack(side="left", padx=8, pady=7)

        self.rule_widgets = {}
        point_rules = self.rules.setdefault(point_id, {})
        for row_idx, header in enumerate(headers):
            rule = point_rules.get(header, CorrectionRule("", ""))
            op_var = tk.StringVar(value=OP_LABELS.get(rule.op, "不修改"))
            value_var = tk.StringVar(value=rule.value)

            row_bg = COLORS["content_bg"] if row_idx % 2 == 0 else COLORS["row_alt"]
            row_frame = tk.Frame(table, bg=row_bg)
            row_frame.pack(fill="x")
            tk.Label(row_frame, text=header, font=("Microsoft YaHei UI", 10), bg=row_bg, fg=COLORS["text_primary"], width=22, anchor="w").grid(row=0, column=0, padx=8, pady=5, sticky="w")
            ttk.Combobox(row_frame, textvariable=op_var, values=list(OP_LABELS.values()), state="readonly", width=12).grid(row=0, column=1, padx=6, pady=5)
            ttk.Entry(row_frame, textvariable=value_var, width=28).grid(row=0, column=2, padx=6, pady=5)
            self.rule_widgets[(point_id, header)] = (op_var, value_var)

        btn_row = tk.Frame(table, bg=COLORS["content_bg"])
        btn_row.pack(fill="x", padx=8, pady=12)
        self._primary_btn(btn_row, "应用当前点号设置", self._collect_visible_rules).pack(side="left")

    def _change_point(self, _event: tk.Event) -> None:
        self._collect_visible_rules_if_present()
        self._render_rule_table()

    def _apply_current_rules_to_all_points(self) -> None:
        self._collect_visible_rules_if_present()
        source_point = self.current_point.get()
        self.rules = copy_rules_to_all_points(source_point, self.workbooks, self.rules)
        self._ensure_rules_for_workbooks()
        self._render_rule_table()
        self.status.set(f"已将 {source_point} 的修正规则应用到全部点号。")

    def _collect_visible_rules(self) -> None:
        for (point_id, header), (op_var, value_var) in self.rule_widgets.items():
            op = LABEL_TO_OP.get(op_var.get(), "")
            value = value_var.get().strip()
            self.rules.setdefault(point_id, {})[header] = CorrectionRule(op, value)
        self.status.set(f"已应用 {self.current_point.get()} 的修正规则。")

    def _save_rules_dialog(self) -> None:
        self._collect_visible_rules_if_present()
        if not self.rules:
            messagebox.showinfo("没有参数", "请先选择文件夹并设置修正规则。")
            return
        path = filedialog.asksaveasfilename(
            title="保存参数文件",
            defaultextension=".json",
            filetypes=[("参数文件", "*.json"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            save_rules(path, self.rules)
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))
            return
        self.status.set(f"参数已保存：{path}")

    def _load_rules_dialog(self) -> None:
        path = filedialog.askopenfilename(title="加载参数文件", filetypes=[("参数文件", "*.json"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            loaded = load_rules(path)
        except Exception as exc:
            messagebox.showerror("加载失败", str(exc))
            return
        self.rules.update(loaded)
        self._ensure_rules_for_workbooks()
        self.status.set(f"参数已加载：{path}")
        if self.workbooks:
            self._show_step_2()

    def _export_a0_summary_dialog(self) -> None:
        path = filedialog.askopenfilename(
            title="选择要汇总 A0 的 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            output = export_a0_summary(path)
        except Exception as exc:
            messagebox.showerror("汇总失败", str(exc))
            return

        message = f"A0 汇总表已生成：\n{output}"
        messagebox.showinfo("汇总完成", message)
        self.status.set(message.replace("\n", " "))

    def _generate_files(self) -> None:
        self._collect_visible_rules_if_present()
        if self.folder is None:
            messagebox.showinfo("请选择文件夹", "请先在 Step 1 选择包含原始 Excel 的文件夹。")
            return
        try:
            summary = apply_corrections_to_folder(self.folder, self.rules)
        except Exception as exc:
            messagebox.showerror("生成失败", str(exc))
            return

        message = f"已生成 {summary.processed_files} 个文件。"
        if summary.skipped_cells:
            message += f"\n有 {summary.skipped_cells} 个单元格不是数字，已保持原值。"
        if summary.skipped_files:
            message += f"\n有 {len(summary.skipped_files)} 个文件没有修正规则，已跳过。"
        if summary.outputs:
            message += f"\n输出文件夹：{summary.outputs[0].parent}"
        messagebox.showinfo("生成完成", message)
        self.status.set(message.replace("\n", " "))

    def _collect_visible_rules_if_present(self) -> None:
        if self.rule_widgets:
            self._collect_visible_rules()

    def _ensure_rules_for_workbooks(self) -> None:
        for info in self.workbooks:
            point_rules = self.rules.setdefault(info.point_id, {})
            for header in info.headers:
                point_rules.setdefault(header, CorrectionRule("", ""))

    def _headers_for_point(self, point_id: str) -> list[str]:
        for info in self.workbooks:
            if info.point_id == point_id:
                return info.headers
        return []

    def _folder_text(self) -> str:
        return str(self.folder) if self.folder else "尚未选择"


def main() -> None:
    app = CorrectionApp()
    app.mainloop()


if __name__ == "__main__":
    main()
