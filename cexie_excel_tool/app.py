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


class CorrectionApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("测斜 Excel 修正工具")
        self.geometry("980x680")
        self.minsize(900, 600)

        self.folder: Path | None = None
        self.workbooks: list[WorkbookInfo] = []
        self.rules: dict[str, dict[str, CorrectionRule]] = {}
        self.current_point = tk.StringVar()
        self.status = tk.StringVar(value="请选择包含原始 Excel 的文件夹。")
        self.rule_widgets: dict[tuple[str, str], tuple[tk.StringVar, tk.StringVar]] = {}

        self._build_shell()
        self._show_step_1()

    def _build_shell(self) -> None:
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(self, padding=16)
        sidebar.grid(row=0, column=0, sticky="ns")
        ttk.Label(sidebar, text="测斜 Excel\n修正工具", font=("Microsoft YaHei UI", 16, "bold")).pack(anchor="w", pady=(0, 24))
        ttk.Button(sidebar, text="1 选择文件夹", command=self._show_step_1).pack(fill="x", pady=4)
        ttk.Button(sidebar, text="2 设置修正", command=self._show_step_2).pack(fill="x", pady=4)
        ttk.Button(sidebar, text="3 生成文件", command=self._show_step_3).pack(fill="x", pady=4)
        ttk.Separator(sidebar).pack(fill="x", pady=16)
        ttk.Button(sidebar, text="保存参数", command=self._save_rules_dialog).pack(fill="x", pady=4)
        ttk.Button(sidebar, text="加载参数", command=self._load_rules_dialog).pack(fill="x", pady=4)

        main = ttk.Frame(self, padding=(8, 16, 16, 16))
        main.grid(row=0, column=1, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=1)

        self.content = ttk.Frame(main)
        self.content.grid(row=0, column=0, sticky="nsew")
        self.content.columnconfigure(0, weight=1)
        self.content.rowconfigure(0, weight=1)

        ttk.Label(main, textvariable=self.status, foreground="#555").grid(row=1, column=0, sticky="ew", pady=(12, 0))

    def _clear_content(self) -> None:
        for child in self.content.winfo_children():
            child.destroy()
        self.content.columnconfigure(0, weight=1)
        self.content.rowconfigure(0, weight=1)

    def _show_step_1(self) -> None:
        self._clear_content()
        frame = ttk.Frame(self.content, padding=24)
        frame.grid(sticky="nsew")
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="Step 1 选择原始 Excel 文件夹", font=("Microsoft YaHei UI", 18, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(frame, text="程序只读取该文件夹下的 .xlsx 文件，不会修改原始文件。").grid(row=1, column=0, sticky="w", pady=(8, 20))
        ttk.Button(frame, text="选择文件夹", command=self._choose_folder).grid(row=2, column=0, sticky="w")

        self.folder_label = ttk.Label(frame, text=self._folder_text(), foreground="#333")
        self.folder_label.grid(row=3, column=0, sticky="w", pady=(16, 8))
        self.files_box = tk.Listbox(frame, height=12)
        self.files_box.grid(row=4, column=0, sticky="nsew")
        frame.rowconfigure(4, weight=1)
        self._refresh_files_box()

    def _show_step_2(self) -> None:
        self._clear_content()
        if not self.workbooks:
            self.status.set("请先在 Step 1 选择包含 .xlsx 的文件夹。")
            self._show_step_1()
            return

        frame = ttk.Frame(self.content, padding=16)
        frame.grid(sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)

        ttk.Label(frame, text="Step 2 按点号和列设置修正规则", font=("Microsoft YaHei UI", 18, "bold")).grid(row=0, column=0, sticky="w")

        point_ids = sorted({info.point_id for info in self.workbooks})
        if self.current_point.get() not in point_ids:
            self.current_point.set(point_ids[0])
        selector = ttk.Frame(frame)
        selector.grid(row=1, column=0, sticky="ew", pady=(12, 10))
        ttk.Label(selector, text="点号：").pack(side="left")
        point_combo = ttk.Combobox(selector, textvariable=self.current_point, values=point_ids, state="readonly", width=24)
        point_combo.pack(side="left")
        point_combo.bind("<<ComboboxSelected>>", self._change_point)
        ttk.Button(selector, text="应用到全部点号", command=self._apply_current_rules_to_all_points).pack(side="left", padx=(12, 0))

        ttk.Label(
            frame,
            text="公式说明：选择“公式”后，在值/公式里写表达式；x 表示原值，rand() 表示 0~1 均匀随机数，例如 x + rand() * 0.5。",
            foreground="#555",
        ).grid(row=2, column=0, sticky="w", pady=(0, 8))

        self.rules_holder = ttk.Frame(frame)
        self.rules_holder.grid(row=3, column=0, sticky="nsew")
        self.rules_holder.columnconfigure(0, weight=1)
        self.rules_holder.rowconfigure(0, weight=1)
        self._render_rule_table()

    def _show_step_3(self) -> None:
        self._clear_content()
        frame = ttk.Frame(self.content, padding=24)
        frame.grid(sticky="nsew")
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="Step 3 生成修正后 Excel", font=("Microsoft YaHei UI", 18, "bold")).grid(row=0, column=0, sticky="w")
        folder_text = self._folder_text()
        ttk.Label(frame, text=f"原始文件夹：{folder_text}").grid(row=1, column=0, sticky="w", pady=(12, 4))
        ttk.Label(frame, text="输出位置：原始文件夹下的 corrected 子文件夹").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Label(frame, text="文件命名：原文件名 + -修正后.xlsx；如已存在会自动追加时间后缀。").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Button(frame, text="生成修正后 Excel", command=self._generate_files).grid(row=4, column=0, sticky="w", pady=(20, 0))

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
            self.files_box.insert(tk.END, f"{info.point_id} | {info.path.name} | {len(info.headers)} 列")

    def _render_rule_table(self) -> None:
        for child in self.rules_holder.winfo_children():
            child.destroy()
        point_id = self.current_point.get()
        headers = self._headers_for_point(point_id)

        canvas = tk.Canvas(self.rules_holder, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.rules_holder, orient="vertical", command=canvas.yview)
        table = ttk.Frame(canvas, padding=(0, 0, 12, 0))
        table.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=table, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        ttk.Label(table, text="列名", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, sticky="w", padx=6, pady=6)
        ttk.Label(table, text="操作", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=1, sticky="w", padx=6, pady=6)
        ttk.Label(table, text="值/公式", font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=2, sticky="w", padx=6, pady=6)

        self.rule_widgets = {}
        point_rules = self.rules.setdefault(point_id, {})
        for row, header in enumerate(headers, start=1):
            rule = point_rules.get(header, CorrectionRule("", ""))
            op_var = tk.StringVar(value=OP_LABELS.get(rule.op, "不修改"))
            value_var = tk.StringVar(value=rule.value)
            ttk.Label(table, text=header, width=24).grid(row=row, column=0, sticky="w", padx=6, pady=4)
            ttk.Combobox(table, textvariable=op_var, values=list(OP_LABELS.values()), state="readonly", width=12).grid(row=row, column=1, sticky="w", padx=6, pady=4)
            ttk.Entry(table, textvariable=value_var, width=28).grid(row=row, column=2, sticky="w", padx=6, pady=4)
            self.rule_widgets[(point_id, header)] = (op_var, value_var)

        ttk.Button(table, text="应用当前点号设置", command=self._collect_visible_rules).grid(row=len(headers) + 1, column=0, sticky="w", padx=6, pady=16)

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
