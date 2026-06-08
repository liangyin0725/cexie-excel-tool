from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Mapping

from openpyxl import load_workbook


POINT_ID_RE = re.compile(r"\b\d+[A-Z]?-ZQTS?\d+\b", re.IGNORECASE)
VALID_OPS = {"add", "sub", "mul", "div", "replace"}


@dataclass(frozen=True)
class CorrectionRule:
    op: str
    value: str


@dataclass(frozen=True)
class WorkbookInfo:
    path: Path
    point_id: str
    sheet_name: str
    headers: list[str]


@dataclass(frozen=True)
class ProcessSummary:
    processed_files: int
    outputs: list[Path]
    skipped_cells: int
    skipped_files: list[str]


RuleMap = dict[str, dict[str, CorrectionRule]]


def discover_workbooks(folder: str | Path) -> list[WorkbookInfo]:
    root = Path(folder)
    if not root.exists() or not root.is_dir():
        raise ValueError(f"文件夹不存在: {root}")

    workbooks: list[WorkbookInfo] = []
    for path in sorted(root.glob("*.xlsx")):
        if path.name.startswith("~$"):
            continue
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        headers = ["" if cell.value is None else str(cell.value) for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        point_id = _extract_point_id(path.name) or _extract_point_id(ws.title) or path.stem
        workbooks.append(WorkbookInfo(path=path, point_id=point_id, sheet_name=ws.title, headers=headers))
        wb.close()
    return workbooks


def apply_corrections_to_folder(folder: str | Path, rules: Mapping[str, Mapping[str, CorrectionRule]]) -> ProcessSummary:
    root = Path(folder)
    output_dir = root / "corrected"
    output_dir.mkdir(exist_ok=True)

    outputs: list[Path] = []
    skipped_files: list[str] = []
    skipped_cells = 0

    for info in discover_workbooks(root):
        point_rules = dict(rules.get(info.point_id, {}))
        if not _has_effective_rules(point_rules):
            skipped_files.append(info.path.name)
            continue

        _validate_rules(point_rules)
        wb = load_workbook(info.path)
        ws = wb.worksheets[0]
        header_to_column = {
            "" if cell.value is None else str(cell.value): cell.column
            for cell in ws[1]
        }

        for header, rule in point_rules.items():
            if not rule.op or header not in header_to_column:
                continue
            column_index = header_to_column[header]
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=column_index)
                new_value, skipped = _apply_rule(cell.value, rule)
                if skipped:
                    skipped_cells += 1
                    continue
                cell.value = new_value

        output_path = _unique_output_path(output_dir, info.path)
        wb.save(output_path)
        wb.close()
        outputs.append(output_path)

    return ProcessSummary(
        processed_files=len(outputs),
        outputs=outputs,
        skipped_cells=skipped_cells,
        skipped_files=skipped_files,
    )


def save_rules(path: str | Path, rules: Mapping[str, Mapping[str, CorrectionRule]]) -> None:
    payload = {
        "version": 1,
        "points": {
            point_id: {
                header: {"op": rule.op, "value": rule.value}
                for header, rule in column_rules.items()
                if rule.op and rule.value != ""
            }
            for point_id, column_rules in rules.items()
        },
    }
    Path(path).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def load_rules(path: str | Path) -> RuleMap:
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    if payload.get("version") != 1:
        raise ValueError("不支持的参数文件版本")
    loaded: RuleMap = {}
    for point_id, column_rules in payload.get("points", {}).items():
        loaded[point_id] = {}
        for header, raw_rule in column_rules.items():
            rule = CorrectionRule(str(raw_rule.get("op", "")), str(raw_rule.get("value", "")))
            if rule.op:
                _validate_rule(rule)
                loaded[point_id][header] = rule
    return loaded


def _extract_point_id(text: str) -> str | None:
    match = POINT_ID_RE.search(text)
    return match.group(0).upper() if match else None


def _has_effective_rules(rules: Mapping[str, CorrectionRule]) -> bool:
    return any(rule.op and rule.value != "" for rule in rules.values())


def _validate_rules(rules: Mapping[str, CorrectionRule]) -> None:
    for rule in rules.values():
        if rule.op:
            _validate_rule(rule)


def _validate_rule(rule: CorrectionRule) -> None:
    if rule.op not in VALID_OPS:
        raise ValueError(f"不支持的操作: {rule.op}")
    if rule.op != "replace":
        number = _as_number(rule.value)
        if number is None:
            raise ValueError(f"数值规则必须填写数字: {rule.value}")
        if rule.op == "div" and number == 0:
            raise ValueError("除法规则不能填写 0")


def _apply_rule(value: object, rule: CorrectionRule) -> tuple[object, bool]:
    if value is None:
        return value, False
    if rule.op == "replace":
        return rule.value, False

    original = _as_number(value)
    operand = _as_number(rule.value)
    if original is None or operand is None:
        return value, True

    if rule.op == "add":
        return original + operand, False
    if rule.op == "sub":
        return original - operand, False
    if rule.op == "mul":
        return original * operand, False
    if rule.op == "div":
        return original / operand, False
    return value, True


def _as_number(value: object) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, int | float):
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def _unique_output_path(output_dir: Path, source_path: Path) -> Path:
    target = output_dir / f"{source_path.stem}-修正后{source_path.suffix}"
    if not target.exists():
        return target
    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    target = output_dir / f"{source_path.stem}-修正后-{stamp}{source_path.suffix}"
    counter = 2
    while target.exists():
        target = output_dir / f"{source_path.stem}-修正后-{stamp}-{counter}{source_path.suffix}"
        counter += 1
    return target
