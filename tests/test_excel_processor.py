import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from cexie_excel_tool.processor import (
    CorrectionRule,
    apply_corrections_to_folder,
    discover_workbooks,
    load_rules,
    save_rules,
)


def make_workbook(path: Path, sheet_name: str = "手动测斜-18T-ZQT05-202606060806") -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(["深度（m）", "A0", "A180", "设备编号"])
    ws.append([0.5, -22.5, 23.0, "CM024"])
    ws.append([1.0, None, "bad", None])
    ws.column_dimensions["A"].width = 18
    wb.save(path)


class ExcelProcessorTests(unittest.TestCase):
    def test_discovers_xlsx_files_headers_and_point_ids(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            make_workbook(folder / "手动测斜-18T-ZQT05-原始数据.xlsx")
            (folder / "notes.txt").write_text("ignore", encoding="utf-8")

            found = discover_workbooks(folder)

            self.assertEqual(len(found), 1)
            self.assertEqual(found[0].point_id, "18T-ZQT05")
            self.assertEqual(found[0].headers, ["深度（m）", "A0", "A180", "设备编号"])

    def test_applies_numeric_and_replace_rules_without_overwriting_source(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "手动测斜-18T-ZQT05-原始数据.xlsx"
            make_workbook(source)
            rules = {
                "18T-ZQT05": {
                    "深度（m）": CorrectionRule("add", "1"),
                    "A0": CorrectionRule("mul", "2"),
                    "A180": CorrectionRule("div", "2"),
                    "设备编号": CorrectionRule("replace", "DEVICE-X"),
                }
            }

            summary = apply_corrections_to_folder(folder, rules)

            self.assertEqual(summary.processed_files, 1)
            self.assertEqual(summary.skipped_cells, 1)
            self.assertTrue(source.exists())
            self.assertEqual(len(summary.outputs), 1)
            out = summary.outputs[0]
            self.assertIn("corrected", str(out))
            self.assertIn("修正后", out.name)

            wb = load_workbook(out, data_only=True)
            ws = wb.active
            self.assertEqual(ws["A2"].value, 1.5)
            self.assertEqual(ws["B2"].value, -45.0)
            self.assertEqual(ws["C2"].value, 11.5)
            self.assertEqual(ws["D2"].value, "DEVICE-X")
            self.assertIsNone(ws["B3"].value)
            self.assertEqual(ws["C3"].value, "bad")
            self.assertEqual(ws.column_dimensions["A"].width, 18)

            original = load_workbook(source, data_only=True).active
            self.assertEqual(original["A2"].value, 0.5)
            self.assertEqual(original["D2"].value, "CM024")

    def test_generates_unique_output_name_when_file_exists(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            make_workbook(folder / "手动测斜-18T-ZQT05-原始数据.xlsx")
            corrected = folder / "corrected"
            corrected.mkdir()
            existing = corrected / "手动测斜-18T-ZQT05-原始数据-修正后.xlsx"
            existing.write_text("existing", encoding="utf-8")

            summary = apply_corrections_to_folder(folder, {"18T-ZQT05": {"A0": CorrectionRule("add", "1")}})

            self.assertEqual(summary.processed_files, 1)
            self.assertNotEqual(summary.outputs[0], existing)
            self.assertTrue(summary.outputs[0].name.startswith("手动测斜-18T-ZQT05-原始数据-修正后-"))

    def test_rejects_division_by_zero(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            make_workbook(folder / "手动测斜-18T-ZQT05-原始数据.xlsx")

            with self.assertRaises(ValueError):
                apply_corrections_to_folder(folder, {"18T-ZQT05": {"A0": CorrectionRule("div", "0")}})

    def test_saves_and_loads_rules(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "rules.json"
            rules = {"18T-ZQT05": {"A0": CorrectionRule("add", "1.2"), "设备编号": CorrectionRule("replace", "X")}}

            save_rules(path, rules)
            raw = json.loads(path.read_text(encoding="utf-8"))
            loaded = load_rules(path)

            self.assertEqual(raw["version"], 1)
            self.assertEqual(loaded["18T-ZQT05"]["A0"], CorrectionRule("add", "1.2"))
            self.assertEqual(loaded["18T-ZQT05"]["设备编号"], CorrectionRule("replace", "X"))


if __name__ == "__main__":
    unittest.main()
