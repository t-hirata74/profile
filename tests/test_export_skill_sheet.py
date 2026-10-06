import importlib.util
import re
import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("exporter", ROOT / "scripts/export_skill_sheet.py")
exporter = importlib.util.module_from_spec(spec)
spec.loader.exec_module(exporter)


class ExportTest(unittest.TestCase):
    def test_all_source_lines_and_project_links_are_preserved(self):
        wb = load_workbook(BytesIO(exporter.workbook_bytes()))
        self.assertEqual(wb.sheetnames, ["プロフィール・自己PR", "スキル経験", "案件経歴", "案件詳細"])
        values = [str(c.value) for ws in wb for row in ws for c in row if c.value is not None]
        all_text = "\n".join(values)
        _, sections, details = exporter.sources(ROOT)
        texts = [sections[name] for name in exporter.SECTIONS]
        texts += [p.read_text(encoding="utf-8") for p in details]
        for text in texts:
            for line in text.splitlines():
                line = line.strip()
                if not line or line.startswith("|"):
                    continue
                line = re.sub(r"^#+\s*|^-\s*", "", line)
                line = exporter.plain(line)
                # 項目ラベルと値のセル分割を許容して、本文の欠落を検出する。
                normalized = re.sub(r"[\s:：]", "", line)
                self.assertIn(normalized, re.sub(r"[\s:：]", "", all_text))
        readme = (ROOT / "README.md").read_text(encoding="utf-8")
        links = [url for _, url in exporter.LINK.findall(readme) if "/project/" in url]
        actual = [c.hyperlink.target for row in wb["案件経歴"] for c in row if c.hyperlink]
        self.assertEqual(actual, links)
        for path in details:
            self.assertIn(path.relative_to(ROOT).as_posix(), values)
        for ws in wb:
            self.assertTrue(ws.print_area)
            self.assertEqual(ws.page_setup.fitToWidth, 1)
            self.assertEqual(ws.freeze_panes, "B2")

    def test_same_sources_are_reproducible(self):
        self.assertEqual(exporter.workbook_bytes(), exporter.workbook_bytes())

    def test_markdown_edits_change_workbook(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "project").mkdir()
            readme = (ROOT / "README.md").read_text(encoding="utf-8")
            (root / "README.md").write_text(readme, encoding="utf-8")
            detail = root / "project/test.md"
            detail.write_text("### 案件内容\n- 元の経歴\n", encoding="utf-8")
            original = exporter.workbook_bytes(root)
            detail.write_text("### 案件内容\n- 更新した経歴\n", encoding="utf-8")
            self.assertNotEqual(original, exporter.workbook_bytes(root))
            detail_original = exporter.workbook_bytes(root)
            (root / "README.md").write_text(readme.replace("38歳", "39歳"), encoding="utf-8")
            self.assertNotEqual(detail_original, exporter.workbook_bytes(root))

    def test_formula_like_text_is_not_executed(self):
        wb = Workbook()
        exporter.render(wb.active, "- =1+1\n  - 担当：=SUM(A1:A2)\n")
        buffer = BytesIO()
        wb.save(buffer)
        loaded = load_workbook(BytesIO(buffer.getvalue()))
        self.assertEqual(loaded.active["A1"].value, "=1+1")
        self.assertEqual(loaded.active["A1"].data_type, "s")
        self.assertEqual(loaded.active["B2"].data_type, "s")


if __name__ == "__main__":
    unittest.main()
