import re
import sys
import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from pypdf import PdfReader
from reportlab.pdfgen.canvas import Canvas

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from export_skill_sheet import ROOT, SECTIONS, parse_projects, project_phase_marks, read_sections
from render_skill_sheet_pdf import (
    PAGE_HEIGHT, PAGE_WIDTH, balanced_pages,
    draw_page, paragraph, pdf_bytes,
)
import test_export_skill_sheet as excel_tests


def compact(text):
    return re.sub(r"\s+", "", text)


def projects(root):
    _, sections = read_sections(root)
    return (parse_projects(sections[SECTIONS[4]]) + parse_projects(sections[SECTIONS[5]])
            + parse_projects(sections[SECTIONS[6]], side_job=True))


class SkillSheetPDFTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.generated = pdf_bytes()

    def assert_project_content(self, reader, entries):
        groups = (balanced_pages([p for p in entries if not p.side_job])
                  + balanced_pages([p for p in entries if p.side_job]))
        self.assertEqual(len(reader.pages), 1 + len(groups))
        for page, group in zip(reader.pages[1:], groups):
            text = compact(page.extract_text())
            links = [str(item.get_object()["/A"]["/URI"])
                     for item in page.get("/Annots", [])]
            self.assertEqual(links, [p.url for p in group if p.url])
            for project in group:
                for value in [project.name, project.start, project.end, project.duration]:
                    self.assertIn(compact(value), text)
                for key, value in project.fields.items():
                    if key != "担当工程":
                        with self.subTest(project=project.name, field=key):
                            self.assertIn(compact(value), text)
            self.assertEqual(text.count("●"),
                             sum(sum(project_phase_marks(p.fields)) for p in group))

    def test_saved_and_generated_pdf_preserve_all_projects_fields_and_links(self):
        entries = projects(ROOT)
        for data in [self.generated, (ROOT / "docs/skill-sheet.pdf").read_bytes()]:
            self.assert_project_content(PdfReader(BytesIO(data)), entries)
        generated = PdfReader(BytesIO(self.generated))
        saved = PdfReader(ROOT / "docs/skill-sheet.pdf")
        self.assertEqual([compact(p.extract_text()) for p in saved.pages],
                         [compact(p.extract_text()) for p in generated.pages])

    def test_portrait_a4_embedded_japanese_font_and_page_numbers(self):
        reader = PdfReader(BytesIO(self.generated))
        for number, page in enumerate(reader.pages, 1):
            self.assertAlmostEqual(float(page.mediabox.width), PAGE_WIDTH, places=3)
            self.assertAlmostEqual(float(page.mediabox.height), PAGE_HEIGHT, places=3)
            self.assertIn(f"{number} / {len(reader.pages)}", page.extract_text())
            fonts = [font.get_object() for font in page["/Resources"]["/Font"].values()]
            self.assertTrue(any("/FontFile2" in font.get("/FontDescriptor", {})
                                for font in fonts))

    def test_ten_main_projects_and_six_side_projects_fit_five_five_six(self):
        with tempfile.TemporaryDirectory() as directory:
            root = excel_tests.ExportSkillSheetTest.fixture(directory)
            prefix = excel_tests.FIXTURE.split("## 【現案件】")[0]

            def entry(index):
                return (f"- 検証案件{index}(2025年1月〜2025年6月：6ヶ月)\n"
                        "  - 担当：エンジニア\n  - 言語：Ruby\n"
                        "  - 工程/作業：開発\n  - 担当工程：実装、単体テスト\n")

            text = (prefix + "## 【現案件】\n\n## 【過去案件: フルタイム案件】\n"
                    + "".join(entry(i) for i in range(1, 11))
                    + "\n## 【過去案件: 副業案件】\n"
                    + "".join(entry(i) for i in range(11, 17)))
            (root / "README.md").write_text(text, encoding="utf-8")
            reader = PdfReader(BytesIO(pdf_bytes(root)))
            self.assertEqual(len(reader.pages), 4)
            self.assertEqual([len(re.findall(r"検証案件\d+", p.extract_text()))
                              for p in reader.pages], [0, 5, 5, 6])
            self.assert_project_content(reader, projects(root))

    def test_unknown_fields_and_literal_markup_survive_without_extra_blank_pages(self):
        with tempfile.TemporaryDirectory() as directory:
            root = excel_tests.ExportSkillSheetTest.fixture(directory)
            path = root / "README.md"
            path.write_text(excel_tests.FIXTURE.replace(
                "  - 担当：リードエンジニア",
                "  - 補足：<検証> & 新規項目も保持\n  - 担当：リードエンジニア", 1),
                encoding="utf-8")
            reader = PdfReader(BytesIO(pdf_bytes(root)))
            self.assertEqual(len(reader.pages), 3)
            self.assertIn("<検証> & 新規項目も保持", reader.pages[1].extract_text())

    def test_overflow_is_reported_without_clipping_or_silent_shrinking(self):
        canvas = Canvas(BytesIO())
        # 全文を残すとページを超える場合はエラーにし、下端を切って出力しない。
        with self.assertRaisesRegex(ValueError, "1ページに収まりません"):
            draw_page(canvas, "検証", [paragraph("長い本文\n" * 200)], 1, 1)


if __name__ == "__main__":
    unittest.main()
