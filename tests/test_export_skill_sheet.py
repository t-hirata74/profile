"""原本保持、経歴の反映、将来の更新、同期検証を確認する。"""

import hashlib
import subprocess
import sys
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from shutil import copy2, copytree

from openpyxl import load_workbook

from scripts.export_skill_sheet import workbook_bytes


ROOT = Path(__file__).resolve().parents[1]
SOURCE_NAME = "スキルシート_当てはめ済み_202605_最寄駅修正.xlsx"
SOURCE_HASH = "bc10cb1906270e4c16cd6f15cf55b494036993b61e7d0772ef5389e68f7812db"

# 本人の経歴を更新してもテストの修正が不要になるよう、入力は独立した例を使う。
FIXTURE = """2026/10 更新
# スキルシート

## 【プロフィール】
| 項目名 | 内容 |
| --- | --- |
| 氏名 | テスト 太郎 |
| フリガナ | テスト タロウ |
| エリア | 神奈川 |
| 最寄駅 | テスト駅 |
| 稼働時期 | 即日 or 相談 |
| 性別 | 男性 |
| 年齢 | 38歳 |
| 所属 | フリーランス |
| キャリア年数 | 14年目 |
| 学歴 | テスト大学卒 |
| 稼働希望 | フルリモート |

## 【職務要約/自己PR】
再生成を検証するための職務要約です。

## 【得意分野】
- API設計と開発。

## 【スキル経験】
### キャリア
- フリーランス 2017/10 ~ 現在
### 業務資格
- テスト資格
### プログラミング言語
- Ruby 8年
- Kotlin 1年未満
### FW/ライブラリ
- Ruby on Rails 8年
### DB
- MySQL 6年
### クラウド
- AWS 6年
  - Amazon Transcribe
### チーム開発、バージョン管理ツール
- Git 8年
### その他ソフトウェアツール、サービス
- Docker 6年

## 【現案件】
- 継続開発(2024年10月〜、随時更新）
  - 担当：リードエンジニア
  - 言語：Ruby、TypeScript
  - FW/ライブラリ：Ruby on Rails、React
  - DB：MySQL
  - インフラ：AWS
  - エディタ/IDE：Cursor、Orca
  - その他：Langfuse
  - 工程/作業：設計、開発、単体テスト、結合テスト、障害対応
  - 開発手法：フルリモート
  - チーム体制：エンジニア

## 【過去案件: フルタイム案件】
- [過去の開発(2024年4月〜2024年8月：5ヶ月)](https://github.com/t-hirata74/profile/blob/master/project/main-job/202404-202408.md)
  - 担当：エンジニア
  - 言語：Ruby
  - 工程/作業：基本設計、開発、結合テスト リリース作業

## 【過去案件: 副業案件】
- CRM新規開発(2025年8月〜2026年6月：11ヶ月）
  - 担当：フルスタックエンジニア
  - 言語：Ruby
  - 工程/作業：開発、コードレビュー
"""


def open_sheet(data):
    return load_workbook(BytesIO(data))


def all_text(sheet):
    return "\n".join(str(cell.value) for row in sheet for cell in row if cell.value is not None)


def project_rows(sheet):
    return [row for row in range(14, sheet.max_row + 1)
            if isinstance(sheet.cell(row, 2).value, int)]


def project_row(sheet, title):
    return next(row for row in project_rows(sheet) if title in str(sheet.cell(row, 6).value))


class ExportSkillSheetTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.original = ROOT / "excel/templates" / SOURCE_NAME
        cls.directory = tempfile.TemporaryDirectory()
        cls.root = cls.fixture(cls.directory.name)
        cls.generated = workbook_bytes(cls.root)

    @classmethod
    def tearDownClass(cls):
        cls.directory.cleanup()

    @staticmethod
    def fixture(directory):
        root = Path(directory)
        (root / "README.md").write_text(FIXTURE, encoding="utf-8")
        copytree(ROOT / "excel/templates", root / "excel/templates")
        copytree(ROOT / "project", root / "project")
        (root / 'scripts').mkdir()
        copy2(ROOT / 'scripts/export_skill_sheet.py', root / 'scripts/export_skill_sheet.py')
        return root

    def test_actual_source_bytes_and_single_sheet_format(self):
        self.assertEqual(hashlib.sha256(self.original.read_bytes()).hexdigest(), SOURCE_HASH)
        source = load_workbook(self.original)
        result = open_sheet(self.generated)
        self.assertEqual(result.sheetnames, ["スキルシート（エンジニア）"])
        sheet = result.active
        for column, dimension in source.active.column_dimensions.items():
            if dimension.min >= 12:
                continue  # 担当工程の列幅は横書きの見出しに合わせて拡張する。
            actual = sheet.column_dimensions[column]
            self.assertEqual((actual.width, actual.min, actual.max),
                             (dimension.width, dimension.min, dimension.max))
        for region in ["B2:S2", "D8:S8", "D9:S9", "D10:S10"]:
            self.assertIn(region, {str(value) for value in sheet.merged_cells.ranges})
        self.assertEqual(sheet['H12'].value, '言語')
        self.assertEqual(sheet['O13'].value, '実装')
        self.assertEqual(sheet['P13'].value.replace('\n', ''), '単体テスト')
        self.assertEqual(len(sheet.data_validations.dataValidation), 0)
        self.assertNotIn('言語ゲンゴ', all_text(sheet))
        self.assertTrue(sheet.print_area)
        source.close()
        result.close()

    def test_current_profile_all_projects_and_past_side_job(self):
        book = open_sheet(self.generated)
        sheet = book.active
        self.assertEqual(sheet['D4'].value, 'テスト 太郎')
        self.assertEqual(sheet['I5'].value, '38歳')
        self.assertEqual(sheet['I3'].value, '14年目')
        content = all_text(sheet)
        self.assertIn('2026/10 更新', content)
        self.assertNotIn('基本リモード', content)
        self.assertEqual(len(project_rows(sheet)), 3)
        rows = project_rows(sheet)
        self.assertIn('継続開発', sheet.cell(rows[0], 6).value)
        crm = project_row(sheet, 'CRM新規開発')
        side_separator = next(row for row in range(14, crm)
                              if '副業' in str(sheet.cell(row, 2).value))
        self.assertLess(side_separator, crm)
        self.assertEqual(sheet.cell(crm, 3).value, '2025/08')
        self.assertEqual(sheet.cell(crm, 5).value, '2026/06')
        self.assertEqual(sheet.cell(crm + 2, 3).value, '11ヶ月')
        for text in ['Cursor', 'Orca', 'Langfuse', 'Amazon Transcribe',
                     'Kotlin 1年未満', '稼働時期', '即日 or 相談',
                     'テスト大学卒']:
            self.assertIn(text, content)
        self.assertEqual(sum(sheet.cell(row, 6).hyperlink is not None for row in rows), 1)
        book.close()

    def test_phase_marks_do_not_inherit_unstated_experience(self):
        book = open_sheet(self.generated)
        sheet = book.active
        dx = project_row(sheet, '継続開発')
        self.assertEqual([sheet.cell(dx, col).value for col in range(12, 20)],
                         [None, None, None, '●', '●', '●', None, None])
        crm = project_row(sheet, 'CRM新規開発')
        self.assertEqual([sheet.cell(crm, col).value for col in range(12, 20)],
                         [None, None, None, '●', None, None, None, None])
        book.close()

    def test_explicit_phases_override_work_descriptions(self):
        with tempfile.TemporaryDirectory() as directory:
            root = self.fixture(directory)
            path = root / 'README.md'
            work = '要件定義、基本設計、詳細設計、開発、単体テスト、結合テスト、総合テスト、保守運用、コードレビュー'
            path.write_text(FIXTURE.replace(
                '  - 工程/作業：設計、開発、単体テスト、結合テスト、障害対応',
                f'  - 工程/作業：{work}\n  - 担当工程：実装、単体テスト', 1))
            book = open_sheet(workbook_bytes(root))
            sheet = book.active
            row = project_row(sheet, '継続開発')
            self.assertEqual([sheet.cell(row, col).value for col in range(12, 20)],
                             [None, None, None, '●', '●', None, None, None])
            self.assertIn(work, sheet.cell(row + 1, 6).value)
            book.close()
            with self.subTest('一般的な設計を明示工程として受け付けない'):
                path.write_text(path.read_text().replace('  - 担当工程：実装、単体テスト',
                                                        '  - 担当工程：設計', 1))
                with self.assertRaises(ValueError):
                    workbook_bytes(root)

    def test_explicit_empty_phases_do_not_fall_back_to_work_descriptions(self):
        with tempfile.TemporaryDirectory() as directory:
            root = self.fixture(directory)
            path = root / 'README.md'
            path.write_text(FIXTURE.replace(
                '  - 工程/作業：設計、開発、単体テスト、結合テスト、障害対応',
                '  - 工程/作業：設計、開発、単体テスト、結合テスト、障害対応\n  - 担当工程：', 1))
            book = open_sheet(workbook_bytes(root))
            sheet = book.active
            row = project_row(sheet, '継続開発')
            self.assertEqual([sheet.cell(row, col).value for col in range(12, 20)], [None] * 8)
            self.assertIn('設計、開発、単体テスト、結合テスト、障害対応', sheet.cell(row + 1, 6).value)
            book.close()

    def test_updates_unknown_fields_and_formula_like_text_are_preserved(self):
        with tempfile.TemporaryDirectory() as directory:
            root = self.fixture(directory)
            path = root / 'README.md'
            text = path.read_text()
            text = text.replace('| 年齢 | 38歳 |', '| 年齢 | 39歳 |')
            text = text.replace('| 氏名 | テスト 太郎 |', '| 氏名 | =1+1 |')
            text = text.replace('  - 担当：リードエンジニア',
                                '  - 検証用の追加項目：追加した内容を消さずに反映\n  - 担当：リードエンジニア', 1)
            path.write_text(text)
            data = workbook_bytes(root)
            self.assertNotEqual(data, self.generated)
            book = open_sheet(data)
            self.assertEqual(book.active['I5'].value, '39歳')
            self.assertEqual(book.active['D4'].value, '=1+1')
            self.assertEqual(book.active['D4'].data_type, 's')
            self.assertIn('追加した内容を消さずに反映', all_text(book.active))
            book.close()

    def test_project_count_can_grow_and_shrink_without_stale_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            root = self.fixture(directory)
            path = root / 'README.md'
            initial = path.read_text()
            extra = ('\n- 追加検証案件(2026年7月〜2026年9月：3ヶ月) ＜副業＞\n'
                     '  - 担当：エンジニア\n  - 言語：Ruby\n  - 工程/作業：開発\n')
            path.write_text(initial + extra)
            book = open_sheet(workbook_bytes(root))
            self.assertEqual(len(project_rows(book.active)), 4)
            self.assertIn('追加検証案件', all_text(book.active))
            book.close()
            path.write_text(initial.split('## 【過去案件: 副業案件】')[0]
                            + '## 【過去案件: 副業案件】\n')
            book = open_sheet(workbook_bytes(root))
            self.assertEqual(len(project_rows(book.active)), 2)
            self.assertNotIn('CRM新規開発', all_text(book.active))
            self.assertNotIn('追加検証案件', all_text(book.active))
            book.close()

    def test_missing_required_heading_fails_instead_of_omitting_content(self):
        with tempfile.TemporaryDirectory() as directory:
            root = self.fixture(directory)
            path = root / 'README.md'
            path.write_text(path.read_text().replace('## 【スキル経験】', '## 技術'))
            with self.assertRaises(ValueError):
                workbook_bytes(root)

    def test_relative_project_links_are_downloadable_workbook_links(self):
        with tempfile.TemporaryDirectory() as directory:
            root = self.fixture(directory)
            path = root / 'README.md'
            url = 'https://github.com/t-hirata74/profile/blob/master/project/main-job/202404-202408.md'
            path.write_text(path.read_text().replace(url, 'project/main-job/202404-202408.md'))
            book = open_sheet(workbook_bytes(root))
            row = project_row(book.active, '過去の開発')
            self.assertEqual(book.active.cell(row, 6).hyperlink.target, url)
            book.close()

    def test_current_side_jobs_and_empty_current_section(self):
        with tempfile.TemporaryDirectory() as directory:
            root = self.fixture(directory)
            path = root / 'README.md'
            path.write_text(FIXTURE.replace('継続開発(2024年10月〜、随時更新）',
                                            '継続開発(2024年10月〜、随時更新）＜副業＞'))
            book = open_sheet(workbook_bytes(root))
            row = project_row(book.active, '継続開発')
            self.assertEqual(book.active.cell(row, 5).value, '現在')
            self.assertTrue(any('副業' in str(book.active.cell(index, 2).value)
                                for index in range(14, row)))
            book.close()
            path.write_text(FIXTURE.split('## 【現案件】')[0] + '## 【現案件】\n\n'
                            + '## 【過去案件: フルタイム案件】'
                            + FIXTURE.split('## 【過去案件: フルタイム案件】')[1])
            book = open_sheet(workbook_bytes(root))
            self.assertEqual(len(project_rows(book.active)), 2)
            self.assertNotIn('継続開発', all_text(book.active))
            book.close()

    def test_generation_is_repeatable_and_check_detects_a_changed_workbook(self):
        self.assertEqual(workbook_bytes(self.root), self.generated)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'skill-sheet.xlsx'
            path.write_bytes(self.generated)
            command = [sys.executable, str(self.root / 'scripts/export_skill_sheet.py'),
                       '--check', '--output', str(path)]
            check = subprocess.run(command, capture_output=True, text=True)
            self.assertEqual(check.returncode, 0, check.stderr)
            book = load_workbook(path)
            book.active['D4'] = 'Excelだけを変更'
            book.save(path)
            book.close()
            before = path.read_bytes()
            check = subprocess.run(command, capture_output=True, text=True)
            self.assertNotEqual(check.returncode, 0)
            self.assertEqual(path.read_bytes(), before)
        self.assertEqual(hashlib.sha256(self.original.read_bytes()).hexdigest(), SOURCE_HASH)


if __name__ == '__main__':
    unittest.main()
