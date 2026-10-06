#!/usr/bin/env python3
"""README を正本に、指定された Excel 書式でスキルシートを再生成する。"""

import argparse
import hashlib
import math
import os
import re
import unicodedata
from copy import copy
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from urllib.parse import quote
from xml.etree.ElementTree import canonicalize
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile, ZipInfo

# オプション依存の有無で出力が変わらないよう、標準の XML 実装に統一する。
# 呼出し元が openpyxl を先に import 済みの場合も、末尾の XML 正規化で吸収する。
os.environ["OPENPYXL_LXML"] = "False"

from openpyxl import load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE, MergedCell
from openpyxl.packaging.core import DocumentProperties
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.datavalidation import DataValidationList
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break, RowBreak


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "excel/skill-sheet.xlsx"
TEMPLATE = Path("excel/templates/スキルシート_当てはめ済み_202605_最寄駅修正.xlsx")
SHEET_NAME = "スキルシート（エンジニア）"
PROJECT_BASE_URL = "https://github.com/t-hirata74/profile/blob/master/"
SECTIONS = (
    "【プロフィール】", "【職務要約/自己PR】", "【得意分野】", "【スキル経験】",
    "【現案件】", "【過去案件: フルタイム案件】", "【過去案件: 副業案件】",
)
PROFILE_CELLS = {
    "D3": "フリガナ", "D4": "氏名", "I3": "キャリア年数",
    "I4": "性別", "I5": "年齢", "I6": "学歴",
}
TECH_FIELDS = {"FW/ライブラリ", "エディタ/IDE", "AIエージェント", "生成AI",
               "その他", "コミュニケーション", "アプリ"}
PHASE_NAMES = ("要件定義", "基本設計", "詳細設計", "実装",
               "単体テスト", "結合テスト", "総合テスト", "保守運用")
PHASE_HEADERS = ("要件\n定義", "基本\n設計", "詳細\n設計", "実装",
                 "単体\nテスト", "結合\nテスト", "総合\nテスト", "保守\n運用")
BLUE_DARK = "17365D"
BLUE_LIGHT = "E9F1F8"
WHITE = "FFFFFF"
BLACK = "000000"
FIXED_TIME = datetime(2000, 1, 1)
XSI_TYPE = "{http://www.w3.org/2001/XMLSchema-instance}type"


@dataclass
class Project:
    title: str
    name: str
    url: str | None
    start: str
    end: str
    duration: str
    fields: dict[str, str]
    side_job: bool


def plain(text: str) -> str:
    """表記内容を保持したまま、表示に不要な Markdown 装飾を取り除く。"""
    text = re.sub(r"\[([^\]]+)\]\(([^\s)]+)\)", r"\1", text)
    return text.replace("**", "").replace("`", "").replace("\\|", "|")


def read_sections(root: Path) -> tuple[str, dict[str, str]]:
    text = (root / "README.md").read_text(encoding="utf-8")
    first_line = text.splitlines()[0] if text else ""
    if not re.fullmatch(r"\d{4}/\d{1,2}\s+更新", first_line):
        raise ValueError("README の先頭に更新日を YYYY/M 更新 の形式で記載してください")
    parts = re.split(r"^##[ \t]+(.+?)[ \t]*$", text, flags=re.MULTILINE)
    sections: dict[str, str] = {}
    for title, content in zip(parts[1::2], parts[2::2]):
        if title in sections:
            raise ValueError(f"README のセクションが重複しています: {title}")
        if title not in SECTIONS:
            raise ValueError(f"未対応の README セクションです: {title}")
        sections[title] = content.strip()
    missing = [title for title in SECTIONS if title not in sections]
    if missing:
        raise ValueError("README の必須セクションがありません: " + "、".join(missing))
    for title in SECTIONS[:4]:
        if not sections[title]:
            raise ValueError(f"README の必須セクションが空です: {title}")
    return first_line, sections


def parse_profile(text: str) -> dict[str, str]:
    profile: dict[str, str] = {}
    header_seen = False
    for line in text.splitlines():
        if not line.strip():
            continue
        if not line.strip().startswith("|") or not line.strip().endswith("|"):
            raise ValueError(f"プロフィールは2列表で記載してください: {line.strip()}")
        cells = [plain(value.strip()) for value in re.split(r"(?<!\\)\|", line.strip()[1:-1])]
        if len(cells) != 2:
            raise ValueError("プロフィール表は2列で記載してください")
        if cells == ["項目名", "内容"]:
            header_seen = True
            continue
        if all(re.fullmatch(r":?-+:?", value) for value in cells):
            continue
        key, value = cells
        if not key or key in profile:
            raise ValueError(f"プロフィールの項目名が空または重複しています: {key}")
        profile[key] = value
    required = set(PROFILE_CELLS.values()) | {"最寄駅", "稼働希望"}
    missing = sorted(key for key in required if not profile.get(key))
    if not header_seen or missing:
        raise ValueError("プロフィールの必須項目を確認してください: " + "、".join(missing))
    return profile


def parse_skills(text: str) -> dict[str, list[str]]:
    """全カテゴリを抽出する。AWS 等の子項目は所属を保ったまま括弧内にまとめる。"""
    categories: dict[str, list[str]] = {}
    category: str | None = None
    item: str | None = None
    children: list[str] = []

    def flush() -> None:
        nonlocal item
        if item is not None and category is not None:
            suffix = "（" + "、".join(children) + "）" if children else ""
            categories[category].append(item + suffix)
        item = None
        children.clear()

    for line in text.splitlines():
        if not line.strip():
            continue
        if line.startswith("### "):
            flush()
            category = plain(line[4:].strip())
            if not category or category in categories:
                raise ValueError(f"スキル経験のカテゴリ名が空または重複しています: {category}")
            categories[category] = []
        elif line.startswith("- ") and category is not None:
            flush()
            item = plain(line[2:].strip())
            if not item:
                raise ValueError(f"スキル経験の項目が空です: {category}")
        elif line.startswith("  - ") and item is not None:
            value = plain(line[4:].strip())
            if not value:
                raise ValueError(f"スキル経験の子項目が空です: {category}")
            children.append(value)
        else:
            raise ValueError(f"未対応のスキル経験の記述です: {line.strip()}")
    flush()
    if not categories.get("業務資格"):
        raise ValueError("スキル経験に ### 業務資格 と資格の箇条書きが必要です")
    if any(not items for items in categories.values()):
        raise ValueError("スキル経験に内容が空のカテゴリがあります")
    return categories


def parse_period(title: str) -> tuple[str, str, str, str]:
    match = re.search(r"[（(](\d{4}年\d{1,2}月[^()（）]*)[）)]", title)
    if not match:
        raise ValueError(f"案件名に参画期間がありません: {title}")
    period = re.fullmatch(
        r"(?P<start>\d{4}年\d{1,2}月)\s*[〜～~]\s*"
        r"(?P<end>\d{4}年\d{1,2}月|現在)?(?P<rest>.*)", match[1]
    )
    if not period:
        raise ValueError(f"未対応の案件期間です: {match[1]}")

    def date(value: str) -> str:
        found = re.fullmatch(r"(\d{4})年(\d{1,2})月", value)
        if not found or not 1 <= int(found[2]) <= 12:
            raise ValueError(f"案件期間の年月が不正です: {value}")
        return f"{int(found[1]):04}/{int(found[2]):02}"

    start = date(period["start"])
    end = date(period["end"]) if period["end"] not in (None, "現在") else "現在"
    rest = period["rest"].strip()
    duration = ""
    if rest.startswith(("：", ":")):
        duration = rest[1:].strip()
        if not re.fullmatch(r"[０-９0-9]+[ヶかケ箇]月", duration):
            raise ValueError(f"未対応の案件月数です: {rest}")
    elif rest.startswith(("、", ",", "，")):
        duration = rest[1:].strip()
    elif rest:
        raise ValueError(f"未対応の案件期間の補足です: {rest}")
    # 記載された月数だけを使う。期間の原文は業務内容にも残す。
    name = (title[:match.start()] + title[match.end():]).strip()
    if not name:
        raise ValueError("案件名が空です")
    return name, start, end, duration


def parse_projects(text: str, *, side_job: bool = False) -> list[Project]:
    projects: list[Project] = []
    for line in text.splitlines():
        if not line.strip():
            continue
        if line.startswith("- "):
            raw = line[2:].strip()
            link = re.fullmatch(r"\[(.+)\]\((\S+)\)", raw)
            if raw.startswith("[") and not link:
                raise ValueError(f"案件リンクの形式を確認してください: {raw}")
            title, url = (plain(link[1]), link[2]) if link else (plain(raw), None)
            if url and not (url.lower().startswith(("https://", "http://")) or url.startswith("project/")):
                raise ValueError(f"未対応の案件リンクです: {url}")
            if url and url.startswith("project/"):
                # README 基準のリンクを、Excel の保存場所に依存しない URL にする。
                # 日本語・裸の % はエンコードし、既存の %HH は二重にエンコードしない。
                relative = re.sub(r"%(?![0-9A-Fa-f]{2})", "%25", url)
                url = PROJECT_BASE_URL + quote(relative, safe="/%?#=&")
            name, start, end, duration = parse_period(title)
            projects.append(Project(title, name, url, start, end, duration, {},
                                    side_job or bool(re.search(r"[＜<]副業[＞>]", title))))
        elif line.startswith("  - ") and projects:
            field = re.fullmatch(r"([^:：]+)[:：]\s*(.*)", plain(line[4:].strip()))
            if not field:
                raise ValueError(f"未対応の案件項目です: {line.strip()}")
            key, value = field[1].strip(), field[2]
            if not key or key in projects[-1].fields:
                raise ValueError(f"案件項目名が空または重複しています: {key}")
            projects[-1].fields[key] = value
        else:
            raise ValueError(f"未対応の案件の記述です: {line.strip()}")
    for project in projects:
        missing = [key for key in ("担当", "言語", "工程/作業") if not project.fields.get(key)]
        if missing:
            raise ValueError(f"案件の必須項目がありません（{project.name}）: " + "、".join(missing))
    return projects


def put(ws, address: str, value: str | int | None):
    if isinstance(value, str):
        if len(value) > 32767:
            raise ValueError(f"Excel のセル文字数上限を超えています: {address}")
        if ILLEGAL_CHARACTERS_RE.search(value):
            raise ValueError(f"Excel に保存できない制御文字があります: {address}")
    cell = ws[address]
    cell.value = value
    if isinstance(value, str):
        cell.data_type = "s"  # '=' 等で始まる入力も数式として保存しない。
    return cell


def column_points(ws, first: int, last: int) -> float:
    """元書式の I:J、L:S のグループ指定を含めて列幅を読む。変更はしない。"""
    total = 0.0
    for column in range(first, last + 1):
        width = ws.sheet_format.defaultColWidth or 9
        for dimension in ws.column_dimensions.values():
            if dimension.min <= column <= dimension.max:
                width = dimension.width
        total += math.floor(width * 7 + 5) * 0.75
    return total


def text_height(ws, value: str, first: int, last: int, size: float = 11) -> float:
    width = max(10.0, (column_points(ws, first, last) - 10) * 0.93)
    line_count = 0
    for line in value.split("\n"):
        length = sum(size if unicodedata.east_asian_width(char) in "WFA" else size * 0.58
                     for char in line)
        line_count += max(1, math.ceil(length / width))
    return max(24.0, math.ceil(line_count * size * 1.45 + 9))


def row_height(ws, row: int, height: float) -> None:
    if height > 409:
        raise ValueError(f"Excel の行高上限を超えています（{row}行目）。README の記述量を確認してください")
    ws.row_dimensions[row].height = height


def phase_marks(text: str) -> list[bool]:
    """工程名が独立した項目として明記されている場合だけ印を付ける。"""
    tokens = re.split(r"[、,，\s]+", text)

    def stated(*aliases: str) -> bool:
        for token in tokens:
            for alias in aliases:
                if token == alias:
                    return True
                note = re.fullmatch(re.escape(alias) + r"[（(](.+)[）)]", token)
                if note and not re.search(r"未担当|未実施|未経験|対象外|なし|除く", note[1]):
                    return True
        return False

    return [stated("要件定義"), stated("基本設計"), stated("詳細設計"), stated("開発", "実装"),
            stated("単体テスト"), stated("結合テスト"), stated("総合テスト"),
            stated("保守運用", "保守・運用", "保守/運用")]


def project_phase_marks(fields: dict[str, str]) -> list[bool]:
    """明示した担当工程を優先し、項目自体がない案件だけ従来の作業説明から判定する。"""
    if "担当工程" not in fields:
        return phase_marks(fields["工程/作業"])
    text = fields["担当工程"].strip()
    if not text:
        return [False] * len(PHASE_NAMES)
    names = {name for name in re.split(r"[、,，\s]+", text) if name}
    if not names:
        raise ValueError("担当工程が区切り文字だけです。工程名を記載するか空欄にしてください")
    unknown = sorted(names - set(PHASE_NAMES))
    if unknown:
        raise ValueError("担当工程に未対応の工程名があります: " + "、".join(unknown)
                         + "（指定可能: " + "、".join(PHASE_NAMES) + "）")
    return [name in names for name in PHASE_NAMES]


def apply_blue_theme(ws) -> None:
    """罫線の形状・文字サイズ・レイアウトを保ち、生成シートだけを青系に統一する。"""
    white_fill = PatternFill("solid", fgColor=WHITE)

    def base_style(target, *, marked: bool = False) -> None:
        target.fill = white_fill
        font = copy(target.font)
        font.color = BLACK if marked else BLUE_DARK
        target.font = font
        border = copy(target.border)
        for edge in ("left", "right", "top", "bottom", "diagonal", "vertical", "horizontal", "start", "end"):
            side = getattr(border, edge)
            if side is not None and (side.style is not None or side.color is not None):
                side = copy(side)
                side.color = BLUE_DARK
                setattr(border, edge, side)
        target.border = border

    # 列・行の既定書式に含まれる緑色も置き換える。幅・高さ等は変更しない。
    for dimension in list(ws.column_dimensions.values()) + list(ws.row_dimensions.values()):
        base_style(dimension)
    for cells in ws:
        for cell in cells:
            base_style(cell, marked=cell.value == "●")

    def shade(region: str, fill: str, text: str | None = None) -> None:
        for cells in ws[region]:
            for cell in cells:
                cell.fill = PatternFill("solid", fgColor=fill)
                if text is not None:
                    font = copy(cell.font)
                    font.color = text
                    cell.font = font

    shade("B2:S2", BLUE_DARK, WHITE)
    for region in ("B3:C6", "G3:H6", "B8:C10"):
        shade(region, BLUE_LIGHT, BLUE_DARK)
    shade("B12:K13", BLUE_DARK, WHITE)
    shade("L12:S12", BLUE_DARK, WHITE)
    shade("L13:S13", BLUE_LIGHT, BLUE_DARK)
    for row in range(14, ws.max_row + 1):
        if ws.cell(row, 2).value == "副業案件":
            shade(f"B{row}:S{row}", BLUE_DARK, WHITE)
        elif isinstance(ws.cell(row, 2).value, int):
            shade(f"B{row}:B{row + 2}", BLUE_DARK, WHITE)
            shade(f"F{row}:G{row}", BLUE_LIGHT, BLUE_DARK)
            for column in "LMNOPQRS":
                if ws[f"{column}{row}"].value == "●":
                    shade(f"{column}{row}:{column}{row + 2}", BLUE_LIGHT)
    ws.sheet_properties.tabColor = BLUE_DARK


def prepare_template(wb):
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"指定のテンプレートに {SHEET_NAME} シートがありません")
    ws = wb[SHEET_NAME]
    for sheet in list(wb):
        if sheet != ws:
            wb.remove(sheet)
    required = {"D3:F3", "D4:F4", "D5:F5", "D6:F6", "D8:S8", "D9:S9", "D10:S10",
                "B12:E13", "F12:F13", "F15:F16", "G15:G16", "B14:B16", "H14:H16"}
    if not required.issubset({str(value) for value in ws.merged_cells.ranges}):
        raise ValueError("元 Excel のセル結合が想定と異なります。指定された原本を使用してください")
    styles = {(offset, column): copy(ws.cell(14 + offset, column)._style)
              for offset in range(3) for column in range(2, 20)}
    merges = sorted(
        [(value.min_row - 14, value.max_row - 14, value.min_col, value.max_col)
         for value in ws.merged_cells.ranges if value.min_row >= 14 and value.max_row <= 16]
    )
    labels = {"B2", "B3", "G3", "B4", "G4", "B5", "G5", "B6", "G6", "B8", "B9", "B10",
              "B12", "F12", "G12", "H12", "I12", "J12", "K12", "L12"}
    labels.update(f"{column}13" for column in "LMNOPQRS")
    # 元ファイルは記入済み。プロフィール・案件文・人数・古い●も含め、入力値を全消去する。
    for cells in ws:
        for cell in cells:
            if not isinstance(cell, MergedCell):
                if cell.coordinate not in labels:
                    cell.value = None
                cell.hyperlink = None
                cell.comment = None
    for merged in list(ws.merged_cells.ranges):
        if merged.min_row >= 14:
            ws.unmerge_cells(str(merged))
    ws.delete_rows(14, max(1, ws.max_row - 13))
    for row in list(ws.row_dimensions):
        if row >= 14:
            del ws.row_dimensions[row]
    ws.row_breaks = RowBreak()
    # 提出用シートに入力用プルダウンを残さず、選択中も丸印を隠さない。
    ws.data_validations = DataValidationList()
    for address, label in {"H12": "言語", "J12": "機種\n・\nOS", "O13": "実装",
                           "P13": "単体テスト"}.items():
        put(ws, address, label)
    return ws, styles, merges


def style_phase_header(ws) -> None:
    """担当工程を回転なしで読める列幅と2行見出しにする。"""
    for index, (column, label) in enumerate(zip("LMNOPQRS", PHASE_HEADERS), 12):
        # 原本のL:Sは1つの列幅グループなので、各列の範囲を明示する。
        dimension = ws.column_dimensions[column]
        dimension.min = index
        dimension.max = index
        dimension.width = 8.5
        ws.cell(12, index).fill = PatternFill("solid", fgColor=BLUE_DARK)
        cell = put(ws, f"{column}13", label)
        cell.font = Font(name="メイリオ", size=11, bold=True, color=BLUE_DARK)
        cell.fill = PatternFill("solid", fgColor=BLUE_LIGHT)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True,
                                   textRotation=0, shrink_to_fit=False)
    ws["L12"].font = Font(name="メイリオ", size=12, bold=True, color=WHITE)
    ws["L12"].alignment = Alignment(horizontal="center", vertical="center")


def engineer_sheet(wb, updated: str, sections: dict[str, str]) -> None:
    profile = parse_profile(sections[SECTIONS[0]])
    skills = parse_skills(sections[SECTIONS[3]])
    entries = (parse_projects(sections[SECTIONS[4]])
               + parse_projects(sections[SECTIONS[5]])
               + parse_projects(sections[SECTIONS[6]], side_job=True))
    if not entries:
        raise ValueError("README に案件がありません")
    projects = [entry for entry in entries if not entry.side_job] + [entry for entry in entries if entry.side_job]
    ws, styles, merges = prepare_template(wb)
    style_phase_header(ws)
    ws.merge_cells("M1:S1")
    put(ws, "M1", updated)
    ws["M1"].font = Font(name="メイリオ", size=11)
    ws["M1"].alignment = Alignment(horizontal="right", vertical="center")
    for address, key in PROFILE_CELLS.items():
        put(ws, address, profile[key])
    career = skills.get("キャリア", [])
    if career:
        put(ws, "I3", profile["キャリア年数"] + "：" + "、".join(career))
    put(ws, "D5", profile["最寄駅"] + " / " + profile["稼働希望"])
    put(ws, "D6", "\n".join(skills["業務資格"]))
    put(ws, "D8", plain(sections[SECTIONS[1]]))
    put(ws, "D9", plain(sections[SECTIONS[2]]))
    put(ws, "D10", "\n".join("・" + category + "：" + "、".join(items)
                              for category, items in skills.items()
                              if category not in {"業務資格", "キャリア"}))
    for row in (3, 4, 5, 6):
        row_height(ws, row, max(text_height(ws, ws[f"D{row}"].value, 4, 6),
                               text_height(ws, ws[f"I{row}"].value, 9, 19)))
    for row in (8, 9, 10):
        row_height(ws, row, text_height(ws, ws[f"D{row}"].value, 4, 19))
    for row, height in ((1, 24), (7, 10), (11, 10), (12, 26), (13, 46)):
        row_height(ws, row, height)

    # A3 横で幅を1ページに収める。3行の案件ブロックをページの途中で切らない。
    scale = min(1.0, (1190.55 - 36) / column_points(ws, 2, 19))
    page_height = (841.89 - 50.4) / scale
    project_capacity = page_height - ws.row_dimensions[12].height - ws.row_dimensions[13].height
    ws.row_breaks.append(Break(id=11))
    row, used_height, in_side = 14, 0.0, False
    for index, project in enumerate(projects, 1):
        if project.side_job and not in_side:
            in_side = True
            if used_height:
                ws.row_breaks.append(Break(id=row - 1))
            for column in range(2, 20):
                ws.cell(row, column)._style = copy(styles[(0, column)])
                ws.cell(row, column).fill = PatternFill("solid", fgColor=BLUE_LIGHT)
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=19)
            put(ws, f"B{row}", "副業案件")
            ws[f"B{row}"].font = Font(name="メイリオ", size=11, bold=True, color=BLUE_DARK)
            row_height(ws, row, 28)
            row += 1
            used_height = 28
        for (offset, column), style in styles.items():
            ws.cell(row + offset, column)._style = copy(style)
        for first, last, left, right in merges:
            ws.merge_cells(start_row=row + first, end_row=row + last,
                           start_column=left, end_column=right)
        fields = project.fields
        consumed = {"担当", "言語", "DB", "インフラ", "チーム体制", "担当工程"} | TECH_FIELDS
        paragraphs = ["≪プロジェクト内容≫\n" + project.title]
        # 新規の案件項目も省略せず、業務内容欄に追加する。
        for key, value in fields.items():
            if key not in consumed:
                label = "担当業務" if key == "工程/作業" else key
                paragraphs.append(f"≪{label}≫\n{value}")
        body = "\n\n".join(paragraphs)
        tools = "\n".join(f"{key}：{value}" for key, value in fields.items() if key in TECH_FIELDS)
        team = "チーム体制\n" + fields["チーム体制"] if fields.get("チーム体制") else ""
        values = {"B": index, "C": project.start, "D": "−", "E": project.end,
                  "F": "■" + project.name, "G": fields["担当"], "H": fields["言語"],
                  "I": fields.get("DB", ""), "J": fields.get("インフラ", ""), "K": tools}
        for column, value in values.items():
            put(ws, f"{column}{row}", value)
        put(ws, f"F{row + 1}", body)
        put(ws, f"G{row + 1}", team)
        put(ws, f"C{row + 2}", project.duration)
        for column, marked in zip("LMNOPQRS", project_phase_marks(fields)):
            cell = put(ws, f"{column}{row}", "●" if marked else None)
            cell.font = Font(name="メイリオ", size=18, bold=True, color=BLACK)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            fill = PatternFill("solid", fgColor=BLUE_LIGHT if marked else WHITE)
            for offset in range(3):
                ws[f"{column}{row + offset}"].fill = fill
        for column in "FG":
            cell = ws[f"{column}{row}"]
            cell.fill = PatternFill("solid", fgColor=BLUE_LIGHT)
            cell.font = Font(name="メイリオ", size=11, bold=True, color=BLUE_DARK)
        if project.url:
            ws[f"F{row}"].hyperlink = project.url
        title_height = max(text_height(ws, values["F"], 6, 6), text_height(ws, fields["担当"], 7, 7))
        content_height = max(text_height(ws, body, 6, 6), text_height(ws, team, 7, 7), 48)
        technical_height = max(text_height(ws, str(values[column]), number, number)
                               for column, number in (("H", 8), ("I", 9), ("J", 10), ("K", 11)))
        total = max(title_height + content_height, technical_height)
        if total > project_capacity:
            raise ValueError(f"案件が印刷1ページの高さを超えています。README の記述量を確認してください: {project.name}")
        if used_height and used_height + total > project_capacity:
            ws.row_breaks.append(Break(id=row - 1))
            used_height = 0.0
        bottom = max(24, total - title_height - 409)
        row_height(ws, row, title_height)
        row_height(ws, row + 1, total - title_height - bottom)
        row_height(ws, row + 2, bottom)
        used_height += total
        row += 3

    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=19)
    put(ws, f"B{row}", "●：担当工程。空欄：未記載（経験なしを意味しません）。案件名のリンクから詳細を確認できます。")
    ws[f"B{row}"].font = Font(name="メイリオ", size=10, color=BLUE_DARK)
    row_height(ws, row, 28)
    if used_height + 28 > project_capacity:
        ws.row_breaks.append(Break(id=row - 1))
    for cells in ws:
        for cell in cells:
            if isinstance(cell, MergedCell):
                continue
            if isinstance(cell.value, str):
                cell.data_type = "s"
            if cell.row not in (1, 12, 13) and not (cell.row >= 14 and cell.column >= 12):
                alignment = copy(cell.alignment)
                alignment.wrap_text = True
                alignment.shrink_to_fit = False
                if cell.column in (4, 6, 8, 9, 10, 11):
                    alignment.vertical = "top"
                cell.alignment = alignment
    ws.freeze_panes = "F3"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.35, bottom=0.35, header=0.15, footer=0.15)
    ws.print_title_rows = "12:13"
    ws.print_area = f"B1:S{row}"
    ws.oddFooter.center.text = "&P / &N"
    apply_blue_theme(ws)


def normalize_archive(raw: BytesIO) -> bytes:
    result = BytesIO()
    with ZipFile(raw) as original, ZipFile(result, "w", ZIP_DEFLATED, compresslevel=9) as archive:
        for name in sorted(original.namelist()):
            data = original.read(name)
            if name == "docProps/core.xml":
                data = re.sub(rb"(<dcterms:modified[^>]*>)[^<]+", rb"\g<1>2000-01-01T00:00:00Z", data)
            if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                # openpyxl の結合範囲は set。PYTHONHASHSEED による順序差を取り除く。
                def sort_merges(match):
                    cells = re.findall(rb'<mergeCell ref="([^"]+)"\s*/>', match[1])
                    cells.sort(key=lambda ref: range_boundaries(ref.decode()))
                    return (b'<mergeCells count="' + str(len(cells)).encode() + b'">'
                            + b"".join(b'<mergeCell ref="' + ref + b'"/>' for ref in cells)
                            + b"</mergeCells>")

                data = re.sub(rb'<mergeCells\b[^>]*>(.*?)</mergeCells>', sort_merges, data)
            if name.endswith((".xml", ".rels")):
                data = canonicalize(data, qname_aware_attrs={XSI_TYPE}).encode("utf-8")
            info = ZipInfo(name, (2000, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o600 << 16
            archive.writestr(info, data, compresslevel=9)
    return result.getvalue()


def workbook_bytes(root: Path = ROOT) -> bytes:
    root = Path(root)
    updated, sections = read_sections(root)
    wb = load_workbook(root / TEMPLATE)
    try:
        wb.properties = DocumentProperties(creator="profile", lastModifiedBy="profile", title="スキルシート",
                                           language="ja-JP", created=FIXED_TIME, modified=FIXED_TIME)
        engineer_sheet(wb, updated, sections)
        raw = BytesIO()
        wb.save(raw)
        return normalize_archive(raw)
    finally:
        wb.close()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=OUTPUT, help="出力先の .xlsx ファイル")
    parser.add_argument("--check", action="store_true", help="保存済み Excel が README と同期しているか検証する")
    args = parser.parse_args()
    try:
        data = workbook_bytes()
        if args.check:
            if not args.output.is_file() or args.output.read_bytes() != data:
                parser.exit(1, "Excel が未生成または古い状態です。python scripts/export_skill_sheet.py を実行してください。\n")
            print("README と Excel は同期しています。")
        else:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_bytes(data)
            print(f"出力: {args.output} (SHA256: {hashlib.sha256(data).hexdigest()})")
    except (OSError, ValueError, BadZipFile) as error:
        parser.exit(1, f"Excel 生成エラー: {error}\n")


if __name__ == "__main__":
    main()
