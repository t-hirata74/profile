#!/usr/bin/env python3
"""Markdown のスキルシートを再現可能な .xlsx に出力する。"""

import argparse
import hashlib
import math
import re
from datetime import datetime
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.page import PageMargins

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "excel/skill-sheet.xlsx"
SECTIONS = ["【プロフィール】", "【職務要約/自己PR】", "【スキル経験】",
            "【現案件】", "【過去案件: フルタイム案件】", "【過去案件: 副業案件】"]
LINK = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")
FIXED_TIME = datetime(2000, 1, 1)


def plain(text):
    return LINK.sub(r"\1", text).replace("**", "").replace("`", "")


def append_row(ws, label, value=None, *, heading=False, indent=0, url=None):
    # Excel のセル上限を超える場合は切り捨てずエラーにする。
    for text in (label, value):
        if text is not None and len(text) > 32767:
            raise ValueError("Excel のセル文字数上限を超えています")
    ws.append([label, value])
    row = ws.max_row
    for cell in ws[row]:
        if isinstance(cell.value, str):
            cell.data_type = "s"  # '=' などで始まる経歴も数式として実行しない
        cell.font = Font(name="Yu Gothic", size=11, color="243746")
        cell.alignment = Alignment(vertical="top", wrap_text=True, indent=indent)
        cell.border = Border(bottom=Side(style="hair", color="D9E2EC"))
    if value is None:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    if heading:
        for cell in ws[row]:
            cell.fill = PatternFill("solid", fgColor="17365D")
            cell.font = Font(name="Yu Gothic", size=12, bold=True, color="FFFFFF")
        ws.row_dimensions[row].height = 32
    else:
        # 日本語の全角幅を考慮し、長文は複数行に分けて描画する。
        widths = (115, 88) if value is None else (25, 88)
        lines = max(sum(max(1, math.ceil(sum(2 if ord(c) > 127 else 1 for c in line) / width))
                        for line in str(text or "").split("\n"))
                    for text, width in zip((label, value), widths))
        ws.row_dimensions[row].height = min(409, max(24, lines * 17 + 8))
    if url:
        ws.cell(row, 1).hyperlink = url


def render(ws, text):
    """既存の見出し、2列表、箇条書き、段落を欠落させず出力する。"""
    paragraph = []

    def flush():
        if paragraph:
            append_row(ws, plain("\n".join(paragraph)))
            paragraph.clear()

    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            flush()
        elif stripped.startswith("#"):
            flush()
            append_row(ws, plain(stripped.lstrip("#").strip()), heading=True)
        elif stripped.startswith("|"):
            flush()
            cells = [c.strip() for c in stripped.strip("|").split("|")]
            if all(re.fullmatch(r":?-+:?", c) for c in cells):
                continue
            if len(cells) != 2:
                raise ValueError("プロフィール表は2列で記述してください")
            append_row(ws, plain(cells[0]), plain(cells[1]))
        elif stripped.startswith("- "):
            flush()
            body = stripped[2:]
            indent = min(10, (len(line) - len(line.lstrip())) // 2)
            link = LINK.search(body)
            label = plain(body)
            # 案件名の期間内の ':' は項目区切りにしない。
            field = re.match(r"^([^:：()（）]+)[:：]\s*(.*)$", label)
            if field and indent:
                append_row(ws, field[1], field[2], indent=max(0, indent - 1))
            else:
                append_row(ws, label, indent=indent, url=link[2] if link else None)
        else:
            paragraph.append(stripped)
    flush()


def sources(root):
    readme = (root / "README.md").read_text(encoding="utf-8")
    parts = re.split(r"^## (.+)\s*$", readme, flags=re.MULTILINE)
    sections = dict(zip(parts[1::2], parts[2::2]))
    missing = set(SECTIONS) - sections.keys()
    if missing:
        raise ValueError(f"README の必須セクションがありません: {sorted(missing)}")
    details = sorted((root / "project").rglob("*.md"), reverse=True)
    if not details:
        raise ValueError("project/ の案件詳細がありません")
    return readme.splitlines()[0], sections, details


def styled_cell(ws, row, column, value, *, header=False, shade=False):
    if isinstance(value, str) and len(value) > 32767:
        raise ValueError("Excel のセル文字数上限を超えています")
    cell = ws.cell(row, column, value)
    if isinstance(value, str):
        cell.data_type = "s"
    cell.font = Font(name="Yu Gothic", size=11, bold=header,
                     color="FFFFFF" if header else "243746")
    cell.fill = PatternFill("solid", fgColor="17365D" if header else "EDF3F8" if shade else "FFFFFF")
    cell.alignment = Alignment(vertical="top", wrap_text=True)
    cell.border = Border(*(Side(style="thin", color="CED9E5") for _ in range(4)))
    return cell


def merged_row(ws, text, width, *, header=False, height=26):
    if height > 409:
        raise ValueError("Excel の行高上限を超えています。長文を複数行に分割してください")
    row = ws.max_row + 1
    styled_cell(ws, row, 1, text, header=header)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=width)
    ws.row_dimensions[row].height = height
    return row


def wrapped_lines(text, width):
    return sum(max(1, math.ceil(sum(2 if ord(c) > 127 else 1 for c in line) / width))
               for line in str(text or "").split("\n"))


def overview(wb, updated, sections):
    ws = wb.create_sheet("スキルシート")
    for column, width in zip("ABCD", (17, 39, 17, 39)):
        ws.column_dimensions[column].width = width
    # 新規シートの空 A1 は意図的に使い、タイトルを先頭へ配置する。
    styled_cell(ws, 1, 1, "スキルシート", header=True)
    ws.merge_cells("A1:D1")
    ws.row_dimensions[1].height = 34
    merged_row(ws, updated, 4, height=24)
    profile = {}
    for line in sections[SECTIONS[0]].splitlines():
        if line.strip().startswith("|"):
            cells = [plain(c.strip()) for c in line.strip().strip("|").split("|")]
            if len(cells) != 2:
                raise ValueError("プロフィール表は2列で記述してください")
            if cells[0] != "項目名" and not re.fullmatch(r":?-+:?", cells[0]):
                profile[cells[0]] = cells[1]
    rows = [(("氏名", profile.get("氏名")), ("フリガナ", profile.get("フリガナ"))),
            (("最寄駅", profile.get("最寄駅")), ("学歴", profile.get("学歴")))]
    pairs = list(profile.items())
    rows += [(pairs[i], pairs[i + 1] if i + 1 < len(pairs) else ("", None))
             for i in range(0, len(pairs), 2)]
    for left, right in rows:
        row = ws.max_row + 1
        for col, (label, value) in zip((1, 3), (left, right)):
            styled_cell(ws, row, col, label, shade=True)
            styled_cell(ws, row, col + 1, value)
        ws.row_dimensions[row].height = max(30, max(wrapped_lines(p[1], 36) for p in (left, right)) * 17 + 10)
    merged_row(ws, "空欄はMarkdownに情報がない項目です。", 4, height=24)
    parts = re.split(r"^### (.+)\s*$", sections[SECTIONS[2]], flags=re.MULTILINE)
    skill_parts = dict(zip(parts[1::2], parts[2::2]))
    for title in ("キャリア", "業務資格"):
        merged_row(ws, title, 4, header=True)
        for line in skill_parts.get(title, "").splitlines():
            if line.strip():
                merged_row(ws, plain(line.strip().removeprefix("- ")), 4, height=26)
    merged_row(ws, "職務要約・自己PR", 4, header=True)
    for paragraph in sections[SECTIONS[1]].strip().split("\n\n"):
        text = plain(paragraph.strip())
        if text:
            merged_row(ws, text, 4, height=wrapped_lines(text, 103) * 17 + 12)
    merged_row(ws, "本業経歴 / 副業経歴 / スキル一覧 / 案件詳細 を別シートに掲載しています。", 4, height=28)
    return ws


def projects(text):
    result = []
    for line in text.splitlines():
        if line.startswith("- "):
            body = line[2:]
            link = LINK.search(body)
            result.append({"title": plain(body), "url": link[2] if link else None, "fields": []})
        elif line.strip().startswith("- ") and result:
            body = plain(line.strip()[2:])
            field = re.match(r"^([^:：]+)[:：]\s*(.*)$", body)
            if not field:
                raise ValueError(f"未対応の案件項目: {body}")
            result[-1]["fields"].append((field[1].strip(), field[2]))
    return result


def careers(wb, name, updated, entries):
    ws = wb.create_sheet(name)
    widths = (17, 54, 18, 16, 13, 22, 28)
    for column, width in zip("ABCDEFG", widths):
        ws.column_dimensions[column].width = width
    styled_cell(ws, 1, 1, name + " — " + updated, header=True)
    ws.merge_cells("A1:G1")
    ws.row_dimensions[1].height = 32
    merged_row(ws, "期間・経験・担当工程はMarkdownの記載を使用。詳細な取り組みは「案件詳細」を参照。", 7, height=26)
    page_height = 26
    for index, entry in enumerate(entries, 1):
        fields = dict(entry["fields"])
        title = entry["title"]
        period = re.search(r"[（(](\d{4}年.+?)[）)]", title)
        body = "担当業務\n" + fields.get("工程/作業", "記載なし")
        body += "\n\n開発手法\n" + fields.get("開発手法", "記載なし")
        role = fields.get("担当", "記載なし")
        if fields.get("チーム体制"):
            role += "\n\nチーム構成\n" + fields["チーム体制"]
        tools = []
        consumed = {"担当", "言語", "DB", "インフラ", "工程/作業", "開発手法", "チーム体制"}
        for key, value in entry["fields"]:
            if key not in consumed:
                tools.append(key + "：" + (value or "記載なし"))
        values = [period[1] if period else "記載なし", body, role,
                  fields.get("言語", "記載なし"), fields.get("DB", "記載なし"),
                  fields.get("インフラ", "") or "記載なし", "\n".join(tools)]
        # 項目内の列挙だけ改行し、各値の単語・表記は変更しない。
        for i in (2, 3, 4, 5):
            values[i] = values[i].replace("、", "、\n")
        height = max(125, max(wrapped_lines(v, w - 2) for v, w in zip(values, widths)) * 15 + 16)
        # 単一行の高さ上限による表示切れを事前に検知する。
        if height > 400:
            raise ValueError(f"案件の表示量が1行の上限を超えています。行分割が必要です: {title}")
        heading_height = max(32, wrapped_lines(title, 150) * 17 + 10)
        phase_height = max(32, wrapped_lines(fields.get("工程/作業", ""), 140) * 17 + 10)
        card_height = height + heading_height + 30 + phase_height + (22 if entry["url"] else 0) + 12
        if page_height + card_height > 490 and page_height > 26:
            ws.row_breaks.append(Break(id=ws.max_row))
            page_height = 0
        heading = merged_row(ws, f"{index:02d}  {title}", 7, header=True,
                             height=heading_height)
        row = ws.max_row + 1
        for col, label in enumerate(("期間", "業務内容", "役割・体制", "言語", "DB", "インフラ", "FW・ツール等"), 1):
            styled_cell(ws, row, col, label, header=True)
        ws.row_dimensions[row].height = 30
        row += 1
        for col, value in enumerate(values, 1):
            styled_cell(ws, row, col, value, shade=index % 2 == 0)
        ws.row_dimensions[row].height = height
        row += 1
        styled_cell(ws, row, 1, "担当工程", shade=True)
        styled_cell(ws, row, 2, fields.get("工程/作業", "記載なし"))
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
        ws.row_dimensions[row].height = phase_height
        if entry["url"]:
            link_row = merged_row(ws, "案件詳細をGitHubで開く", 7, height=22)
            ws.cell(link_row, 1).hyperlink = entry["url"]
            ws.cell(link_row, 1).font = Font(name="Yu Gothic", size=10, color="0563C1", underline="single")
        ws.append([None])
        ws.row_dimensions[ws.max_row].height = 12
        page_height += card_height
    ws.page_setup.orientation = "landscape"
    return ws


def skills(wb, updated, text):
    ws = wb.create_sheet("スキル一覧")
    for column, width in zip("ABC", (32, 45, 18)):
        ws.column_dimensions[column].width = width
    for col, value in enumerate(("分類", "技術・サービス", "経験年数"), 1):
        styled_cell(ws, 1, col, value, header=True)
    category = ""
    for line in text.splitlines():
        if line.startswith("### "):
            category = line[4:]
        elif line.strip().startswith("- "):
            value = plain(line.strip()[2:])
            experience = re.match(r"^(.+?)\s+(\d+年(?:未満)?)$", value)
            row = ws.max_row + 1
            values = (category, experience[1] if experience else value, experience[2] if experience else "")
            for col, value in enumerate(values, 1):
                styled_cell(ws, row, col, value, shade=row % 2 == 0)
            ws.row_dimensions[row].height = max(25, max(wrapped_lines(v, w - 2) for v, w in zip(values, (32, 45, 18))) * 17 + 8)
    ws.auto_filter.ref = f"A1:C{ws.max_row}"
    ws.freeze_panes = "B2"
    return ws


def workbook_bytes(root=ROOT):
    updated, sections, details = sources(root)
    wb = Workbook()
    wb.remove(wb.active)
    wb.properties.creator = "profile"
    wb.properties.title = "スキルシート"
    wb.properties.created = FIXED_TIME
    wb.properties.modified = FIXED_TIME

    overview(wb, updated, sections)
    careers(wb, "本業経歴", updated, projects(sections[SECTIONS[3]]) + projects(sections[SECTIONS[4]]))
    careers(wb, "副業経歴", updated, projects(sections[SECTIONS[5]]))
    skills(wb, updated, sections[SECTIONS[2]])

    ws = wb.create_sheet("案件詳細")
    append_row(ws, "案件詳細 — " + updated, heading=True)
    for path in details:
        if ws.max_row > 1:
            ws.row_breaks.append(Break(id=ws.max_row))
        append_row(ws, path.relative_to(root).as_posix(), heading=True)
        render(ws, path.read_text(encoding="utf-8"))

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 95
    for ws in wb:
        ws.freeze_panes = "B2"
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        if ws.title not in ("本業経歴", "副業経歴"):
            ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.print_title_rows = "1:1"
        ws.print_area = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.4, bottom=0.4, header=0.15, footer=0.15)
        ws.oddFooter.center.text = "&P / &N"

    raw = BytesIO()
    wb.save(raw)
    # ZIP の時刻と openpyxl が更新する modified を固定し、再生成の差分を安定させる。
    result = BytesIO()
    with ZipFile(raw) as original, ZipFile(result, "w", ZIP_DEFLATED) as normalized:
        for name in sorted(original.namelist()):
            data = original.read(name)
            if name == "docProps/core.xml":
                data = re.sub(rb"(<dcterms:modified[^>]*>)[^<]+", rb"\g<1>2000-01-01T00:00:00Z", data)
            info = ZipInfo(name, (2000, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            normalized.writestr(info, data)
    return result.getvalue()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=OUTPUT)
    parser.add_argument("--check", action="store_true", help="保存済み Excel が最新か検証する")
    args = parser.parse_args()
    data = workbook_bytes()
    if args.check:
        if not args.output.is_file() or args.output.read_bytes() != data:
            parser.exit(1, "Excel が未生成または古い状態です。python scripts/export_skill_sheet.py を実行してください。\n")
        print("Markdown と Excel は同期しています。")
    else:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_bytes(data)
        print(f"出力: {args.output} (SHA256: {hashlib.sha256(data).hexdigest()})")


if __name__ == "__main__":
    main()
