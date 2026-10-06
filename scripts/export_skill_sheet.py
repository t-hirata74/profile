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


def workbook_bytes(root=ROOT):
    updated, sections, details = sources(root)
    wb = Workbook()
    wb.remove(wb.active)
    wb.properties.creator = "profile"
    wb.properties.title = "スキルシート"
    wb.properties.created = FIXED_TIME
    wb.properties.modified = FIXED_TIME

    for name, titles in [
        ("プロフィール・自己PR", SECTIONS[:2]),
        ("スキル経験", SECTIONS[2:3]),
        ("案件経歴", SECTIONS[3:]),
    ]:
        ws = wb.create_sheet(name)
        append_row(ws, "スキルシート — " + updated, heading=True)
        for title in titles:
            append_row(ws, title, heading=True)
            render(ws, sections[title])

    ws = wb.create_sheet("案件詳細")
    append_row(ws, "案件詳細 — " + updated, heading=True)
    for path in details:
        if ws.max_row > 1:
            ws.row_breaks.append(Break(id=ws.max_row))
        append_row(ws, path.relative_to(root).as_posix(), heading=True)
        render(ws, path.read_text(encoding="utf-8"))

    for ws in wb:
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 95
        ws.freeze_panes = "B2"
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = "portrait"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.print_title_rows = "1:1"
        ws.print_area = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
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
