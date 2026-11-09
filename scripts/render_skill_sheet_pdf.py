#!/usr/bin/env python3
"""共通のMarkdown正本とExcel生成処理から、余白を詰めたA4縦のPDFを生成する。"""

import argparse
import math
import os
from html import escape
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import Paragraph, Spacer, Table, TableStyle

from export_skill_sheet import (
    BLUE_DARK, BLUE_LIGHT, PHASE_NAMES, ROOT, SECTIONS, Project,
    parse_projects, project_phase_marks, read_sections, workbook_bytes,
)


PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN = 24
CONTENT_WIDTH = PAGE_WIDTH - MARGIN * 2
CONTENT_TOP = PAGE_HEIGHT - 50
CONTENT_BOTTOM = 32
FONT_NAME = "SkillSheetJapanese"
DARK = colors.HexColor("#" + BLUE_DARK)
LIGHT = colors.HexColor("#" + BLUE_LIGHT)


def register_font(path: Path | None = None) -> None:
    candidates = [path, os.environ.get("SKILL_SHEET_PDF_FONT"),
                  "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
                  "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            pdfmetrics.registerFont(TTFont(FONT_NAME, str(candidate)))
            pdfmetrics.registerFontFamily(FONT_NAME, normal=FONT_NAME, bold=FONT_NAME)
            return
    raise ValueError("IPAゴシックが必要です。fonts-ipafont-gothicを導入するか、"
                     "--fontで日本語TTFフォントを指定してください")


def paragraph(text: str, *, size: float = 10, color=DARK, center: bool = False,
              leading: float | None = None) -> Paragraph:
    style = ParagraphStyle(
        "compact", fontName=FONT_NAME, fontSize=size, leading=leading or size * 1.25,
        textColor=color, wordWrap="CJK", splitLongWords=True,
        alignment=TA_CENTER if center else 0, spaceBefore=0, spaceAfter=0,
        # 日本語の禁則処理で長音・小書き仮名が1字ぶら下がっても罫線を越えない。
        rightIndent=0 if center else size,
    )
    return Paragraph(escape(str(text)).replace("\n", "<br/>"), style)


def balanced_pages(projects: list[Project]) -> list[list[Project]]:
    """6件以下で均等に分ける。現在のメイン10件は5+5、副業6件は1ページ。"""
    if not projects:
        return []
    count = math.ceil(len(projects) / 6)
    base, remainder = divmod(len(projects), count)
    pages, offset = [], 0
    for index in range(count):
        length = base + (index < remainder)
        pages.append(projects[offset:offset + length])
        offset += length
    return pages


def table_style() -> list:
    return [("BOX", (0, 0), (-1, -1), 0.6, DARK),
            ("INNERGRID", (0, 0), (-1, -1), 0.35, DARK),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]


def profile_flowables(sheet) -> list:
    rows = [[paragraph(sheet[a].value or "", size=9.5, center=True),
             paragraph(sheet[b].value or "", size=9.5),
             paragraph(sheet[c].value or "", size=9.5, center=True),
             paragraph(sheet[d].value or "", size=9.5)]
            for a, b, c, d in [(f"B{row}", f"D{row}", f"G{row}", f"I{row}")
                               for row in range(3, 7)]]
    profile = Table(rows, colWidths=[44, 210, 48, CONTENT_WIDTH - 302])
    profile.setStyle(TableStyle(table_style() + [
        ("BACKGROUND", (0, 0), (0, -1), LIGHT),
        ("BACKGROUND", (2, 0), (2, -1), LIGHT),
    ]))
    flowables = [profile]
    for row in (8, 9, 10):
        flowables.extend([Spacer(1, 7), paragraph(sheet[f"B{row}"].value, size=10),
                          Spacer(1, 3)])
        body = Table([[paragraph(sheet[f"D{row}"].value or "", size=9.5)]],
                     colWidths=[CONTENT_WIDTH])
        body.setStyle(TableStyle(table_style()))
        flowables.append(body)
    return flowables


def project_table(project: Project, number: int) -> Table:
    fields = project.fields
    period = f"期間：{project.start} - {project.end}"
    if project.duration:
        period += "（" + project.duration + "）"
    title = paragraph(project.name, size=10.5, color=colors.white)
    if project.url:
        # リンク先もエスケープし、READMEの文字列をPDFマークアップとして解釈しない。
        title = Paragraph(
            f'<link href="{escape(project.url, quote=True)}" color="#FFFFFF">'
            f'{escape(project.name)}</link>', title.style,
        )
    metadata = Table([[
        paragraph(period, size=9), paragraph("担当：" + fields["担当"], size=9),
    ]], colWidths=[248, CONTENT_WIDTH - 256])
    metadata.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    rows = [[paragraph(f"No.{number}", size=10.5, center=True, color=colors.white), title],
            [metadata, ""]]

    def detail(label: str, value: str) -> None:
        if value:
            rows.append([paragraph(label, size=8.5, center=True), paragraph(value)])

    detail("担当業務", fields["工程/作業"])
    detail("開発手法", fields.get("開発手法", ""))
    detail("チーム体制", fields.get("チーム体制", ""))
    technology_keys = {"言語", "FW/ライブラリ", "DB", "インフラ", "生成AI", "アプリ"}
    environment_keys = {"エディタ/IDE", "AIエージェント", "その他", "コミュニケーション"}
    display_labels = {"FW/ライブラリ": "FW", "エディタ/IDE": "IDE", "コミュニケーション": "連絡"}
    for label, keys in (("使用技術", technology_keys), ("開発環境", environment_keys)):
        value = "　/　".join(f"{display_labels.get(key, key)}：{value}"
                            for key, value in fields.items() if key in keys and value)
        detail(label, value)
    known = technology_keys | environment_keys | {"担当", "工程/作業", "開発手法", "チーム体制", "担当工程"}
    for key, value in fields.items():
        if key not in known:
            detail(key, value)

    right_width = CONTENT_WIDTH - 54
    phases = Table([[
        paragraph(name + (" ●" if marked else ""), size=7.6, center=True)
        for name, marked in zip(PHASE_NAMES, project_phase_marks(fields))
    ]], colWidths=[(right_width - 8) / 8] * 8)
    phases.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, DARK),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1),
        ("RIGHTPADDING", (0, 0), (-1, -1), 1),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    rows.append([paragraph("担当工程", size=8.5, center=True), phases])
    table = Table(rows, colWidths=[54, right_width])
    table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.6, DARK),
        ("LINEBELOW", (0, 1), (-1, -2), 0.25, colors.HexColor("#CEDAE7")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("SPAN", (0, 1), (1, 1)),
        ("BACKGROUND", (0, 0), (-1, 0), DARK),
        ("BACKGROUND", (0, 1), (-1, 1), LIGHT),
        ("BACKGROUND", (0, 2), (0, -1), colors.HexColor("#F3F6FA")),
    ]))
    return table


def draw_page(canvas: Canvas, title: str, flowables: list, number: int, total: int) -> None:
    heights = [flowable.wrap(CONTENT_WIDTH, PAGE_HEIGHT)[1] for flowable in flowables]
    used = sum(heights)
    if used > CONTENT_TOP - CONTENT_BOTTOM:
        raise ValueError(f"{title}の内容が1ページに収まりません（{used:.1f}pt）。"
                         "文字や案件を省略せず、PDFレイアウトを調整してください")
    canvas.setFillColor(DARK)
    canvas.setFont(FONT_NAME, 13)
    canvas.drawString(MARGIN, PAGE_HEIGHT - 32, title)
    y = CONTENT_TOP
    for flowable, height in zip(flowables, heights):
        y -= height
        flowable.drawOn(canvas, MARGIN, y)
    canvas.setFont(FONT_NAME, 8)
    canvas.drawCentredString(PAGE_WIDTH / 2, 17, f"{number} / {total}")
    canvas.showPage()


def pdf_bytes(root: Path = ROOT, font: Path | None = None) -> bytes:
    register_font(font)
    _, sections = read_sections(root)
    entries = (parse_projects(sections[SECTIONS[4]]) + parse_projects(sections[SECTIONS[5]])
               + parse_projects(sections[SECTIONS[6]], side_job=True))
    # 氏名などの既定値・スキルの転記はExcel生成処理と共有する。
    workbook = load_workbook(BytesIO(workbook_bytes(root)))
    try:
        pages = [("スキルシート（エンジニア）", profile_flowables(workbook.active))]
    finally:
        workbook.close()
    number = 1
    for label, is_side in (("メイン案件", False), ("副業案件", True)):
        groups = balanced_pages([project for project in entries if project.side_job == is_side])
        for index, group in enumerate(groups, 1):
            tables = []
            first = number
            for project in group:
                if tables:
                    tables.append(Spacer(1, 5))
                tables.append(project_table(project, number))
                number += 1
            title = f"{label} {index} / {len(groups)}　（No.{first}-{number - 1}・{len(group)}件）"
            pages.append((title, tables))
    raw = BytesIO()
    canvas = Canvas(raw, pagesize=A4, invariant=1, pageCompression=1)
    canvas.setTitle("スキルシート（エンジニア）")
    canvas.setAuthor("profile")
    for index, (title, flowables) in enumerate(pages, 1):
        draw_page(canvas, title, flowables, index, len(pages))
    canvas.save()
    return raw.getvalue()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "docs/skill-sheet.pdf")
    parser.add_argument("--font", type=Path, help="埋め込む日本語TTFフォント（既定: IPAゴシック）")
    args = parser.parse_args()
    try:
        data = pdf_bytes(font=args.font)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_bytes(data)
        print(f"出力: {args.output}（A4縦）")
    except (OSError, ValueError) as error:
        parser.exit(1, f"PDF生成エラー: {error}\n")


if __name__ == "__main__":
    main()
