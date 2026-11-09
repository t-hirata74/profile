#!/usr/bin/env bash
# Markdown・Excelと同じ内容を、閲覧しやすいA4縦のPDFへ出力する。
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
python_bin="${PYTHON:-python3}"
"$python_bin" "$root/scripts/export_skill_sheet.py" --check
"$python_bin" "$root/scripts/render_skill_sheet_pdf.py" "$@"
