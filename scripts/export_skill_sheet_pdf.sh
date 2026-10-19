#!/usr/bin/env bash
# Excel を正本とし、印刷設定を維持した PDF を出力する。
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
source_file="$root/excel/skill-sheet.xlsx"
target_file="$root/docs/skill-sheet.pdf"

if [[ ! -f "$source_file" ]]; then
  echo "Excel の元ファイルが見つかりません: $source_file" >&2
  exit 1
fi
if ! command -v libreoffice >/dev/null 2>&1; then
  echo "LibreOffice が必要です（例: sudo apt-get install libreoffice-calc fonts-noto-cjk）。" >&2
  exit 1
fi

workdir="$(mktemp -d)"
trap 'rm -rf "$workdir"' EXIT

libreoffice "-env:UserInstallation=file://$workdir/lo-profile" \
  --headless --convert-to 'pdf:calc_pdf_Export' \
  --outdir "$workdir" "$source_file" >/dev/null

pdf_file="$workdir/skill-sheet.pdf"
if [[ ! -s "$pdf_file" ]]; then
  echo "PDF 変換に失敗しました。" >&2
  exit 1
fi
if command -v pdfinfo >/dev/null 2>&1; then
  pages="$(pdfinfo "$pdf_file" | awk '$1 == "Pages:" { print $2 }')"
  if [[ ! "$pages" =~ ^[0-9]+$ ]] || (( pages < 1 )); then
    echo "PDF のページ数を検証できませんでした。" >&2
    exit 1
  fi
  echo "PDF ページ数: ${pages}"
fi

mkdir -p "$(dirname "$target_file")"
cp "$pdf_file" "$target_file"
echo "出力: $target_file"
