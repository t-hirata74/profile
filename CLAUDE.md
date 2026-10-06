# CLAUDE.md

このリポジトリは職務経歴書（スキルシート）を Markdown と Excel で管理します。
共通の編集・同期・公開情報のルールは、以下のファイルを読み込んで遵守してください。

@AGENTS.md

標準Excelは指定された5月版の「スキルシート（エンジニア）」を主シートとし、
スキル一覧と案件詳細を補助シートに付けます。
書式・項目対応と未記載情報の扱いは `docs/skill-sheet-management.md` を参照してください。
基準となる実ファイルと書式は `excel/templates/template-source.json` に記録しています。
添付がある場合は最新の添付ファイルを確認し、同名の別ファイルに置き換えないでください。

## Claude Code の作業手順

- 経歴の正本は `README.md` と `project/` 配下の Markdown です。
- Markdown を更新したら `python scripts/export_skill_sheet.py` を実行し、
  `excel/skill-sheet.xlsx` も同じ PR に含めます。
- 依存関係は `python -m pip install -r requirements.txt` で導入します。
- 検証は `python -m unittest discover -s tests -v` と
  `python scripts/export_skill_sheet.py --check` で行います。
- 標準 Excel の直接編集や、提出先指定テンプレートの上書きはしません。
- 指定フォーマットの受領時や提出用ファイルの作成時は
  `docs/skill-sheet-management.md` を参照します。
- PR 本文には経歴・書式の変更内容と検証結果を日本語で記載します。
