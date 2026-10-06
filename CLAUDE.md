# CLAUDE.md

Claude Codeは、作業前に[AGENTS.md](AGENTS.md)を読み、共通のコンテンツ規約・正本の扱い・GitHub操作方針に従う。

## リポジトリ概要

このリポジトリは日本語の職務経歴書（スキルシート）を管理する。Markdownを正本とし、指定のExcel書式へ転記するPythonスクリプトと検証用テストを含む。

## 作業の進め方

1. プロフィール・スキル・案件の要約は`README.md`、案件詳細は`project/main-job/`または`project/side-job/`で更新する。
2. READMEの更新年月、案件の期間・区分・詳細リンクを確認する。経験年数や担当工程は記載された事実に基づく。
3. `python scripts/export_skill_sheet.py`で`excel/skill-sheet.xlsx`を再生成する。
4. `python scripts/export_skill_sheet.py --check`と`python -m unittest discover -s tests -v`を実行し、Excelの表示も確認する。
5. 正本と生成Excelを同じPRに含める。

原本Excelは過去の経歴を含む書式見本として保持し、最新経歴の情報源にしない。セットアップ、項目追加、案件終了時の手順は[スキルシート管理ガイド](docs/skill-sheet-management.md)を参照する。
