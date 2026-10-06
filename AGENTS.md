# AGENTS.md

## リポジトリの目的と構成

職務経歴書（スキルシート）を日本語の Markdown と Excel で管理する。
経歴情報の正本は Markdown、標準 Excel はその生成物とする。

- `README.md`: プロフィール、職務要約、スキル、現案件・過去案件の一覧
- `project/main-job/`、`project/side-job/`: 案件詳細（`YYYYMM-YYYYMM.md`）
- `excel/skill-sheet.xlsx`: 提出用の標準 Excel（Git 管理する）
- `scripts/export_skill_sheet.py`: Markdown から標準 Excel を生成
- `excel/templates/skill-sheet-template.xlsx`: 個人情報・経歴を含まない指定書式の空テンプレート
- `excel/templates/template-source.json`: 基準にした実ファイルの名前・指紋・列幅・結合範囲
- `requirements.txt`: Excel 生成の依存関係
- `docs/skill-sheet-management.md`: 更新手順、提出先指定フォーマットの扱い

標準Excelの主シートは、ユーザー指定の
`スキルシート_当てはめ済み_202605_最寄駅修正.xlsx` の書式を使う。
セル配置・結合・列幅と担当工程の8列を維持し、シート内で本業と副業を区分する。
`excel/templates/skill-sheet-template.xlsx` は経歴と個人情報を除いた書式だけのテンプレート。
補助シートはスキル一覧と案件詳細の2つ。5月版の古い年齢・期間は引き継がない。
氏名・最寄駅等もREADMEを正本とし、担当工程の丸印はMarkdownに明記された内容だけに付ける。
書式改善の際もMarkdownにある経歴情報を欠落させない。
ユーザーが実ファイルを添付した場合は、その添付を基準にする。
同名のファイルや別の見本をファイル名だけで採用しない。
基準書式の変更時は `template-source.json` の出所と書式情報も更新する。

## 経歴編集のルール

- 日本語で記述し、経験年数、担当範囲、実績、期間を推測で補わない。
- README の案件一覧は新しい順に並べ、詳細がある案件にはリンクを付ける。
- 進行中の案件は `【現案件】`、終了後は `【過去案件】` に記載する。
- 担当、言語、FW/ライブラリ、DB、インフラ、その他、コミュニケーション、
  工程/作業、開発手法、チーム体制など、既存の項目形式を維持する。
- 経歴の更新時は README 先頭の更新年月を合わせ、年齢・キャリア年数も必要に応じて確認する。
- README と案件詳細に矛盾がある場合は、根拠や本人への確認により Markdown を修正する。
  Excel 生成時にどちらかを推測で書き換えない。
- 個人の住所・連絡先、顧客名、非公開情報をユーザーの指示なしに追加しない。
  このリポジトリは公開されている。

## Markdown と Excel の同期

1. README と該当する案件詳細の Markdown を更新する。
2. `python -m pip install -r requirements.txt` で依存関係を準備する（Python 3.10 以上）。
3. `python scripts/export_skill_sheet.py` で標準 Excel を再生成する。
4. `python -m unittest discover -s tests -v` と
   `python scripts/export_skill_sheet.py --check` を実行する。
5. Excel の全シートについて、文字の欠落、折り返し、リンク、印刷範囲を確認する。
   Excel / LibreOffice で目視確認できなければ、PR にその制約を記載する。
6. Markdown、生成 Excel、必要なスクリプト修正を同じ PR に含める。

標準 Excel を直接編集しない。Excel 側から内容の訂正を受けた場合は、
先に Markdown に反映してから再生成する。書式変更は生成スクリプトで行う。
生成できなかった場合は、同期済みと報告せず原因と未完了事項を明記する。
Excel はバイナリなので、PR 本文に変更内容と確認結果を日本語で記載する。

## 提出先指定の Excel

標準 Excel と提出先指定テンプレートは別物として扱う。
指定がある場合は受領したテンプレートのセル配置、結合、数式、印刷設定を維持し、
最新の Markdown を転記する。未確認の事実や空欄を推測で埋めない。
標準生成コマンドで指定テンプレートを上書きしない。
保存先、命名、公開可否は `docs/skill-sheet-management.md` に従う。

## エージェント指示の維持

共通の運用ルールはこのファイルを正本とし、Claude Code 向け入口は `CLAUDE.md` に置く。
構成・コマンド・同期方針を変える場合は、両ファイルの参照と説明も更新する。
