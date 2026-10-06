# スキルシートの管理手順

## 管理するファイル

| ファイル | 用途 | 更新方法 |
| --- | --- | --- |
| `README.md` | プロフィール、職務要約、スキル、案件一覧の正本 | Markdown を編集 |
| `project/**/*.md` | 案件詳細の正本 | Markdown を編集 |
| `excel/skill-sheet.xlsx` | 提出用の標準 Excel | スクリプトで再生成 |
| `scripts/export_skill_sheet.py` | 標準 Excel の書式と変換処理 | Python を編集 |

標準 Excel は「プロフィール・自己PR」「スキル経験」「案件経歴」「案件詳細」の4シートです。
README の既存の経歴セクションと `project/` の全 Markdown を出力します。
README にだけ存在する現案件・副業案件も案件経歴に含まれます。
更新日は README 先頭の更新年月を使い、経験年数や案件期間を自動計算で変更しません。
Markdown 間の既存の表記差は、そのまま各シートに出力します。

これは新規の標準形式です。以前使用した Excel や提出先指定の書式を再現するものではありません。
既存フォーマットを採用する場合は、その実ファイルを確認してから変換処理を調整します。

## 初回セットアップと更新

Python 3.10 以上で、リポジトリのルートから実行します。

```sh
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python scripts/export_skill_sheet.py
python -m unittest discover -s tests -v
python scripts/export_skill_sheet.py --check
```

Windows では仮想環境の有効化に `.venv\\Scripts\\activate` を使用します。

1. README と必要な案件詳細を修正し、README 先頭の更新年月を合わせる。
2. 生成コマンドを実行する。
3. テストと `--check` で、生成処理と Markdown / Excel の同期を確認する。
4. Excel / LibreOffice で折り返し、長文の表示、案件順、リンク、印刷プレビューを確認する。
5. Markdown と Excel を同じ PR に含め、バイナリで見えない変更点を PR 本文に記載する。

`--check` はファイルを書き換えず、再生成した内容と保存済み Excel を比較します。
同じ入力と依存関係からは同じバイト列を生成するため、内容変更のない再生成で差分が出ません。
内部の生成時刻は固定値です。経歴の更新日はシート上の README 更新年月で確認します。
README の経歴に関係しない管理案内は出力しません。
必須セクションの削除、2列以外のプロフィール表、セル文字数上限超過はエラーにします。

標準 Excel の内容を直接修正しても次回の生成で上書きされます。
訂正は Markdown、書式変更は生成スクリプトへ反映してください。
`--output` で一時ファイルへ出力できますが、提出先テンプレートへの転記機能ではありません。

## 提出先指定フォーマット

指定テンプレートを受領した場合は、標準 Excel とは別に管理します。

- 再利用する空のテンプレート: `excel/templates/<提出先>-template.xlsx`
- 提出用コピー: `output/<提出先>/skill-sheet-YYYYMMDD.xlsx`

`output/` は既存の .gitignore により Git 管理対象外です。
テンプレートは公開してよいことを確認できたものだけコミットします。
本人の住所、連絡先、顧客名などを含む提出用コピーや公開不可のテンプレートはコミットしません。

指定テンプレートはコピーしてから、最新の Markdown の情報を所定セルへ転記します。
セル配置、結合、数式、印刷設定を保持し、独自欄の情報が不足している場合は本人に確認します。
テンプレートの追加時には対応する項目マッピングと作成手順をこの文書に追記します。
この PR の標準生成スクリプトは指定テンプレートへの自動転記には対応していません。
