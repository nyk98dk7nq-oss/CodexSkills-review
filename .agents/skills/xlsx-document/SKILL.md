---
name: xlsx-document
description: XLSXワークブックを位置情報付きMarkdownへ変換し、限定したセル・シート編集またはレビュー結果の追記を行う。input配下の.xlsxを読んでレビューするとき、.xlsxをoutput/editedへ安全に編集するとき、結果記入済みチェックリストを作るときに使用する。.xlsの変換、CSV、オンライン表計算の操作には使用しない。
---

# XLSX文書

原本を上書きせず、`openpyxl` で `.xlsx` を読み取り・編集する。シート名、使用範囲、結合セル、セル位置、数式、コメントを失わないMarkdownを作る。

## 実行前に確認する

1. リポジトリルートで `python --version` を実行し、Python 3.12以上であることを確認する。
2. `python -c "import openpyxl"` を実行する。失敗したら処理を始めず、READMEのインストール手順を案内する。
3. 入力が `.xlsx` であることを確認する。`.xls` は先に `convert-legacy-office` を使用する。
4. 入力と出力に別の未作成パスを指定する。既存のMarkdown、編集済みXLSX、結果XLSXは上書きしない。出力パス、画像出力ディレクトリ、それらの祖先にシンボリックリンクを使用しない。

## Markdownへ変換する

次を実行する。`ROLE` は配置フォルダに合わせて `checklist`、`reference`、`target` から選ぶ。

```powershell
python .agents/skills/xlsx-document/scripts/xlsx_document.py to-markdown INPUT.xlsx OUTPUT.md --role ROLE --repo-root . --images-dir work/images
```

生成物のYAML frontmatterとセル位置を保持したままレビューに使用する。変換先は通常 `work/markdown/` とする。`--images-dir` を指定した場合は全画像の出力先を先に検証し、1件でも既存、壊れたリンク、重複、またはシンボリックリンク経由ならMarkdownを含めて何も書き込まず停止する。抽出後はシート・アンカー・相対リンクをMarkdownに残し、`image-document` でOCRする。

## 限定編集を行う

操作JSONを作り、次を実行する。

```powershell
python .agents/skills/xlsx-document/scripts/xlsx_document.py edit INPUT.xlsx OUTPUT.xlsx --operations operations.json
```

許可する操作は `set_cell`、`rename_sheet`、`add_sheet` だけとする。操作JSONの正確な形式は[操作仕様](references/operations.md)を読む。

## レビュー結果を記入する

標準結果JSONを検証した後、チェックリストのコピーを作る。

```powershell
python .agents/skills/xlsx-document/scripts/xlsx_document.py write-review INPUT.xlsx OUTPUT.xlsx --results results.json --checklist-path input/checklists/checklist.xlsx
```

対象ファイルごとに右端へ3列を追加する。判定は `適合`、`不適合`、`対象外`、`要確認` だけを渡す。各コメントには根拠と確認可能な位置を含める。

## 出力を確認する

1. Markdownの `source_path` がリポジトリ相対パスであり、6つの必須frontmatter項目があることを確認する。
2. シート、セル、数式、コメント、結合範囲が読み取れることを確認する。
3. 編集済みファイルを再度開き、指定した操作以外が変化していないことを確認する。
4. 結果記入済みチェックリストで、対象ごとの3列と全判定を確認する。
