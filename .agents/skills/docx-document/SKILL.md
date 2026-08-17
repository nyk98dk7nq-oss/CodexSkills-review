---
name: docx-document
description: DOCX文書の段落・リスト・ハイパーリンク・表・ヘッダー・フッター・埋め込み画像を位置情報付きMarkdownへ変換し、限定したテキスト編集を行う。input配下の.docxをレビューするとき、画像を抽出するとき、利用者が依頼した.docx編集をoutput/editedへ保存するときに使用する。.docの変換、PDF、文書レイアウトの厳密な再現には使用しない。
---

# DOCX文書

原本と既存出力を上書きせず、`python-docx` で `.docx` を読み取り・編集する。レビューで根拠を示せるよう、段落番号、リスト種別・レベル、ハイパーリンクの表示文字とURL、表番号、行列、ヘッダー・フッター位置をMarkdownへ残す。

## 実行前に確認する

1. `python --version` でPython 3.12以上を確認する。
2. `python -c "import docx"` を実行する。失敗したら処理を始めず、READMEのインストール手順を案内する。
3. 入力が `.docx` であることを確認する。`.doc` は先に `convert-legacy-office` を使用する。
4. 入力と出力に別の未作成パスを指定する。既存のMarkdownまたはDOCXは上書きしない。出力パス、画像出力ディレクトリ、それらの祖先にシンボリックリンクを使用しない。

## Markdownへ変換する

次を実行する。`ROLE` は配置フォルダに合わせる。埋め込み画像を保存する場合は `--images-dir` を指定する。

```powershell
python .agents/skills/docx-document/scripts/docx_document.py to-markdown INPUT.docx OUTPUT.md --role target --repo-root . --images-dir work/images
```

全抽出画像の出力先を先に検証する。1件でも既存、壊れたリンク、重複、またはシンボリックリンク経由ならMarkdownを含めて何も書き込まず停止する。成功時は抽出画像への相対リンクをMarkdownに残す。画像内文字の認識は `image-document` を使用し、視覚的意味を推測しない。

## 限定編集を行う

```powershell
python .agents/skills/docx-document/scripts/docx_document.py edit INPUT.docx OUTPUT.docx --operations operations.json
```

許可する操作は `replace_text`、`add_paragraph`、`set_table_cell` だけとする。リンクを壊さないため、`replace_text` の対象文字がハイパーリンクを含む段落にある場合は明示的に拒否する。操作JSONの形式と索引の数え方は[操作仕様](references/operations.md)を読む。

## 出力を確認する

1. Markdownに6つの必須frontmatter項目があり、`source_path` がリポジトリ相対パスであることを確認する。
2. 段落と表の位置、リスト種別、ハイパーリンク、ヘッダー・フッター、抽出画像リンクを確認する。
3. 編集後のDOCXを再度読み、指定した変更だけが反映されたことを確認する。
4. 入力ファイルが変更されていないことを確認する。
