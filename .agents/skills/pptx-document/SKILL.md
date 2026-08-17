---
name: pptx-document
description: PPTXプレゼンテーションのスライド、図形テキスト、表、ノート、通常画像と画像プレースホルダーを位置情報付きMarkdownへ変換し、背景画像参照を要確認として記録し、限定したテキスト編集を行う。input配下の.pptxをレビューするとき、画像を抽出するとき、利用者が依頼した.pptx編集をoutput/editedへ保存するときに使用する。.pptの変換、PDF、アニメーション、背景画像の抽出、視覚配置の意味判断には使用しない。
---

# PPTX文書

原本を上書きせず、`python-pptx` で `.pptx` を読み取り・編集する。レビューの根拠として、スライド番号、図形ID、座標、段落、表の行列、ノート位置をMarkdownへ残す。

## 実行前に確認する

1. `python --version` でPython 3.12以上を確認する。
2. `python -c "import pptx"` を実行する。失敗したら処理を始めず、READMEのインストール手順を案内する。
3. 入力が `.pptx` であることを確認する。`.ppt` は先に `convert-legacy-office` を使用する。
4. 入力と出力に別のパスを指定し、未作成の出力先を使用する。出力パス、画像出力ディレクトリ、それらの祖先にシンボリックリンクを使用しない。

## Markdownへ変換する

埋め込み画像を保存する場合は `--images-dir` を指定する。

```powershell
python .agents/skills/pptx-document/scripts/pptx_document.py to-markdown INPUT.pptx OUTPUT.md --role target --repo-root . --images-dir work/images
```

通常画像と画像が挿入済みの画像プレースホルダーを抽出する。抽出予定パスをすべて先に確認し、1件でも既存、壊れたリンク、重複、またはシンボリックリンク経由なら画像もMarkdownも書き込まない。

背景はスライド、レイアウト、マスターの順に明示背景を確認する。`p:bg/p:bgPr/a:blipFill/a:blip` と画像relationshipを検出した場合だけ、スライド番号、由来パート、relationship IDと参照先、「背景画像は抽出対象外・要確認」をMarkdownに記録する。単色、グラデーション等の背景には画像マーカーを出さない。背景画像自体の抽出は初版の対象外とする。

図形の重なり、矢印の意味、レイアウトの妥当性はテキストだけから推測しない。画像内文字は `image-document` でOCRし、判断できなければ `要確認` とする。

## 限定編集を行う

```powershell
python .agents/skills/pptx-document/scripts/pptx_document.py edit INPUT.pptx OUTPUT.pptx --operations operations.json
```

許可する操作は `replace_text` と `set_shape_text` だけとする。`set_shape_text` ではスライド番号と図形名を優先し、後方互換用に図形IDも使用できる。操作JSONの指定方法は[操作仕様](references/operations.md)を読む。既存の出力ファイルは上書きしない。

## 出力を確認する

1. Markdownに6つの必須frontmatter項目があり、`source_path` がリポジトリ相対パスであることを確認する。
2. 全スライドの図形名、図形ID、位置、表、ノート、抽出画像リンク、背景画像の要確認マーカーを確認する。
3. 編集後のPPTXを開き、指定したテキスト以外のレイアウトが変化していないことを確認する。
4. 入力ファイルが変更されていないことを確認する。
