---
name: image-document
description: >-
  PNG、JPEG、TIFF、BMP、WebPの単一または複数フレーム画像を読み取り、寸法・モード・メタデータとTesseract OCRの文字・座標・信頼度を根拠位置付きMarkdownへ変換し、切り抜き・回転・サイズ変更・グレースケール・モード変換を新規画像へ適用する。設計書レビューで画像をチェックリスト、参考資料、レビュー対象として扱う場合、画像内文字をOCRする場合、または限定的な画像編集を依頼された場合に使用する。PDF、SVG、動画、生成画像には使用せず、OCRだけから図形・色・配置・写真の意味を断定しない。
---

# 画像文書を処理する

## 原則

- Python 3.12 以上と `Pillow` を先に確認する。
- OCR を行う場合だけ `pytesseract`、PATH 上の Tesseract、`jpn`、`eng`、必要に応じて `jpn_vert` を確認する。
- 入力画像を上書きしない。既存の出力先にも書き込まない。
- 主出力、画像出力フォルダ、予定する全抽出フレームについて、リンク切れを含むシンボリックリンクと、シンボリックリンクである既存祖先フォルダを処理前に拒否する。
- 全抽出フレームの競合を変換前に確認し、主出力は同じフォルダの一時ファイルで完成させてから新規確定する。
- 複数フレーム TIFF と WebP を全フレーム処理する。
- OCR が不明瞭な箇所と、OCR では判断できない視覚情報を `要確認` として扱う。

## Markdown へ変換する

1. リポジトリルート基準の入力パスと、未作成の Markdown 出力先を決める。
2. 各フレームの PNG コピーが必要なら `--images-dir` を指定する。
3. 画像中の文字をレビューに使う場合は `--ocr` を指定する。
4. 次を実行する。

```powershell
py -3 .agents/skills/image-document/scripts/image_document.py to-markdown INPUT OUTPUT `
  --role target --repo-root REPO_ROOT --images-dir IMAGES_DIR --ocr --lang jpn+eng
```

5. frontmatter の必須項目、元拡張子を保持した `source_format`、`image_width_px`、`image_height_px`、`ocr_executed`、配列の `ocr_languages`、`frame_count` を確認する。
6. 各フレームに標準 Markdown 画像リンクがあり、`--images-dir` 指定時は抽出 PNG、未指定時は元画像を Markdown からの相対パスで参照していることを確認する。
7. OCR 表の座標と信頼度を根拠に使い、低信頼または未認識の文字を推測しない。

## 編集する

1. [references/operations.md](references/operations.md) を最後まで読む。
2. 許可された操作だけを JSON に記述する。
3. 入力と異なる未作成の出力先を指定して実行する。

```powershell
py -3 .agents/skills/image-document/scripts/image_document.py edit INPUT OUTPUT `
  --operations OPERATIONS_JSON
```

4. 出力画像のフレーム数、寸法、モード、形式を確認する。
5. 複数フレームを出力する場合は TIFF または WebP を選ぶ。
6. 対象ファイルの編集を依頼された場合だけ `output/edited/` に保存する。

## エラー時に止める

- Pillow、pytesseract、Tesseract、言語データが不足したら、不足項目を利用者へ示して停止する。
- 既存出力、入力外の切り抜き、無効なフレーム番号、未対応操作を検出したら停止する。
- OCR 結果が空でも成功扱いで断定せず、Markdown に `要確認` を残す。
