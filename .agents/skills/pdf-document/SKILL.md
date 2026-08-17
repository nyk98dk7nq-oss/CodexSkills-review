---
name: pdf-document
description: >-
  PDF文書をページ単位で読み取り、メタデータ、テキスト、表、座標、画像領域、必要なOCR結果を根拠位置付きMarkdownへ変換し、回転・削除・並べ替え・結合・メタデータ変更を新規PDFへ適用する。設計書レビューで.pdfをチェックリスト、参考資料、レビュー対象として扱う場合、画像PDFへOCRする場合、または限定的なPDF編集を依頼された場合に使用する。Word、Excel、PowerPoint、単体画像、PDF以外の文書には使用せず、抽出結果だけから図形・色・配置の意味を断定しない。
---

# PDF 文書を処理する

## 原則

- Python 3.12 以上と、処理に必要な `pypdf`、`pdfplumber` を先に確認する。
- OCR を行う場合だけ `Pillow`、`pytesseract`、PATH 上の Tesseract と必要な言語データを確認する。
- 入力 PDF を上書きしない。既存の出力先にも書き込まない。
- 主出力、画像出力フォルダ、予定する全ページ画像について、リンク切れを含むシンボリックリンクと、シンボリックリンクである既存祖先フォルダを処理前に拒否する。
- 全ページ画像の競合を変換前に確認し、主出力は同じフォルダの一時ファイルで完成させてから新規確定する。
- 文書の役割は配置フォルダから `checklist`、`reference`、`target` のいずれかに決める。
- OCR が不明瞭な箇所と、テキスト抽出では判断できない視覚情報を `要確認` として扱う。

## Markdown へ変換する

1. リポジトリルート基準の入力パスと、未作成の Markdown 出力先を決める。
2. ページ画像が必要なら `--images-dir` を指定する。
3. 画像 PDF、埋め込み画像を含む PDF、またはテキストの少ないページを含む場合は `--ocr` を指定する。
4. 次を実行する。

```powershell
py -3 .agents/skills/pdf-document/scripts/pdf_document.py to-markdown INPUT OUTPUT `
  --role target --repo-root REPO_ROOT --images-dir IMAGES_DIR --ocr
```

5. YAML frontmatter の `source_path`、`source_name`、`source_format`、`document_role`、`converted_at`、`converter_skill` を確認する。
6. 各ページの抽出テキスト、表、座標、OCR の信頼度と根拠位置を確認する。
7. テキストが少なく OCR 未実施のページ、OCR できないページ、視覚情報だけの判断をレビューで推測しない。

`--ocr` はテキストの少ないページと、本文量にかかわらず埋め込み画像領域があるページを OCR する。全ページを OCR する必要がある場合は `--ocr-all` を指定する。言語は既定の `jpn+eng` を使い、必要な場合だけ `--lang` で変更する。

## 編集する

1. [references/operations.md](references/operations.md) を最後まで読む。
2. 許可された操作だけを JSON に記述する。
3. 入力と異なる未作成の出力先を指定して実行する。

```powershell
py -3 .agents/skills/pdf-document/scripts/pdf_document.py edit INPUT OUTPUT `
  --operations OPERATIONS_JSON
```

4. 出力 PDF のページ数、順序、回転、メタデータを確認する。
5. 対象ファイルの編集を依頼された場合だけ `output/edited/` に保存する。

## エラー時に止める

- ライブラリ、Tesseract、言語データ、ページ画像化手段が不足したら処理を開始または継続せず、不足項目を利用者へ示す。
- 暗号化 PDF、既存出力、入力外のページ番号、未対応操作を検出したら停止する。
- 失敗時に生成途中の PDF を成果物として扱わない。
