---
name: convert-legacy-office
description: >-
  Windows上のpywin32とインストール済みMicrosoft Excel、Word、PowerPointのCOMを使い、旧Office形式の.xls、.doc、.pptをそれぞれ.xlsx、.docx、.pptxへ新規保存し、フォルダ単位でも相対構成を保って変換する。設計書レビューの入力に旧Office形式があり、Markdown変換前に新形式へ変換する場合に使用する。新Office形式、PDF、画像には使用せず、LibreOfficeや外部変換サービスへ代替しない。
---

# 旧 Office 形式を変換する

## 原則

- Windows 11、Python 3.12 以上、`pywin32`、対象形式に対応する Microsoft Office アプリの COM 登録を先に確認する。
- 入力原本を上書きしない。既存の出力先にも書き込まない。
- 出力、出力フォルダ、その既存祖先フォルダについて、リンク切れを含むシンボリックリンクを変換前に拒否する。
- 単一変換は同じフォルダの一時ファイルへ保存してから新規確定し、バッチ変換が途中失敗した場合は今回作成した出力だけをすべてロールバックする。
- `.xls` は `.xlsx`、`.doc` は `.docx`、`.ppt` は `.pptx` にだけ変換する。
- 変換後のファイルを `work/converted-office/` に置き、元入力と中間ファイルの対応はオーケストレータの manifest に保持する。
- LibreOffice や外部変換サービスを使用しない。

## 単一ファイルを変換する

1. 入力形式に対応する Office アプリがインストール済みか確認する。
2. 入力と異なる未作成の出力先を `work/converted-office/` 配下に決める。
3. 次を実行する。

```powershell
py -3 .agents/skills/convert-legacy-office/scripts/convert_legacy_office.py convert INPUT OUTPUT
```

4. 正しい Open XML 拡張子の出力が作られたことを確認する。
5. 変換後のファイルへ、対応する形式別 Skill を適用する。

## フォルダを一括変換する

入力フォルダを再帰走査し、相対フォルダ構成を維持して出力する。

```powershell
py -3 .agents/skills/convert-legacy-office/scripts/convert_legacy_office.py batch INPUT_DIR OUTPUT_DIR
```

処理開始前に、計画された出力先に既存ファイルがないことを確認する。1件でも競合すればバッチ全体を開始しない。

## 詳細を確認する

COM アプリ、保存形式番号、終了処理、バッチの扱いを変更する前に [references/operations.md](references/operations.md) を最後まで読む。

## エラー時に止める

- Windows 以外、`pywin32` 不足、Office COM 未登録、拡張子不一致、既存出力を検出したら停止する。
- 成功・失敗を問わず、開いた文書を閉じ、Office アプリを終了し、COM を解放する。
- 失敗途中に作成された出力を成果物として残さない。
