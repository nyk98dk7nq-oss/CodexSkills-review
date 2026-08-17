# 作業用フォルダ

Skill が実行中に使用する中間物を格納します。利用者が事前にファイルを置く必要はありません。

実行時に次のサブフォルダを作成します。

- `converted-office/yyyyMMddhhmm/`: 旧 Office 形式から変換した作業用ファイル
- `images/yyyyMMddhhmm/`: Office/PDF から抽出または描画した画像
- `markdown/yyyyMMddhhmm/`: レビュー用に変換した Markdown
- `review-runs/yyyyMMddhhmm/`: `manifest.json`、`review_bundle.json`、AI が完成させる `results.json`

実行 ID ごとに中間物を分けるため、再実行しても過去の中間物を上書きしません。

中間物は通常の Git 管理対象になりません。
