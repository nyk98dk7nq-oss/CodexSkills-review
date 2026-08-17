# レビュー結果JSONスキーマ

`results.json`はUTF-8のJSONオブジェクトとする。`prepare`が作ったbundleを複製し、`items`の空欄をCodexが埋める。

## ルート

| キー | 型 | 内容 |
|---|---|---|
| `run_id` | string | `yyyyMMddhhmm`形式の実行ID |
| `generated_at` | string | 生成日時 |
| `checklists` | array | `{path, header_row, markdown}` |
| `references` | array | `{path, markdown}` |
| `targets` | array | `{path, markdown}` |
| `checklist_items` | array | チェック項目の正規化一覧 |
| `items` | array | チェック項目と対象ファイルの全組合せの結果 |

`materials`と`review_instructions`はレビュー用の補助情報であり、残してよい。

## checklist_items

```json
{
  "checklist_file": "input/checklists/design-checklist.xlsx",
  "sheet": "基本設計",
  "row": 4,
  "check_item": "エラー応答が定義されているか"
}
```

`row`はコピー元のチェックリストに結果を書き戻す行番号である。

## items

```json
{
  "checklist_file": "input/checklists/design-checklist.xlsx",
  "sheet": "基本設計",
  "row": 4,
  "check_item": "エラー応答が定義されているか",
  "target_file": "input/targets/api-design.docx",
  "result": "不適合",
  "comment": "work/markdown/202601021530/targets/api-design.docx.md:L52-L63には正常応答だけがあり、参考資料のエラー応答要件を満たす定義を確認できない。",
  "evidence": [
    "work/markdown/202601021530/targets/api-design.docx.md:L52-L63",
    "work/markdown/202601021530/references/api-standard.pdf.md:L80-L88"
  ],
  "improvement": "HTTPステータス、業務エラーコード、メッセージ、発生条件を対応付けたエラー応答表をAPIごとに追加する。"
}
```

- `result`は`適合`、`不適合`、`対象外`、`要確認`のいずれか。
- `comment`には判断理由と、`evidence`に列挙した証拠位置を少なくとも1件そのまま含める。
- `evidence`は空にしない。
- `不適合`の`improvement`は空にしない。
- 同じ`checklist_file + sheet + row + target_file`を重複させない。
