# XLSX操作仕様

入力と出力には異なる未作成パスを指定する。既存の出力ファイルは上書きしない。出力パス自体または既存祖先がシンボリックリンクなら拒否し、壊れたリンクも既存出力として扱う。JSONは操作の配列、または `operations` 配列を持つオブジェクトとする。操作は記載順に適用する。

## 編集操作

```json
{
  "operations": [
    {"op": "set_cell", "sheet": "設計", "cell": "B3", "value": "更新値"},
    {"op": "rename_sheet", "sheet": "Sheet1", "new_name": "概要"},
    {"op": "add_sheet", "name": "改訂履歴", "index": 0}
  ]
}
```

- `set_cell`: 指定セルの値を置換する。`value` に `=SUM(A1:A3)` のような文字列を渡すと数式になる。
- `rename_sheet`: 既存シートを改名する。同名シートがある場合は失敗する。
- `add_sheet`: 新規シートを追加する。`index` は省略可能な0始まりの位置である。
- これら以外の操作名は拒否する。

## レビュー結果JSON

```json
{
  "checklists": [
    {"path": "input/checklists/design-review.xlsx", "header_row": 1}
  ],
  "targets": [
    {"path": "input/targets/basic-design.docx", "markdown": "work/markdown/input/targets/basic-design.docx.md"}
  ],
  "items": [
    {
      "checklist_file": "input/checklists/design-review.xlsx",
      "sheet": "チェックリスト",
      "row": 2,
      "check_item": "例外時の動作が定義されている",
      "target_file": "input/targets/basic-design.docx",
      "result": "不適合",
      "comment": "例外時の戻り値が記載されていない。対象の『3.2 API仕様』を確認した。",
      "evidence": ["work/markdown/input/targets/basic-design.docx.md#段落-18"],
      "improvement": "例外種別ごとのHTTPステータスとレスポンス例を追記する。"
    }
  ]
}
```

- `header_row` は1始まりで指定する。
- `sheet` と `row` はチェック項目の位置を表す。
- `result` は `適合`、`不適合`、`対象外`、`要確認` のいずれかとする。
- `comment` は空にせず、判断理由と利用者が確認できる位置を含める。
- `evidence` は根拠位置の配列とし、コメント欄の末尾にも追記される。
- 同一の `sheet`、`row`、`target_file` を重複させない。

## 埋め込み画像

`to-markdown` に `--images-dir` を指定すると、シート画像を `{元ファイル名}-sheet-{シート名}-image-{連番}.{拡張子}` で抽出する。画像ディレクトリ、予定画像、それらの既存祖先にあるシンボリックリンクは拒否する。全画像の出力予定を先に検証し、既存、壊れたリンク、または相互に重複するパスが1件でもあれば、画像もMarkdownも作らない。Markdownにはシート名、セルアンカー、抽出画像への相対リンクを記録する。
