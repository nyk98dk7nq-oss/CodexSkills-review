# Excel検査JSONスキーマ

`inspect_excel.py`は、次の構造を持つUTF-8のJSONを出力する。現在の`schema_version`は`1.1`。キーが存在しない場合と値が`null`の場合を区別する。

## 最上位

| キー | 内容 |
|---|---|
| `schema_version` | 検査JSONの形式バージョン |
| `source` | 元ファイル名、拡張子、サイズ、SHA-256 |
| `workbook` | ブック全体のメタデータとシート一覧 |
| `sheets` | シート別の構造、セル、オブジェクト、警告 |
| `warnings` | ブック全体に適用する注意 |

## `workbook`

| キー | 内容 |
|---|---|
| `openpyxl_version` | 検査に使用したopenpyxlのバージョン。対応範囲は`>=3.1,<3.2` |
| `sheet_count` | 全シート数 |
| `included_sheet_count` | セル内容を抽出したシート数 |
| `sheet_order` | 元ブックどおりのシート順 |
| `worksheet_count` / `chartsheet_count` | ワークシート数とグラフシート数 |
| `chartsheets` | グラフシート名と表示状態。セル内容は抽出しない |
| `defined_names` | ブック全体と各ワークシートの名前定義。スコープ、対象シート、種類、非表示状態、値または除外状態を含む |
| `calculation` | Excelの計算モードと再計算フラグ |
| `archive` | ZIP展開後サイズ、部品数、VBA・外部リンク・描画・グラフ・メディア・VML・ActiveX・SmartArt部品の件数 |
| `date_system` | ブックが使用する日付システムの基準 |

## `workbook.defined_names[]`

| キー | 内容 |
|---|---|
| `name` | 名前定義の名前 |
| `scope` | `workbook`または`worksheet` |
| `sheet` | `worksheet`スコープのシート名。`workbook`スコープでは`null` |
| `type` | openpyxlが認識した定義の種類 |
| `hidden` | 名前定義が非表示か |
| `value` | 定義値。安全に出力できる場合だけ存在する |
| `value_excluded` | 値を伏せた場合だけ`true`。この場合は`value`を出力しない |

`wb.defined_names`だけでなく各`ws.defined_names`も抽出する。非表示の名前定義、除外シートをスコープとする名前定義、または除外シートを参照する名前定義は、値を出力せず`value_excluded: true`とする。シート名の引用符、外部ブック修飾、英字の大文字・小文字が異なる参照も保守的に判定する。

## `sheets[]`

| キー | 内容 |
|---|---|
| `name` / `state` | シート名と`visible`、`hidden`、`veryHidden` |
| `included` | `rows`へセル内容を抽出したか |
| `dimensions` | Excelの宣言範囲、内容に基づく実効範囲、抽出範囲、内容セル数 |
| `row_blocks` | 内容のある行が連続する候補領域。意味上の表を保証しない |
| `merged_ranges` | 結合範囲と左上セルの値 |
| `tables` | Excelテーブル名、範囲、スタイル |
| `data_validations` | 入力規則と対象範囲。除外シートまたは完全に非表示の対象範囲では式・エラー文を`details_excluded`で伏せる |
| `hidden_rows` / `hidden_columns` | 使用範囲内の非表示行・列 |
| `charts` / `images` | 種類とアンカー位置。内容そのものではない |
| `rows` | 元の行番号と、内容のあるセルの配列 |
| `warnings` | そのシート固有の注意 |

`dimensions.declared_range`は書式だけの遠方セルを含む場合がある。内容判断には`effective_range`を使い、非表示行・列を除いた実際の抽出範囲は`extracted_range`で確認する。`stored_content_cells`は保存値、数式、コメント、リンク、明示的な空文字のいずれかを持つセル数、`extracted_content_cells`は実際に抽出した数である。`included`が`false`のシートでは、存在と基本メタデータだけを示し、`effective_range`は`null`、`rows`は空になる。非表示範囲または除外シートの結合セルでは`value`を出力せず、`value_excluded: true`を示す。

## `cells[]`

| キー | 内容 |
|---|---|
| `coordinate` / `row` / `column` | セル位置 |
| `value` | 数式以外の保存値 |
| `formula` | 通常は`=`で始まる数式文字列。配列数式などは種類、式、対象範囲を持つオブジェクト |
| `cached_value` | Excelが最後に保存した数式結果。欠落または古い場合がある |
| `cached_value_state` | `present`、または未保存と空文字を区別できない`missing_or_empty` |
| `data_type` | openpyxlが認識したセル型 |
| `number_format` | Excelの表示形式 |
| `style` | 太字、斜体、塗り、配置、折り返し、字下げなどの非既定値 |
| `merged_range` | そのセルが左上セルである結合範囲 |
| `comment` | コメント本文と作成者 |
| `hyperlink` | リンク先と表示値 |

数式セルでは`formula`、`cached_value`、`cached_value_state`を確認し、`cached_value: null`を数値のゼロや確定した空文字と解釈しない。数式以外の明示的な空文字は`value: ""`として残す。`style`は意味判定の補助情報であり、見出しや分類を保証しない。

非表示シート、行、列の直接のセル値は既定で除外する。一方、表示セルの`cached_value`が除外セルに由来するかは追跡しない。派生値まで除外する必要がある場合は、該当する`cached_value`を使用しない。

## 安全な読み方

1. 最上位の`warnings`を読む。
2. `workbook.sheet_order`と各シートの`included`を確認する。
3. `dimensions`、`tables`、`merged_ranges`、`row_blocks`で構造候補を把握する。
4. `rows`の行番号とセル座標を保ったまま内容を読む。
5. `charts`または`images`、描画部品の警告があれば、可能な限りシートを視覚確認する。系列や内部設定は別の検査根拠がない限り断定しない。
6. Markdown化した各領域をセル範囲へ対応付ける。

検査前処理は`[Content_Types].xml`と内部Relationshipを解決し、拡張子やRelationshipの種類だけに依存せずXMLらしい参照部品を走査する。UTF-8・UTF-16・UTF-32相当の宣言文字列を正規化し、DTDとエンティティを拒否してからopenpyxlへ渡す。
