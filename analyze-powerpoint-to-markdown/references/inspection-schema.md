# PowerPoint検査JSONスキーマ

`inspect_powerpoint.py`は、次の構造を持つUTF-8のJSONを出力する。配列順は決定的であり、キーが存在しない場合と値が`null`の場合を区別する。

## 目次

- [1. 最上位](#1-最上位)
- [2. presentation](#2-presentation)
- [3. slides](#3-slides)
- [4. shapes](#4-shapes)
- [5. paragraphsとruns](#5-paragraphsとruns)
- [6. chart](#6-chart)
- [7. 安全な読み方](#7-安全な読み方)

## 1. 最上位

| キー | 内容 |
|---|---|
| `schema_version` | 検査JSONの形式バージョン |
| `source` | 元ファイル名、拡張子、サイズ、SHA-256 |
| `presentation` | プレゼンテーション全体のメタデータと安全検査結果 |
| `slides` | スライド別の構造、図形、ノート、警告 |
| `warnings` | プレゼンテーション全体に適用する注意 |

## 2. `presentation`

| キー | 内容 |
|---|---|
| `slide_count` | 全スライド数 |
| `included_slide_count` | 内容を抽出したスライド数 |
| `hidden_slide_count` | 非表示スライド数 |
| `slide_size` | 幅と高さ。EMUとインチを含む |
| `selected_slides` | `--slide`で明示した1始まりのスライド番号 |
| `document_properties_included` | 文書プロパティを抽出したか |
| `document_properties` | 明示指定時だけ含む作成者、タイトル、更新日時など |
| `archive` | 実際とZIP宣言上の展開後サイズ、部品数、外部Relationship、VBA、OLE、ActiveX、SmartArt、グラフ、画像、コメント、音声、動画などの件数。`relationship_types`は既知の固定ラベルだけを使い、未知の種類は`other`へ集約 |
| `limits` | 実行時に適用したスライド、図形、表セル、グラフ点、文字数の上限 |

`archive`の在庫は、プレゼンテーション全体の部品数である。非表示スライドまたは除外したスライドに属する部品も件数へ含まれるが、その内容、リンク先、ファイル名、任意のRelationship Type文字列は出力しない。`has_vba`は、`.pptm`拡張子、既知のVBA Relationship、Content-Type、標準部品名のいずれかを検出した場合に保守的に`true`となる。

## 3. `slides[]`

| キー | 内容 |
|---|---|
| `number` / `slide_id` | 1始まりの順番とPowerPoint内部のスライドID |
| `hidden` | PowerPointで非表示に設定されたスライドか |
| `included` / `exclusion_reason` | 内容を抽出したか、除外理由 |
| `layout_name` | 抽出対象スライドが使用するレイアウト名 |
| `title` | タイトルプレースホルダーから取得したタイトルと図形参照 |
| `has_notes` / `notes_included` | ノートの存在と内容抽出の有無 |
| `notes` | 明示指定時の発表者ノート本文と段落 |
| `shape_count` / `included_shape_count` | 全図形数と内容を抽出した図形数。グループ内図形を含む |
| `excluded_shapes` | 非表示または完全に領域外として内容を伏せた図形の件数 |
| `shapes` | Zオーダー順の図形情報 |
| `warnings` | そのスライド固有の注意 |

`included: false`のスライドでは、タイトル、レイアウト名、図形、ノートの内容を出力しない。非表示スライドを`--slide`で明示した場合は、明示選択として抽出する。

## 4. `shapes[]`

| キー | 内容 |
|---|---|
| `shape_id` / `name` / `shape_type` | PowerPoint内部ID、図形名、種類 |
| `path` | グループ階層を含む図形参照用パス |
| `z_order` | 同じ図形集合内の背面から前面への順番。読み順ではない |
| `position` | 左、上、幅、高さ。EMU、インチ、スライド比率を含む |
| `position_resolved` | グループ階層を含む座標をスライド座標へ正規化できたか |
| `hidden` | 図形の非表示属性を検出したか |
| `fully_off_slide` / `partially_off_slide` | スライド領域の外側にあるか |
| `included` / `exclusion_reason` | 内容を抽出したか、除外理由 |
| `placeholder_type` | プレースホルダーの場合の種類 |
| `alt_text` | 明示的な代替説明またはタイトル |
| `text` / `paragraphs` | 図形本文、段落レベル、箇条書き種別、ラン、書式、リンク |
| `table` | 行列、幅、高さ、セル本文、結合元・結合先、スパン |
| `chart` | 種類、タイトル、系列、カテゴリ、保存済み値、抽出警告 |
| `image` | 元ファイル名、Content-Type、拡張子、サイズ、SHA-256、明示抽出したパス、未抽出形式の警告 |
| `click_action` | 図形クリック時のアクション、外部リンク、対象スライドIDを取得できた場合の情報 |
| `children` | グループ図形内の子図形 |
| `warnings` | その図形の未対応または抽出エラー |

図形名、代替テキスト、リンク、表、グラフ、画像情報は、`included: true`の図形だけに含める。除外図形は`shape_id`、`path`、表示状態、除外理由だけを示し、内容を伏せる。グループ図形の子は`off`、`ext`、`chOff`、`chExt`、回転、反転を親から順に合成してスライド座標へ変換する。領域内外は回転後の外接矩形だけでなく、変換後の四角形とスライド矩形の交差で判定し、幅または高さが0のコネクターは線分とスライド矩形の交差で判定する。正規化できない子は`position_resolved: false`および`unresolved_position`として、`--include-off-slide-shapes`を明示しない限り内容を除外する。

## 5. `paragraphs[]`と`runs[]`

| キー | 内容 |
|---|---|
| `index` | 0始まりの段落番号 |
| `text` | 段落本文。ソフト改行は改行文字へ正規化 |
| `level` | PowerPointの段落レベル |
| `bullet` | `character`、`numbered`、`none`、`inherited_or_unspecified` |
| `alignment` | 明示されている場合の配置 |
| `runs` | 文字列ラン。太字、斜体、下線、フォント、サイズ、リンクを含む |

書式値が`null`の場合は、レイアウト、マスター、テーマから継承されている可能性がある。`null`を既定の書式と断定しない。

## 6. `chart`

| キー | 内容 |
|---|---|
| `chart_type` | `python-pptx`が認識したグラフ種類 |
| `title` | グラフタイトル。存在しない場合は`null` |
| `series` | 系列名と`python-pptx`が返した保存済み値 |
| `categories` | 取得できたカテゴリラベル。複数プロットではプロット別 |
| `cached_values_only` | 常に`true`。外部または埋め込みブックを再計算していないことを示す |
| `warnings` | 対応外グラフ、取得失敗、視覚確認が必要な項目 |

## 7. 安全な読み方

1. 最上位の`warnings`を読む。
2. `presentation.archive`でVBA、外部Relationship、OLE、ActiveX、SmartArt、コメント、メディアを確認する。
3. 各スライドの`hidden`、`included`、`notes_included`を確認する。
4. 図形の`included`、位置、種類、段落、表、図表を読む。
5. `z_order`を読み順と誤認せず、複数列や図解はスライド画像で確認する。
6. Markdown化した各記述をスライド番号、図形ID、表セルへ対応付ける。

検査前処理は、ZIP部品名、展開量、重複、暗号化、内部Relationshipの解決先、XML宣言を検査する。拡張子やContent-Typeだけに依存せず、XMLらしい全ての部品を走査し、UTF-8・UTF-16・UTF-32相当のDTDとエンティティ宣言を拒否する。その後、元ファイルをパスで渡さず、検査済みバイト列をメモリから`python-pptx`へ読み込む。ファイルの保存APIは使用しない。
