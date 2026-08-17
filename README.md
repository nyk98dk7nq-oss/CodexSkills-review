# CodexSkills-review

Windows 11 上の Codex で、Excel、Word、PowerPoint、PDF を読み取り、Markdown へ変換し、必要に応じてファイルを編集する Skill 群の要件定義です。

設計書レビューでは、対象ファイルとチェックリストを先に Markdown へ変換し、生成された Markdown をレビュー対象にします。編集を依頼された場合は、形式ごとの Python ライブラリを使って編集済みファイルを出力します。

このリポジトリでは、ルートの `README.md` を Skill Creator へ渡す唯一の指示書とします。完成した Skill、入力ファイル、生成 Markdown、編集済みファイル、レビュー結果は、このリポジトリには格納しません。

## 目次

1. [目的](#1-目的)
2. [基本方針](#2-基本方針)
3. [対応形式](#3-対応形式)
4. [作成する Skill](#4-作成する-skill)
5. [使用するライブラリ](#5-使用するライブラリ)
6. [Windows 11 へのインストール](#6-windows-11-へのインストール)
7. [作業フォルダ構成](#7-作業フォルダ構成)
8. [ファイル読み取りと Markdown 変換](#8-ファイル読み取りと-markdown-変換)
9. [ファイル編集](#9-ファイル編集)
10. [YAML フロントマター](#10-yaml-フロントマター)
11. [設計書レビュー](#11-設計書レビュー)
12. [旧 Office 形式の扱い](#12-旧-office-形式の扱い)
13. [形式別の要件](#13-形式別の要件)
14. [テストと完了条件](#14-テストと完了条件)
15. [Skill Creator への実装指示](#15-skill-creator-への実装指示)

## 1. 目的

次の作業を Codex から一貫して実行できるようにします。

1. Excel、Word、PowerPoint、PDF の内容を読み取る。
2. 各ファイルをレビューしやすい Markdown へ変換する。
3. Excel 形式のチェックリストと、必要に応じて与えられる参考資料を読み取る。
4. 生成した Markdown を使って設計書をレビューする。
5. ユーザーの指示に基づき、元の形式を保った編集済みファイルを作成する。
6. 旧 Office 形式は、Microsoft Office を使って新形式へ変換してから処理する。

## 2. 基本方針

1. Windows 11 上の Codex で動作させる。
2. 作業フォルダを他の利用者へ渡せるよう、設定と成果物のパスは相対パスにする。
3. `.xlsx`、`.docx`、`.pptx`、`.pdf` ごとに独立した Skill を作る。
4. 各形式の Skill は、読み取り、Markdown 変換、編集の3つを担当する。
5. レビューの前に、チェックリスト、レビュー対象、参考資料をすべて Markdown へ変換する。
6. AI は生成済み Markdown を使って、文書の役割とチェック項目の適用先を判断する。
7. 編集済みファイルは `output/edited/` へ別ファイルとして保存する。
8. 通常の処理では `input/` に置いたファイルを上書きしない。
9. ライブラリは、無償で利用でき、商用利用可能なものを使用する。
10. LibreOffice は使用しない。

## 3. 対応形式

| 入力形式 | 読み取り | Markdown 変換 | 編集 | 処理方法 |
|---|---:|---:|---:|---|
| `.xlsx` | 対応 | 対応 | 対応 | `openpyxl` |
| `.docx` | 対応 | 対応 | 対応 | `python-docx` |
| `.pptx` | 対応 | 対応 | 対応 | `python-pptx` |
| `.pdf` | 対応 | 対応 | 一部対応 | `pypdf`、`pdfplumber`、`reportlab` |
| `.xls` | 対応 | 対応 | 対応 | Microsoft Excel で `.xlsx` へ変換後に処理 |
| `.doc` | 対応 | 対応 | 対応 | Microsoft Word で `.docx` へ変換後に処理 |
| `.ppt` | 対応 | 対応 | 対応 | Microsoft PowerPoint で `.pptx` へ変換後に処理 |

初版では、マクロ付き形式、パスワード付きファイル、破損ファイル、画像だけで構成された PDF の OCR は対象外とします。

PDF は Office ファイルのように任意の文章やレイアウトを自由に書き換えられる形式ではありません。初版の PDF 編集は、ページの追加・削除・並べ替え・回転、結合・分割、メタデータ、フォーム、注釈、文字や図形の追記を対象とします。本文そのものを大きく修正する場合は、可能であれば変換元の Word、Excel、PowerPoint を編集して PDF を再作成します。

## 4. 作成する Skill

| Skill 名 | 主な役割 | 入力 | 出力 |
|---|---|---|---|
| `xlsx-document` | XLSX の読み取り、Markdown 変換、編集 | `.xlsx` | `.md` または編集済み `.xlsx` |
| `docx-document` | DOCX の読み取り、Markdown 変換、編集 | `.docx` | `.md` または編集済み `.docx` |
| `pptx-document` | PPTX の読み取り、Markdown 変換、編集 | `.pptx` | `.md` または編集済み `.pptx` |
| `pdf-document` | PDF の読み取り、Markdown 変換、対応可能な編集 | `.pdf` | `.md` または編集済み `.pdf` |
| `convert-legacy-office` | 旧 Office 形式を新形式へ変換 | `.xls/.doc/.ppt` | `.xlsx/.docx/.pptx` |
| `review-markdown-documents` | Markdown 化された文書をチェックリストでレビュー | `.md` 一式 | レビュー結果 `.md` |
| `review-documents-orchestrator` | 形式判定、変換、レビューを順番に実行 | 対応ファイル一式 | Markdown、編集済みファイル、レビュー結果 |

形式ごとの読み取り、変換、編集を1つの Skill にまとめます。読み取り用 Skill と編集用 Skill を別々に増やしません。

## 5. 使用するライブラリ

### 5.1 直接使用するライブラリ

| ライブラリ | 対象 | 読み取りで使う機能 | 編集で使う機能 | ライセンス | 料金 | 商用利用 | 注意事項 |
|---|---|---|---|---|---|---|---|
| `openpyxl` | XLSX | シート、セル、数式、スタイル、コメント、結合セルなど | セル値、数式、書式、シート、コメント、画像、グラフなど | MIT | 無料 | 可 | Excel の計算エンジンではないため、数式の再計算は行わない |
| `python-docx` | DOCX | 段落、見出し、表、スタイル、画像関係など | 文章、段落、表、スタイル、画像、セクションなど | MIT | 無料 | 可 | 変更履歴や一部の高度な Word 機能は直接扱えない |
| `python-pptx` | PPTX | スライド、図形、テキスト、表、画像、ノートなど | テキスト、図形、表、画像、配置、サイズ、スライドなど | MIT | 無料 | 可 | SmartArt、アニメーションなどは完全には編集できない |
| `pypdf` | PDF | ページ、テキスト、メタデータ、フォームなど | 結合、分割、回転、ページ操作、メタデータ、フォーム、注釈など | BSD-3-Clause | 無料 | 可 | 既存本文の自由な書き換えには向かない |
| `pdfplumber` | PDF | ページごとのテキスト、単語、表、座標など | 原則として読み取りに使用 | MIT | 無料 | 可 | スキャン PDF の文字認識は行わない |
| `reportlab` | PDF | 既存 PDF の解析には使用しない | 新しい PDF、追記用ページ、文字・図形の重ね合わせを作成 | BSD | 無料 | 可 | 既存 PDF への反映は `pypdf` と組み合わせる |
| `pywin32` | 旧 Office | Microsoft Office の COM API を呼び出す | `.xls/.doc/.ppt` の変換と、必要な Office 操作 | 複数の許諾ライセンス | 無料 | 可 | Windows とインストール済み Microsoft Office が必要。同梱ライセンス文書に従う |

### 5.2 主な依存関係

依存ライブラリは `pip` が自動的にインストールします。Skill のコードから直接利用しない依存ライブラリを、個別にインストールする手順は作りません。

| 親ライブラリ | 主な依存ライブラリ | 主な用途 |
|---|---|---|
| `openpyxl` | `et-xmlfile` | XLSX 内の XML 処理 |
| `python-docx` | `lxml`、`typing-extensions` | DOCX 内の XML 処理、型補助 |
| `python-pptx` | `lxml`、`Pillow`、`XlsxWriter`、`typing-extensions` | PPTX 内の XML、画像、グラフデータ処理 |
| `pdfplumber` | `pdfminer.six`、`Pillow`、`pypdfium2` | PDF テキスト解析、画像、ページ描画 |
| `reportlab` | `Pillow`、`charset-normalizer` | 画像と文字コード処理 |
| `pypdf` | 追加必須依存なし | PDF の基本操作 |
| `pywin32` | 追加必須依存なし | Windows API と COM 呼び出し |

依存関係はライブラリの更新によって変わる場合があります。実際にインストールされた一覧は、必要になったときだけ次のコマンドで確認します。

```powershell
py -3.12 -m pip list
```

## 6. Windows 11 へのインストール

### 6.1 前提

- Windows 11 64bit
- CPython 3.12 64bit
- `py` コマンドが利用できること
- `.xls`、`.doc`、`.ppt` を扱う場合は Microsoft Office デスクトップ版がインストールされていること

PowerShell を開き、Python を確認します。

```powershell
py -3.12 --version
py -3.12 -m pip --version
```

`pip` を更新します。

```powershell
py -3.12 -m pip install --upgrade pip
```

### 6.2 ライブラリを1つずつグローバルインストールする

XLSX 用ライブラリをインストールします。

```powershell
py -3.12 -m pip install openpyxl
```

DOCX 用ライブラリをインストールします。

```powershell
py -3.12 -m pip install python-docx
```

PPTX 用ライブラリをインストールします。

```powershell
py -3.12 -m pip install python-pptx
```

PDF の基本操作用ライブラリをインストールします。

```powershell
py -3.12 -m pip install pypdf
```

PDF のテキスト・表読み取り用ライブラリをインストールします。

```powershell
py -3.12 -m pip install pdfplumber
```

PDF の追記・生成用ライブラリをインストールします。

```powershell
py -3.12 -m pip install reportlab
```

旧 Office 形式用ライブラリをインストールします。

```powershell
py -3.12 -m pip install pywin32
```

### 6.3 インストール確認

```powershell
py -3.12 -c "import openpyxl; print('openpyxl: OK')"
py -3.12 -c "import docx; print('python-docx: OK')"
py -3.12 -c "import pptx; print('python-pptx: OK')"
py -3.12 -c "import pypdf; print('pypdf: OK')"
py -3.12 -c "import pdfplumber; print('pdfplumber: OK')"
py -3.12 -c "import reportlab; print('reportlab: OK')"
py -3.12 -c "import win32com.client; print('pywin32: OK')"
```

すべて `OK` と表示されれば準備完了です。

## 7. 作業フォルダ構成

```text
<作業フォルダ>/
├─ AGENTS.md
├─ skills/
│  ├─ xlsx-document/
│  ├─ docx-document/
│  ├─ pptx-document/
│  ├─ pdf-document/
│  ├─ convert-legacy-office/
│  ├─ review-markdown-documents/
│  └─ review-documents-orchestrator/
├─ input/
│  └─ files/
├─ work/
│  ├─ converted-office/
│  └─ markdown/
└─ output/
   ├─ edited/
   └─ reviews/
```

各パスは作業フォルダを基準にした相対パスとして扱います。利用者ごとのユーザー名、ドライブ名、絶対パスを設定へ埋め込みません。

## 8. ファイル読み取りと Markdown 変換

処理順序は次のとおりです。

```mermaid
flowchart TD
    A[入力ファイルを列挙] --> B{旧Office形式か}
    B -->|はい| C[新形式へ変換]
    B -->|いいえ| D[形式別Skillで読み取り]
    C --> D
    D --> E[Markdownを作成]
    E --> F[AIが文書の役割を判定]
    F --> G[Markdownをレビュー]
```

1. `input/files/` にある対応ファイルを列挙する。
2. `.xls`、`.doc`、`.ppt` は、作業用の新形式へ変換する。
3. 形式ごとの Skill で、本文、表、見出し、シート名、スライド番号などを読み取る。
4. `work/markdown/` にファイルごとの Markdown を作成する。
5. Markdown の先頭に、変換元ファイル情報を YAML フロントマターで記録する。
6. AI が各 Markdown を `checklist`、`target`、`reference`、`unknown` のいずれかへ分類する。
7. 適用するチェックリストと対象文書の組み合わせを決める。
8. Markdown を使ってレビューし、結果を `output/reviews/` へ保存する。

## 9. ファイル編集

編集処理は、ユーザーの編集指示を形式別 Skill が受け取り、編集済みファイルを作成する流れにします。

1. 対象ファイルと編集内容を確認する。
2. 対象形式のライブラリでファイルを開く。
3. 指定された箇所だけを変更する。
4. `output/edited/` に別名で保存する。
5. 保存したファイルを同じライブラリでもう一度開く。
6. 指定した変更が反映されていることを確認する。
7. 変更したファイル、変更箇所、出力先を Markdown で報告する。

編集済みファイル名は、元の名前へ `_edited` を付けます。

```text
input/files/basic-design.xlsx
output/edited/basic-design_edited.xlsx
```

同名ファイルがすでにある場合は、日時を付けて別名にします。

```text
output/edited/basic-design_edited_20260817-120000.xlsx
```

## 10. YAML フロントマター

生成する Markdown の先頭には、変換元ファイルの情報を記載します。

```yaml
---
source_path: "input/files/basic-design.xlsx"
source_name: "basic-design.xlsx"
source_format: "xlsx"
converted_at: "2026-08-17T12:00:00Z"
converter_skill: "xlsx-document"
---
```

旧 Office 形式を新形式へ変換した場合は、中間ファイルも記載します。

```yaml
---
source_path: "input/files/basic-design.xls"
source_name: "basic-design.xls"
source_format: "xls"
intermediate_path: "work/converted-office/basic-design.xlsx"
intermediate_format: "xlsx"
converted_at: "2026-08-17T12:00:00Z"
converter_skill: "xlsx-document"
---
```

YAML フロントマターには、レビュー時に元ファイルを識別するために必要な項目だけを記録します。

## 11. 設計書レビュー

### 11.1 文書の役割

AI は Markdown の内容を読み、各文書を次のいずれかへ分類します。

| 役割 | 内容 |
|---|---|
| `checklist` | チェック項目、確認観点、判定基準が記載された文書 |
| `target` | レビュー対象の設計書、仕様書、計画書など |
| `reference` | チェック観点や判定基準を補足する文書 |
| `unknown` | 内容だけでは役割を判断できない文書 |

役割はファイル名だけで決めず、見出し、表の列名、本文から判断します。判断できない場合は、ユーザーへ候補を示して確認します。

### 11.2 レビュー結果

判定は次の4種類とします。

- `適合`
- `不適合`
- `対象外`
- `要確認`

レビュー結果には、次を記載します。

| 項目 | 内容 |
|---|---|
| チェック項目 | チェックリストの項目名と位置 |
| 対象ファイル | レビューした元ファイル |
| 判定 | `適合/不適合/対象外/要確認` |
| 理由 | 判定した理由 |
| 根拠 | Markdown の見出し、表、セル、スライド、ページなど |
| 修正案 | 必要な場合だけ記載 |

Excel のチェックリストに結果記入用の列がある場合は、`xlsx-document` Skill を使ってレビュー結果を記入した編集済み XLSX も作成できるようにします。

## 12. 旧 Office 形式の扱い

`.xls`、`.doc`、`.ppt` は、Python ライブラリだけで直接処理しません。Windows 11 にインストールされた Microsoft Office を `pywin32` から操作し、作業用の新形式へ変換します。

| 旧形式 | 使用するアプリ | 変換先 |
|---|---|---|
| `.xls` | Microsoft Excel | `.xlsx` |
| `.doc` | Microsoft Word | `.docx` |
| `.ppt` | Microsoft PowerPoint | `.pptx` |

変換先は `work/converted-office/` とします。変換後は、新形式に対応する Skill で読み取り、Markdown 変換、編集を行います。

Microsoft Office がインストールされていない場合は、対象の旧形式を処理できないことを明確に表示します。LibreOffice への切り替えは行いません。

## 13. 形式別の要件

### 13.1 XLSX

読み取りでは、少なくとも次を Markdown に含めます。

- ブック名、シート名、シート順
- セル座標と値
- 数式
- 結合セル
- 表
- コメント
- ハイパーリンク
- 非表示シート、行、列
- 画像やグラフが存在すること

編集では、少なくとも次へ対応します。

- セル値と数式の変更
- 行と列の追加・削除
- シートの追加・削除・名称変更
- 文字、背景色、罫線、配置、表示形式の変更
- コメントの追加・変更
- 表と画像の追加
- チェックリストへのレビュー結果記入

### 13.2 DOCX

読み取りでは、少なくとも次を文書順に Markdown へ含めます。

- 見出し
- 段落
- 箇条書きと番号付きリスト
- 表
- ヘッダーとフッター
- ハイパーリンク
- 画像が存在すること

編集では、少なくとも次へ対応します。

- 文章の追加、置換、削除
- 見出しと段落の追加
- 表の追加とセル編集
- スタイルと文字書式の変更
- 画像の追加
- ヘッダーとフッターの編集

### 13.3 PPTX

読み取りでは、少なくとも次をスライド順に Markdown へ含めます。

- スライド番号とタイトル
- テキスト
- 図形名、種類、位置
- 表
- 画像が存在すること
- 発表者ノート

編集では、少なくとも次へ対応します。

- スライドの追加
- テキストの追加、置換、削除
- 図形の追加、移動、サイズ変更
- 表の追加とセル編集
- 画像の追加
- 文字、色、線、背景の変更

### 13.4 PDF

読み取りでは、少なくとも次をページ順に Markdown へ含めます。

- ページ番号
- 抽出できる本文
- 表
- メタデータ
- 注釈とフォーム項目

編集では、初版で次へ対応します。

- ページの追加、削除、並べ替え、回転
- PDF の結合と分割
- メタデータの変更
- フォームへの値入力
- 注釈の追加
- 文字、線、図形、画像の追記

本文を抽出できないページは、そのページを読み取れなかったことを Markdown へ記載します。OCR は別 Skill として必要になった時点で検討します。

## 14. テストと完了条件

### 14.1 形式別テスト

各 Skill に、次の最小テストを作成します。

1. サンプルファイルを読み取れる。
2. Markdown を作成できる。
3. YAML フロントマターに元ファイル情報が入る。
4. 代表的な1項目を編集できる。
5. 編集済みファイルを再度開ける。
6. 指定した変更が編集済みファイルに反映されている。
7. `input/` のファイルが通常処理で上書きされていない。

### 14.2 旧形式テスト

Windows 11 と Microsoft Office デスクトップ版を使って、次を確認します。

1. `.xls` を `.xlsx` へ変換できる。
2. `.doc` を `.docx` へ変換できる。
3. `.ppt` を `.pptx` へ変換できる。
4. 変換後の新形式を対応 Skill で読み取れる。
5. 変換後の新形式を編集できる。

### 14.3 レビューの一連テスト

1. `input/files/` にチェックリスト、設計書、参考資料を置く。
2. すべての対応ファイルを Markdown へ変換する。
3. AI が文書の役割と適用関係を判断する。
4. Markdown を使って設計書をレビューする。
5. 根拠付きのレビュー結果を出力する。
6. 必要に応じて、レビュー結果をチェックリスト XLSX へ記入する。
7. 編集済み XLSX を再度開き、記入結果を確認する。

初版の完了条件は、4つの形式別 Skill、旧形式変換 Skill、レビュー Skill、orchestrator Skill が動作し、読み取り、Markdown 変換、編集、レビューの一連の処理を実行できることです。

## 15. Skill Creator への実装指示

Skill Creator へは、この README 全体を要件として渡し、次の順で Skill を作成させます。

1. `xlsx-document`
2. `docx-document`
3. `pptx-document`
4. `pdf-document`
5. `convert-legacy-office`
6. `review-markdown-documents`
7. `review-documents-orchestrator`

各 Skill は次の構成にします。

```text
<skill-name>/
├─ SKILL.md
├─ agents/
│  └─ openai.yaml
├─ scripts/
└─ tests/
```

実装時の共通指示は次のとおりです。

1. `SKILL.md` には、Skill の起動条件、入力、出力、処理手順、制限を記載する。
2. 読み取り、Markdown 変換、編集は再実行可能な Python スクリプトとして作る。
3. スクリプトの引数には相対パスを使用する。
4. 生成先は `work/` または `output/` とする。
5. Markdown には元ファイル情報の YAML フロントマターを付ける。
6. 編集後は、出力ファイルを再度開いて変更内容を確認する。
7. 形式別テストと一連テストを作る。
8. README に記載していない変換サービス、クラウドサービス、LibreOffice は追加しない。

実装を必要以上に複雑にせず、まずは「読める」「Markdown にできる」「編集できる」「レビューできる」の4点を完成させます。
