# MarkdownからPDFを生成するための導入要件

## 目次

- [1. 概要](#1-概要)
- [2. 必要なソフトウェア](#2-必要なソフトウェア)
- [3. Windowsでのセットアップ](#3-windowsでのセットアップ)
- [4. macOSでのセットアップ](#4-macosでのセットアップ)
- [5. PDFの生成](#5-pdfの生成)
- [6. オプション](#6-オプション)
- [7. 対応範囲と制約](#7-対応範囲と制約)
- [8. トラブルシューティング](#8-トラブルシューティング)

## 1. 概要

`write-vscode-markdown/scripts/markdown_to_pdf.mjs`は、MarkdownをHTMLへ変換し、Mermaid図を描画してからPlaywrightのChromiumでPDFを生成します。

外部CDNを使用せず、インストール済みのnpmパッケージだけで変換するため、初回セットアップ後はオフライン環境でも利用できます。

## 2. 必要なソフトウェア

| ソフトウェア | 必須 | 用途 |
|---|---:|---|
| Node.js 20以上25未満 | 必須 | PDF変換スクリプトの実行 |
| npm 9以上12未満 | 必須 | JavaScriptライブラリの導入 |
| Playwright Chromium | 必須 | HTMLの描画とPDF出力 |
| Python 3.10以上 | 推奨 | `validate_markdown.py`による事前検証 |
| VS Code | 任意 | MarkdownとMermaidのプレビュー |

このPDF変換機能ではPythonライブラリを追加インストールしません。Pythonは既存のMarkdownバリデーターにのみ使用します。

利用するnpmパッケージは次のとおりです。

| パッケージ | 固定バージョン | 用途 |
|---|---:|---|
| `markdown-it` | `14.3.0` | MarkdownからHTMLへの変換 |
| `mermaid` | `11.16.0` | Mermaid図の描画 |
| `playwright` | `1.62.1` | ChromiumによるPDF生成 |
| `highlight.js` | `11.11.1` | コードブロックのシンタックスハイライト |

直接依存と推移的依存は`package-lock.json`で固定します。`npm ci`は`package.json`とロックファイルが一致しない場合に失敗するため、意図しない依存更新を防げます。

## 3. Windowsでのセットアップ

PowerShellでリポジトリのスキルディレクトリへ移動します。

```powershell
cd "<repository>\write-vscode-markdown"
npm ci
npx playwright install chromium
```

`npm ci`は`package-lock.json`に固定された依存関係を再現します。`npm install`は依存バージョンを意図的に更新し、ロックファイルを更新する場合だけ使用してください。

Node.jsが未導入の場合は、Node.js公式インストーラーまたはWindows Package Managerを使用します。

```powershell
winget install OpenJS.NodeJS.LTS
```

インストール後、新しいPowerShellを開いて確認します。

```powershell
node --version
npm --version
```

## 4. macOSでのセットアップ

ターミナルでリポジトリのスキルディレクトリへ移動します。

```bash
cd "<repository>/write-vscode-markdown"
npm ci
npx playwright install chromium
```

`npm ci`は`package-lock.json`に固定された依存関係を再現します。`npm install`は依存バージョンを意図的に更新し、ロックファイルを更新する場合だけ使用してください。

Homebrewで対応範囲内のNode.jsを導入する場合は、利用可能なLTS版のバージョン付きformulaを使用します。次はNode.js 24の例です。

```bash
brew install node@24
```

確認します。

```bash
node --version
npm --version
```

## 5. PDFの生成

最初にMarkdownを検証します。

**Windows PowerShell**

```powershell
py -3 scripts\validate_markdown.py "<document.md>"
```

**macOS**

```bash
python3 scripts/validate_markdown.py "<document.md>"
```

PDFを生成します。

**Windows PowerShell**

```powershell
node scripts\markdown_to_pdf.mjs "<document.md>" "<output.pdf>"
```

**macOS**

```bash
node scripts/markdown_to_pdf.mjs "<document.md>" "<output.pdf>"
```

出力先を省略すると、入力Markdownと同じディレクトリへ同名のPDFを生成します。

```bash
node scripts/markdown_to_pdf.mjs docs/sample.md
```

この例では`docs/sample.pdf`が生成されます。

出力先に指定できるのは`.pdf`ファイルだけです。入力Markdownと同じパスは指定できません。既存の出力ファイルは既定で置き換えず、スクリプトをエラーで終了します。置き換えが必要な場合だけ`--force`を指定してください。`--force`では別の一時ファイルへPDFを完全に生成した後、出力先を置き換えます。

## 6. オプション

```text
--title <text>          PDFメタデータとヘッダーに使うタイトル
--format <size>         A4、Letterなど（既定: A4）
--landscape             横向きで出力
--no-header-footer      ヘッダーとページ番号を表示しない
--theme <default|dark>  Mermaidテーマ（既定: default）
--force                 既存のPDFを一時ファイル経由で安全に置き換える
--help                  ヘルプを表示
```

横向きA4で生成する例です。

```bash
node scripts/markdown_to_pdf.mjs input.md output.pdf --format A4 --landscape
```

既存のPDFを意図的に置き換える例です。

```bash
node scripts/markdown_to_pdf.mjs input.md output.pdf --force
```

## 7. 対応範囲と制約

- CommonMarkを基礎としたMarkdownに対応します。
- パイプ形式の表、画像、SVG参照、コードブロック、Mermaid図をPDFへ反映します。
- ローカル画像とSVGはMarkdownファイルからの相対パスで解決します。`..`またはシンボリックリンクでMarkdown文書のディレクトリ外へ出る参照は使用できません。
- 画像URLのクエリとフラグメントは、参照先ファイルの存在確認時にパスから分離し、PDF用のURLには保持します。
- SVG内の`file:`、相対パス、リモートURL、`data:`による外部参照とCSSの外部`url()`は使用できません。同一SVG内の`#fragment`参照は使用できます。
- セキュリティ上の理由から、Markdown内の生HTMLは無効です。
- Mermaid図が大きすぎる場合は、図を分割するか横向き出力を使用します。
- PDF変換器はすべての見出しにVS Codeとバリデーター互換の一意な`id`を付けるため、PDF内の目次と文書内リンクが機能します。重複slugには`-1`、`-2`の接尾辞が付きますが、文書作成時は重複見出しを避けてください。
- OSに存在しないフォントは代替フォントで描画されるため、WindowsとmacOSで改行位置がわずかに異なる場合があります。

## 8. トラブルシューティング

### 8.1 Chromiumが見つからない

```bash
npx playwright install chromium
```

を再実行します。

### 8.2 Mermaid図が描画できない

まずVS CodeのMarkdownプレビューでMermaid構文を確認し、次にバリデーターを実行します。

```bash
python3 scripts/validate_markdown.py input.md
```

Windowsでは`python3`の代わりに`py -3`を使用します。

### 8.3 画像が見つからない

画像パスは、コマンドを実行したディレクトリではなく、入力Markdownファイルの配置ディレクトリを基準に指定します。

### 8.4 社内プロキシでnpm ciが失敗する

所属組織のnpmレジストリおよびプロキシ設定に従ってください。証明書検証を無効化する設定は推奨しません。
