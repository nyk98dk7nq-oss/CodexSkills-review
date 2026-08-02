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
| Node.js 20以上 | 必須 | PDF変換スクリプトの実行 |
| npm | 必須 | JavaScriptライブラリの導入 |
| Playwright Chromium | 必須 | HTMLの描画とPDF出力 |
| Python 3.10以上 | 推奨 | `validate_markdown.py`による事前検証 |
| VS Code | 任意 | MarkdownとMermaidのプレビュー |

このPDF変換機能ではPythonライブラリを追加インストールしません。Pythonは既存のMarkdownバリデーターにのみ使用します。

利用するnpmパッケージは次のとおりです。

| パッケージ | 用途 |
|---|---|
| `markdown-it` | MarkdownからHTMLへの変換 |
| `mermaid` | Mermaid図の描画 |
| `playwright` | ChromiumによるPDF生成 |
| `highlight.js` | コードブロックのシンタックスハイライト |

## 3. Windowsでのセットアップ

PowerShellでリポジトリのスキルディレクトリへ移動します。

```powershell
cd "<repository>\write-vscode-markdown"
npm install
npx playwright install chromium
```

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
npm install
npx playwright install chromium
```

HomebrewでNode.jsを導入する場合は次を実行します。

```bash
brew install node
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

## 6. オプション

```text
--title <text>          PDFメタデータとヘッダーに使うタイトル
--format <size>         A4、Letterなど（既定: A4）
--landscape             横向きで出力
--no-header-footer      ヘッダーとページ番号を表示しない
--theme <default|dark>  Mermaidテーマ（既定: default）
--help                  ヘルプを表示
```

横向きA4で生成する例です。

```bash
node scripts/markdown_to_pdf.mjs input.md output.pdf --format A4 --landscape
```

## 7. 対応範囲と制約

- CommonMarkを基礎としたMarkdownに対応します。
- パイプ形式の表、画像、SVG参照、コードブロック、Mermaid図をPDFへ反映します。
- ローカル画像とSVGはMarkdownファイルからの相対パスで解決します。
- セキュリティ上の理由から、Markdown内の生HTMLは無効です。
- Mermaid図が大きすぎる場合は、図を分割するか横向き出力を使用します。
- PDF内の目次リンクは、Markdown側で生成済みの見出しリンクを利用します。
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

### 8.4 社内プロキシでnpm installが失敗する

所属組織のnpmレジストリおよびプロキシ設定に従ってください。証明書検証を無効化する設定は推奨しません。
