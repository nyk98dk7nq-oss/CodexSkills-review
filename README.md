# Codex スキル

Codexで作成したスキルを格納するリポジトリです。

## 目次

1. [1. 収録スキル](#1-収録スキル)
2. [2. 検証](#2-検証)

## 1. 収録スキル

- [`write-vscode-markdown`](./write-vscode-markdown/SKILL.md): VS Code向けに、番号付き見出し、リンク付き目次、表、Mermaid図を備えたMarkdownを作成・検証し、必要に応じてPDFへ変換します。
  - [MarkdownからPDFを生成するための導入要件](./docs/write-vscode-markdown-pdf-setup.md)
- [`analyze-excel-to-markdown`](./analyze-excel-to-markdown/SKILL.md): Excelの構造と内容を解析し、元セルを追跡できるVS Code向けMarkdownを生成します。
  - [導入要件とセットアップ](./docs/analyze-excel-to-markdown-setup.md)
- [`analyze-powerpoint-to-markdown`](./analyze-powerpoint-to-markdown/SKILL.md): PowerPointの構造と内容を解析し、元スライドと図形を追跡できるVS Code向けMarkdownを生成します。
  - [導入要件とセットアップ](./docs/analyze-powerpoint-to-markdown-setup.md)

## 2. 検証

各スキルの回帰テストに加え、ExcelとPowerPointの検査結果からMarkdownを作成し、Markdown検証とPDF生成までを確認するE2Eテストを収録しています。コマンドはリポジトリのルートで実行します。

**Windows PowerShell**

```powershell
py -3 -m unittest discover -s analyze-excel-to-markdown/tests -p "test_*.py" -v
py -3 -m unittest discover -s analyze-powerpoint-to-markdown/tests -p "test_*.py" -v
py -3 -m unittest discover -s write-vscode-markdown/tests -p "test_*.py" -v
py -3 -m unittest discover -s tests -p "test_*.py" -v
```

**macOS**

```bash
python3 -m unittest discover -s analyze-excel-to-markdown/tests -p 'test_*.py' -v
python3 -m unittest discover -s analyze-powerpoint-to-markdown/tests -p 'test_*.py' -v
python3 -m unittest discover -s write-vscode-markdown/tests -p 'test_*.py' -v
python3 -m unittest discover -s tests -p 'test_*.py' -v
```

PDFまで含むE2Eテストは、先に`write-vscode-markdown`で`npm ci`と`npx playwright install chromium`を完了し、環境変数`RUN_PDF_E2E=1`を指定して実行します。環境変数を指定しない場合は、Office文書の解析からMarkdown検証までを実行し、PDF試験だけをスキップします。
