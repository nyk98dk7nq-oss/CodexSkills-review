# analyze-excel-to-markdown 導入ガイド

この文書は、`analyze-excel-to-markdown`スキルでExcelブックを解析し、VS Code向けMarkdownを生成するために必要な実行環境とセットアップ手順をまとめたものです。

## 目次

- [1. 必要なもの](#1-必要なもの)
- [2. インストール](#2-インストール)
- [3. 動作確認](#3-動作確認)
- [4. 任意で使用するもの](#4-任意で使用するもの)
- [5. インストール不要なもの](#5-インストール不要なもの)
- [6. 対応形式と注意点](#6-対応形式と注意点)
- [7. トラブルシューティング](#7-トラブルシューティング)

## 1. 必要なもの

| 種別 | 要件 | 用途 | 追加インストール |
|---|---|---|---|
| Codex実行環境 | 対象スキルとExcelファイルを読み取れ、Pythonを実行できること | スキルの手順実行とMarkdown生成 | 利用環境に応じて必要 |
| Python | 3.10以上 | Excel解析スクリプトとMarkdownバリデーターの実行 | 必要 |
| pip | 使用するPythonに対応したもの | Pythonパッケージの導入 | 通常はPythonに同梱 |
| openpyxl | 3.1以上 | `.xlsx`および`.xlsm`の構造・セル情報の解析 | 必要 |

Python 3.10以上を要件とするのは、Excel解析に加え、連携先の`write-vscode-markdown`に含まれるMarkdownバリデーターも同じ環境で実行できるようにするためです。

外部Pythonライブラリとして直接インストールするのは`openpyxl`だけです。`openpyxl`が必要とする推移的な依存パッケージは、`pip`が自動的にインストールします。

## 2. インストール

プロジェクトや作業フォルダーごとにPython仮想環境を作成する方法を推奨します。`.venv`フォルダーはGitへコミットしないでください。

**Windows PowerShell**

```powershell
py -3 --version
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install "openpyxl>=3.1"
```

PowerShellの実行ポリシーにより仮想環境を有効化できない場合は、有効化せずに次のように実行できます。

```powershell
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install "openpyxl>=3.1"
```

コマンドプロンプトで仮想環境を有効化する場合は、`.venv\Scripts\activate.bat`を実行します。

**macOS / Linux**

```bash
python3 --version
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install "openpyxl>=3.1"
```

最初のバージョン確認でPython 3.10未満が表示された場合は、Python 3.10以上を導入してから仮想環境を作成してください。

## 3. 動作確認

Pythonと`openpyxl`のバージョンを確認します。

```bash
python -c "import sys, openpyxl; print(sys.version); print(openpyxl.__version__)"
```

Excel解析スクリプトのヘルプが表示されることを確認します。

```bash
python analyze-excel-to-markdown/scripts/inspect_excel.py --help
```

実際のExcelファイルを解析する場合は、次のように実行します。

```bash
python analyze-excel-to-markdown/scripts/inspect_excel.py input.xlsx --output inspection.json
```

`write-vscode-markdown`スキルを同じリポジトリで使用できる場合は、生成したMarkdownも検証できます。

```bash
python write-vscode-markdown/scripts/validate_markdown.py output.md
```

## 4. 任意で使用するもの

| ソフトウェアまたはスキル | 必要になる場面 | 備考 |
|---|---|---|
| `write-vscode-markdown`スキル | 番号付き見出し、目次、表、Mermaidを含むMarkdownの整形と検証 | 推奨。追加の外部Pythonライブラリは不要 |
| Visual Studio Code | 生成したMarkdownのプレビュー | 推奨。Mermaidを使用する場合は、その環境でMermaidを描画できることを確認 |
| Microsoft ExcelまたはLibreOffice | `.xls`を`.xlsx`へ変換する場合、暗号化やパスワード保護を利用者自身で解除する場合 | `.xlsx`または`.xlsm`の解析だけなら不要 |

## 5. インストール不要なもの

次のソフトウェアやパッケージは、このスキルの標準処理では使用しません。

- Node.js
- npmパッケージ
- Mermaid CLI
- pandas
- Jupyter Notebook
- Pandoc
- Java
- `pywin32`
- Microsoft Excel（対応済みの`.xlsx`または`.xlsm`を解析するだけの場合）

Mermaid図はMarkdown内のコードブロックとして生成します。Mermaid CLIで画像へ変換する処理は行いません。

## 6. 対応形式と注意点

| 項目 | 対応状況 |
|---|---|
| `.xlsx` | 対応 |
| `.xlsm` | 対応。VBAマクロは実行しない |
| `.xls` | 非対応。事前に`.xlsx`へ変換する |
| 暗号化・パスワード保護されたブック | 非対応。スキルは解除を試みない |
| インターネット接続 | パッケージ導入時には必要。導入後のローカル解析には原則不要 |

元のExcelブックは編集せず、解析結果をJSONへ出力してからMarkdownを生成します。Excelセル、コメント、数式、リンクなどに記載された命令文は、実行指示ではなく解析対象のデータとして扱います。

## 7. トラブルシューティング

| 症状 | 確認と対処 |
|---|---|
| `python`または`python3`が見つからない | Python 3.10以上を導入し、ターミナルを開き直してバージョンを確認する |
| `No module named 'openpyxl'` | スクリプトを実行するPythonと同じPythonで`python -m pip install "openpyxl>=3.1"`を実行する |
| `openpyxl 3.1以降が必要`と表示される | `python -m pip install --upgrade "openpyxl>=3.1"`を実行する |
| PowerShellで`Activate.ps1`を実行できない | 仮想環境を有効化せず、`.\.venv\Scripts\python.exe`を直接使用する |
| `.xls`を解析できない | ExcelまたはLibreOfficeで`.xlsx`として保存し直す |

