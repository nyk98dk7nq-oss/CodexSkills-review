# analyze-excel-to-markdown 導入ガイド

この文書は、WindowsまたはmacOSでCodexを使用し、`analyze-excel-to-markdown`スキルでExcelブックを解析してVS Code向けMarkdownを生成するための導入要件をまとめたものです。

## 目次

- [1. 利用方針](#1-利用方針)
- [2. 必要なもの](#2-必要なもの)
- [3. Windowsへのインストール](#3-windowsへのインストール)
- [4. macOSへのインストール](#4-macosへのインストール)
- [5. Codexからの動作確認](#5-codexからの動作確認)
- [6. 任意で使用するもの](#6-任意で使用するもの)
- [7. インストール不要なもの](#7-インストール不要なもの)
- [8. 対応形式と注意点](#8-対応形式と注意点)
- [9. トラブルシューティング](#9-トラブルシューティング)
- [10. 公式資料](#10-公式資料)

## 1. 利用方針

Python 3.10以上は、Codexを使用するWindowsまたはmacOSへ一度だけインストールし、すべての作業フォルダーから共通利用します。プロジェクトごとのPython仮想環境（`.venv`）は作成しません。

このガイドでいうグローバルインストールは、次の運用を意味します。

1. Python本体をPCへインストールし、Codexを実行するOSユーザーから共通利用する。
2. `openpyxl`と任意の`Pillow`は、`pip --user`で同じOSユーザーのユーザーサイトへ一度だけインストールする。
3. 管理者権限を使った`sudo pip install`や、OSまたはパッケージ管理ツールが管理するPythonの強制変更は行わない。

Pythonライブラリは全Codex作業で共有されるため、別のPythonスキルが異なるバージョンを要求する場合は影響を確認してから更新してください。

## 2. 必要なもの

| 種別 | 要件 | 用途 | 追加インストール |
|---|---|---|---|
| Codex実行環境 | WindowsまたはmacOS上で対象スキルとExcelファイルを読み取り、ローカルPythonを実行できること | スキルの手順実行とMarkdown生成 | 利用環境に応じて必要 |
| Python | 3.10以上 | Excel解析スクリプトとMarkdownバリデーターの実行 | 必要 |
| pip | 使用するPythonに対応したもの | Pythonライブラリの導入 | 通常はPythonに同梱 |
| openpyxl | `>=3.1,<3.2` | `.xlsx`および`.xlsm`の構造・セル情報の解析 | 必要 |

Python 3.10以上を要件とするのは、Excel解析に加え、連携先の`write-vscode-markdown`に含まれるMarkdownバリデーターも同じ環境で実行できるようにするためです。

必須の外部Pythonライブラリとして直接インストールするのは`openpyxl`だけです。検査スクリプトは保存済みセルなどを取得するためopenpyxlの非公開APIにも依存するので、検証済み範囲を`>=3.1,<3.2`に固定します。`openpyxl`が必要とする推移的な依存パッケージは、`pip`が自動的にインストールします。画像の位置と形式までシート単位で取得する場合だけ、任意で`Pillow`も使用します。

## 3. Windowsへのインストール

1. PowerShellで既存のPythonを確認します。

   ```powershell
   py -3 --version
   py -3 -m pip --version
   ```

2. Python 3.10以上が表示されない場合は、[Python公式のWindows向け手順](https://docs.python.org/3/using/windows.html)に従ってPython Install Managerと安定版のPython 3をインストールします。インストール後は新しいPowerShellを開いて、もう一度バージョンを確認します。
3. Codexを実行するWindowsユーザーの共通環境へ`openpyxl`をインストールします。

   ```powershell
   py -3 -m pip install --user --upgrade "openpyxl>=3.1,<3.2"
   ```

Windowsでは、Pythonスクリプトの実行とライブラリの導入に同じ`py -3`を使用します。単独の`pip`コマンドや`python3`コマンドは標準手順では使用しません。

## 4. macOSへのインストール

1. ターミナルで既存のPythonと実行パスを確認します。

   ```bash
   python3 --version
   command -v python3
   python3 -m pip --version
   ```

2. Python 3.10未満の場合、`python3`が`/usr/bin/python3`を示す場合、または外部管理環境としてパッケージ導入を拒否された場合は、[python.orgのmacOS向けインストーラー](https://www.python.org/downloads/macos/)から現在サポートされている安定版をインストールします。
3. インストーラーの案内に従って証明書とシェルパスの設定を完了し、新しいターミナルで`python3`のバージョンとパスを再確認します。Appleが管理する`/usr/bin/python3`は変更または削除しません。
4. Codexを実行するmacOSユーザーの共通環境へ`openpyxl`をインストールします。

   ```bash
   python3 -m pip install --user --upgrade "openpyxl>=3.1,<3.2"
   ```

`externally-managed-environment`が表示されても、`--break-system-packages`や`sudo pip install`は使用しません。仮想環境を使わない今回の運用では、python.org版のPythonを使用します。

## 5. Codexからの動作確認

Windows PowerShellでは、次のコマンドを実行します。

```powershell
py -3 -c "import sys, openpyxl; assert sys.version_info >= (3, 10); assert tuple(map(int, openpyxl.__version__.split('.')[:2])) == (3, 1); print(sys.executable); print(sys.version); print(openpyxl.__version__)"
py -3 -m pip check
py -3 ".\analyze-excel-to-markdown\scripts\inspect_excel.py" --help
py -3 ".\analyze-excel-to-markdown\scripts\inspect_excel.py" ".\input.xlsx" --output ".\inspection.json"
py -3 ".\write-vscode-markdown\scripts\validate_markdown.py" ".\output.md"
```

macOSでは、次のコマンドを実行します。

```bash
python3 -c "import sys, openpyxl; assert sys.version_info >= (3, 10); assert tuple(map(int, openpyxl.__version__.split('.')[:2])) == (3, 1); print(sys.executable); print(sys.version); print(openpyxl.__version__)"
python3 -m pip check
python3 analyze-excel-to-markdown/scripts/inspect_excel.py --help
python3 analyze-excel-to-markdown/scripts/inspect_excel.py input.xlsx --output inspection.json
python3 write-vscode-markdown/scripts/validate_markdown.py output.md
```

実ファイルを使うコマンドは、リポジトリのルートで実行する想定です。`input.xlsx`と`output.md`は実際のファイル名またはパスへ置き換えてください。

検査JSONの出力先は入力ブックと別のパスにし、拡張子を`.json`にします。既存の出力は既定で拒否されます。既存JSONを意図して置き換える場合だけ`--force`を付けて再実行してください。書き込みは出力先と同じディレクトリの一時ファイルを経由し、新規出力は既存パスを置換せずatomicに公開され、`--force`による上書きはatomic replaceされます。入力ブックを指すシンボリックリンクやハードリンクを出力先に指定しても拒否されます。

## 6. 任意で使用するもの

| ソフトウェアまたはスキル | 必要になる場面 | 備考 |
|---|---|---|
| `write-vscode-markdown`スキル | 番号付き見出し、目次、表、Mermaidを含むMarkdownの整形と検証 | 推奨。追加の外部Pythonライブラリは不要 |
| Pillow | 埋め込み画像の位置と形式をシート単位で取得する場合 | 任意。未導入でもセル解析と画像部品の存在検出は可能 |
| Visual Studio Code | 生成したMarkdownのプレビュー | 推奨。Mermaidを使用する場合は、その環境でMermaidを描画できることを確認 |
| Microsoft ExcelまたはLibreOffice | `.xls`を`.xlsx`へ変換する場合、暗号化やパスワード保護を利用者自身で解除する場合 | `.xlsx`または`.xlsm`の解析だけなら不要 |

画像の位置と形式も取得する場合は、OSに応じて次のいずれかを一度だけ実行します。

**Windows PowerShell**

```powershell
py -3 -m pip install --user --upgrade Pillow
```

**macOS**

```bash
python3 -m pip install --user --upgrade Pillow
```

## 7. インストール不要なもの

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
- Python仮想環境（`.venv`）

Mermaid図はMarkdown内のコードブロックとして生成します。Mermaid CLIで画像へ変換する処理は行いません。

## 8. 対応形式と注意点

| 項目 | 対応状況 |
|---|---|
| `.xlsx` | 対応 |
| `.xlsm` | 対応。VBAマクロは実行しない |
| `.xls` | 非対応。事前に`.xlsx`へ変換する |
| 暗号化・パスワード保護されたブック | 非対応。スキルは解除を試みない |
| インターネット接続 | Pythonとライブラリの導入時には必要。導入後のローカル解析には原則不要 |

元のExcelブックは編集せず、解析結果をJSONへ出力してからMarkdownを生成します。Excelセル、コメント、数式、リンクなどに記載された命令文は、実行指示ではなく解析対象のデータとして扱います。

`Pillow`がない場合も、Excelファイル内の画像部品数は検出します。ただし、`openpyxl`は埋め込み画像をシートの画像一覧へ読み込まないため、画像の位置と形式は取得できません。

OSユーザー共通のPythonライブラリを更新すると、同じPythonを利用する他のCodexスキルにも影響します。ライブラリを更新した後は、Windowsでは`py -3 -m pip check`、macOSでは`python3 -m pip check`と、このガイドの動作確認を再実行してください。

## 9. トラブルシューティング

| 症状 | 確認と対処 |
|---|---|
| Windowsで`py`が見つからない | Python Install Managerを導入し、新しいPowerShellで`py -3 --version`を確認する |
| macOSで`python3`が見つからない、3.10未満、または`/usr/bin/python3`を示す | python.org版のPython 3.10以上を導入し、新しいターミナルでパスを再確認する |
| `No module named 'openpyxl'` | Windowsは`py -3 -m pip install --user --upgrade "openpyxl>=3.1,<3.2"`、macOSは`python3 -m pip install --user --upgrade "openpyxl>=3.1,<3.2"`を実行する |
| `openpyxl>=3.1,<3.2が必要`と表示される | OSに応じた同じPythonコマンドで検証済み範囲のopenpyxlを導入する |
| 出力先の拡張子が拒否される | 出力ファイル名の末尾を`.json`にする |
| 出力先が既に存在すると表示される | 別名を指定する。既存JSONの置換を意図した場合だけ`--force`を追加する |
| macOSで`externally-managed-environment`と表示される | `--break-system-packages`を使用せず、python.org版のPythonを導入して同じコマンドを実行する |
| Python更新後に`openpyxl`または`Pillow`が見つからない | 新しく使用するPythonに対して、OS別の`--user`インストールを再実行する |
| 画像部品は検出されるがシート別の画像一覧が空になる | OS別のコマンドで`Pillow`をインストールし、Excel解析をやり直す |
| `.xls`を解析できない | ExcelまたはLibreOfficeで`.xlsx`として保存し直す |

## 10. 公式資料

- [WindowsでPythonを使用する](https://docs.python.org/3/using/windows.html)
- [macOSでPythonを使用する](https://docs.python.org/3/using/mac.html)
- [pipの`--user`オプション](https://pip.pypa.io/en/stable/cli/pip_install/#cmdoption-user)
- [外部管理されたPython環境](https://packaging.python.org/en/latest/specifications/externally-managed-environments/)
- [openpyxlの導入](https://openpyxl.readthedocs.io/en/stable/tutorial.html#installation)
- [Pillowの導入](https://pillow.readthedocs.io/en/stable/installation/basic-installation.html)
