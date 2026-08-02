# analyze-powerpoint-to-markdown 導入ガイド

この文書は、WindowsまたはmacOSでCodexを使用し、`analyze-powerpoint-to-markdown`スキルでPowerPointを解析してVS Code向けMarkdownを生成するための導入要件をまとめたものです。

## 目次

- [1. 利用方針](#1-利用方針)
- [2. 必要なもの](#2-必要なもの)
- [3. Windowsへのインストール](#3-windowsへのインストール)
- [4. macOSへのインストール](#4-macosへのインストール)
- [5. Codexからの動作確認](#5-codexからの動作確認)
- [6. Pythonライブラリの依存関係](#6-pythonライブラリの依存関係)
- [7. 任意で使用するもの](#7-任意で使用するもの)
- [8. インストール不要なもの](#8-インストール不要なもの)
- [9. 対応形式と注意点](#9-対応形式と注意点)
- [10. トラブルシューティング](#10-トラブルシューティング)
- [11. 公式資料](#11-公式資料)

## 1. 利用方針

Python 3.10以上は、Codexを使用するWindowsまたはmacOSへ一度だけインストールし、すべての作業フォルダーから共通利用します。プロジェクトごとのPython仮想環境（`.venv`）は作成しません。

このガイドでいうグローバルインストールは、次の運用を意味します。

1. Python本体をPCへインストールし、Codexを実行するOSユーザーから共通利用する。
2. `python-pptx`は、`pip --user`で同じOSユーザーのユーザーサイトへ一度だけインストールする。
3. 管理者権限を使った`sudo pip install`や、OSまたはパッケージ管理ツールが管理するPythonの強制変更は行わない。

Pythonライブラリは全Codex作業で共有されるため、別のPythonスキルが異なるバージョンを要求する場合は影響を確認してから更新してください。

## 2. 必要なもの

| 種別 | 要件 | 用途 | 追加インストール |
|---|---|---|---|
| Codex実行環境 | WindowsまたはmacOS上で対象スキルとPowerPointファイルを読み取り、ローカルPythonを実行できること | スキルの手順実行とMarkdown生成 | 利用環境に応じて必要 |
| Python | 3.10以上 | PowerPoint解析スクリプトとMarkdownバリデーターの実行 | 必要 |
| pip | 使用するPythonに対応したもの | Pythonライブラリの導入 | 通常はPythonに同梱 |
| python-pptx | 1.0.2以上、2.0未満 | `.pptx`および`.pptm`のスライド、図形、表、ノート、画像、グラフの読取 | 必要 |

Python 3.10以上を要件とするのは、PowerPoint解析に加え、連携先の`write-vscode-markdown`に含まれるMarkdownバリデーターも同じ環境で実行できるようにするためです。`python-pptx`自体の公開要件はPython 3.8以上ですが、このリポジトリではWindowsとmacOSの共通運用をPython 3.10以上へ統一します。

直接インストールする外部Pythonライブラリは`python-pptx`だけです。標準処理でNode.jsやnpmパッケージは使用しません。

## 3. Windowsへのインストール

1. PowerShellで既存のPythonを確認します。

   ```powershell
   py -3 --version
   py -3 -m pip --version
   ```

2. Python 3.10以上が表示されない場合は、[Python公式のWindows向け手順](https://docs.python.org/3/using/windows.html)に従ってPython Install Managerと安定版のPython 3をインストールします。インストール後は新しいPowerShellを開いて、もう一度バージョンを確認します。
3. Codexを実行するWindowsユーザーの共通環境へ`python-pptx`をインストールします。

   ```powershell
   py -3 -m pip install --user --upgrade "python-pptx>=1.0.2,<2"
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
4. Codexを実行するmacOSユーザーの共通環境へ`python-pptx`をインストールします。

   ```bash
   python3 -m pip install --user --upgrade "python-pptx>=1.0.2,<2"
   ```

`externally-managed-environment`が表示されても、`--break-system-packages`や`sudo pip install`は使用しません。仮想環境を使わない今回の運用では、python.org版のPythonを使用します。

## 5. Codexからの動作確認

Windows PowerShellでは、次のコマンドを実行します。

```powershell
py -3 -c "import sys, pptx; v=tuple(map(int, pptx.__version__.split('.')[:3])); assert sys.version_info >= (3, 10); assert (1, 0, 2) <= v < (2, 0, 0); print(sys.executable); print(sys.version); print(pptx.__version__)"
py -3 -m pip check
py -3 ".\analyze-powerpoint-to-markdown\scripts\inspect_powerpoint.py" --help
py -3 ".\analyze-powerpoint-to-markdown\scripts\inspect_powerpoint.py" ".\input.pptx" --output ".\inspection.json"
py -3 ".\write-vscode-markdown\scripts\validate_markdown.py" ".\output.md"
```

macOSでは、次のコマンドを実行します。

```bash
python3 -c "import sys, pptx; v=tuple(map(int, pptx.__version__.split('.')[:3])); assert sys.version_info >= (3, 10); assert (1, 0, 2) <= v < (2, 0, 0); print(sys.executable); print(sys.version); print(pptx.__version__)"
python3 -m pip check
python3 analyze-powerpoint-to-markdown/scripts/inspect_powerpoint.py --help
python3 analyze-powerpoint-to-markdown/scripts/inspect_powerpoint.py input.pptx --output inspection.json
python3 write-vscode-markdown/scripts/validate_markdown.py output.md
```

実ファイルを使うコマンドは、リポジトリのルートで実行する想定です。`input.pptx`、`inspection.json`、`output.md`は実際のファイル名またはパスへ置き換えてください。`--output`は入力PowerPointと異なる`.json`パスだけを受け入れ、既存出力は上書きしません。内容を確認した既存JSONを意図的に置き換える場合だけ`--force`を追加します。

## 6. Pythonライブラリの依存関係

| 区分 | パッケージ | 用途 | 利用者が直接インストール |
|---|---|---|---|
| 直接依存 | `python-pptx` | PowerPoint Open XMLの読取 | 必要 |
| 推移的依存 | `lxml` | Open XMLの解析 | 不要。pipが自動導入 |
| 推移的依存 | `Pillow` | 画像情報の読取 | 不要。pipが自動導入 |
| 推移的依存 | `XlsxWriter` | `python-pptx`のグラフ機能 | 不要。pipが自動導入 |
| 推移的依存 | `typing_extensions` | Python型機能の互換性 | 不要。pipが自動導入 |

これらの推移的依存は、`python-pptx`の公開パッケージ定義に基づきます。個別にバージョンを固定せず、`pip`に解決させます。導入後は`pip check`で依存関係の整合性を確認してください。

解析スクリプトは、ZIP安全検査、ハッシュ計算、JSON出力などにPython標準ライブラリも使用します。標準ライブラリは追加インストール不要です。

## 7. 任意で使用するもの

| ソフトウェアまたはスキル | 必要になる場面 | 備考 |
|---|---|---|
| `write-vscode-markdown`スキル | 番号付き見出し、目次、表、Mermaidを含むMarkdownの整形と検証 | 推奨。追加の外部Pythonライブラリは不要 |
| Visual Studio Code | 生成したMarkdownのプレビュー | 推奨。Mermaidを使用する場合は、その環境で描画できることを確認 |
| Microsoft PowerPoint | `.ppt`等を`.pptx`へ変換する場合、スライドの見た目・アニメーション・ノートを利用者環境で確認する場合 | `.pptx`または`.pptm`の基本解析だけなら不要 |
| LibreOffice Impress | PowerPointがない環境で`.ppt`を`.pptx`へ変換または視覚確認する場合 | 複雑なレイアウトやアニメーションの再現差に注意 |
| Codexのプレゼンテーション表示・レンダリング機能 | グラフ、SmartArt、複数列、図解、重なり、マスター由来要素を視覚確認する場合 | 利用できる環境でのみ使用 |

## 8. インストール不要なもの

次のソフトウェアやパッケージは、このスキルの標準処理では使用しません。

- Node.js
- npmパッケージ
- Mermaid CLI
- Microsoft MarkItDown
- Pandoc
- Jupyter Notebook
- pandas
- openpyxl
- Java
- `pywin32`
- Microsoft PowerPoint（対応済みの`.pptx`または`.pptm`を基本解析するだけの場合）
- Python仮想環境（`.venv`）

Mermaid図はMarkdown内のコードブロックとして生成します。Mermaid CLIで画像へ変換する処理は行いません。

## 9. 対応形式と注意点

| 項目 | 対応状況 |
|---|---|
| `.pptx` | 対応 |
| `.pptm` | 読み取り専用で対応。VBAマクロは実行・抽出・保存しない |
| `.ppt` / `.pps` / `.pot` | 非対応。事前に`.pptx`へ変換する |
| `.ppsx` / `.ppsm` / `.potx` / `.potm` | 標準対象外。`.pptx`へ変換する |
| 暗号化・パスワード保護 | 非対応。スキルは解除を試みない |
| インターネット接続 | Pythonとライブラリの導入時には必要。導入後のローカル解析には原則不要 |

元のPowerPointは編集しません。ZIP部品名、展開量、内部Relationship、DTD、XMLエンティティを事前検査し、検査済みのバイト列だけをメモリから`python-pptx`へ渡します。Relationship Typeは既知の固定分類だけを集計し、任意の文字列は`other`へまとめます。`python-pptx`の保存APIは使用しません。

Picture図形の画像情報は、図形の`r:embed`または`r:link`からスライドのRelationship先と`[Content_Types].xml`を検証済みZIP内で直接対応付けます。このため`python-pptx`がWebPやSVGを画像としてデコードできない場合も、内部部品を安全に特定できればファイル名、Content-Type、拡張子、バイト数、SHA-256、抽出可否、未抽出理由をJSONへ記録します。外部Relationshipのリンク先は開かず、不正な内部targetは拒否します。

`--extract-images`で自動抽出するのは、PNG、JPEG、GIF、BMP、TIFF、WebP、APNGのうち、拡張子、Content-Type、ファイル署名が一致するラスター画像に限ります。SVG、EMF、WMFなどのベクター画像や未認識・形式不一致の部品は、安全なサニタイズを行わずにVS CodeのMarkdownプレビューへ渡さないため、メタデータとハッシュだけを記録して自動抽出しません。既存の抽出画像は`--force`指定の有無にかかわらず上書きしません。

2026年6月に、`python-pptx 1.0.2`までのディレクトリ読取とZIP書込に関するパストラバーサル報告が公開されています。このスキルは通常ファイルのZIPだけを受け入れ、部品名とRelationshipを検査し、ディレクトリ入力や保存処理を使用しないことで、報告された経路を避けます。未知または信頼できないPowerPointでは、上限値を理由なく引き上げないでください。

非表示スライド、発表者ノート、非表示図形、完全にスライド領域外の図形、文書プロパティは既定で内容を除外します。グループ図形では親子の座標変換を合成し、回転後の四角形とスライド矩形の交差で領域を判定します。座標を正規化できない子図形も既定では内容を除外します。SmartArt、OLE、ActiveX、音声、動画、コメント、アニメーション、マスター由来の要素は、存在を検出できても完全には解析しません。

JSONは出力先と同じディレクトリの一時ファイルへ完全に書き込んでから原子的に公開します。新規出力は既存パスを競合なく拒否し、`--force`指定時だけ既存JSONを原子的に置き換えます。画像抽出を指定した場合、JSON出力先と生成画像のパスが衝突すれば書き込み前に停止します。JSONの保存に失敗した場合は、その実行で新規作成した画像を削除します。

## 10. トラブルシューティング

| 症状 | 確認と対処 |
|---|---|
| Windowsで`py`が見つからない | Python Install Managerを導入し、新しいPowerShellで`py -3 --version`を確認する |
| macOSで`python3`が見つからない、3.10未満、または`/usr/bin/python3`を示す | python.org版のPython 3.10以上を導入し、新しいターミナルでパスを再確認する |
| `No module named 'pptx'` | Windowsは`py -3 -m pip install --user --upgrade "python-pptx>=1.0.2,<2"`、macOSは`python3 -m pip install --user --upgrade "python-pptx>=1.0.2,<2"`を実行する |
| `python-pptx 1.0.2以上2.0未満が必要`と表示される | OSに応じた同じPythonコマンドで`python-pptx`を更新する |
| macOSで`externally-managed-environment`と表示される | `--break-system-packages`を使用せず、python.org版のPythonを導入して同じコマンドを実行する |
| Python更新後に`pptx`が見つからない | 新しく使用するPythonに対して、OS別の`--user`インストールを再実行する |
| `.ppt`、`.pps`、`.pot`を解析できない | PowerPointまたはLibreOffice Impressで`.pptx`として保存し直す |
| `--outputには拡張子.json`と表示される | 出力先を入力PowerPointとは別の`.json`ファイルへ変更する |
| JSONの出力先が既に存在すると表示される | 別名を指定する。既存JSONの置換が意図的な場合だけ内容を確認して`--force`を付ける |
| SVG、EMF、WMFまたは形式不一致画像が抽出されない | 仕様どおり。検査JSONの`image.sha256`、`extractable`、`not_extracted_reason`を確認し、安全な変換は別工程で行う |
| スライドの図解や読み順が正しく再現されない | 元スライドを視覚確認し、検査JSONの座標と図形参照を根拠にMarkdownを修正する |
| 非表示スライドやノートがない | 必要な場合だけ`--include-hidden-slides`または`--include-notes`を明示して再解析する |
| ZIP展開量、部品数、文字数などの上限で停止する | 対象スライドを`--slide`で絞る。上限変更はファイルの安全性と必要量を確認してから行う |

## 11. 公式資料

- [WindowsでPythonを使用する](https://docs.python.org/3/using/windows.html)
- [macOSでPythonを使用する](https://docs.python.org/3/using/mac.html)
- [pipの`--user`オプション](https://pip.pypa.io/en/stable/cli/pip_install/#cmdoption-user)
- [外部管理されたPython環境](https://packaging.python.org/en/latest/specifications/externally-managed-environments/)
- [python-pptxの導入](https://python-pptx.readthedocs.io/en/latest/user/install.html)
- [python-pptxのPyPI情報](https://pypi.org/project/python-pptx/)
- [python-pptxの依存パッケージ定義](https://github.com/scanny/python-pptx/blob/master/pyproject.toml)
- [python-pptxのノートAPI](https://python-pptx.readthedocs.io/en/latest/user/notes.html)
- [python-pptxのShape API](https://python-pptx.readthedocs.io/en/latest/api/shapes.html)
- [python-pptxのパストラバーサル報告](https://github.com/scanny/python-pptx/issues/1137)
