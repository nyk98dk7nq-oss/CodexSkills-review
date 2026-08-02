---
name: write-vscode-markdown
description: VS Codeのプレビューに最適化したMarkdown文書を作成・改訂し、必要に応じてPDFへ変換する。CommonMarkを基礎に、VS Code互換のパイプ形式の表、番号付き見出し、リンク付き目次、目的に応じて選択したMermaid図を使用する。手順書、仕様書、アーキテクチャメモ、ワークフロー、データモデル、解説ガイドなど、構造化、ナビゲーション、視認性、機械検証、PDF納品が必要な`.md`文書を扱う場合に使用する。
---

# VS Code向けMarkdown作成

## ワークフローに従う

1. 文書の目的、対象読者、出力先、必要な詳細度、PDF出力の要否を特定する。
2. 文書を新規作成または再構成する前に、[references/vscode-markdown-profile.md](references/vscode-markdown-profile.md)を読む。
3. 文書に図やチャートが必要な場合は、[references/mermaid-design-rules.md](references/mermaid-design-rules.md)を読む。
4. 編集前に既存の内容と参照先アセットを確認する。ユーザーが書き直しを求めていない限り、正しい内容を保持する。
5. 文書の構成を決めてから、タイトル、リンク付き目次、番号付き見出し、本文、表、図を作成する。
6. 単純な説明には文章、正確な比較にはパイプ形式の表を優先する。関係、順序、状態、構造、責任分担が視覚的に分かりやすくなる場合だけMermaidを使用する。
7. XMLとSVGのソースはフェンス付きコードブロックで記述する。表示するSVGは外部ファイルとして参照する。生HTMLやインラインSVGは出力しない。
8. プロジェクトごとの仮想環境は作成せず、OSへ共通導入されたPython 3.10以上でバリデーターを実行する。見出し、目次と通常のリンク、パイプ形式の表の列数、Mermaid、HTML、画像パス、SVGの外部参照に関するエラーをすべて修正する。Windowsでは`py -3`、macOSでは`python3`を使う。

   **Windows PowerShell**

   ```powershell
   py -3 "<skill-dir>\scripts\validate_markdown.py" "<document.md>"
   ```

   **macOS**

   ```bash
   python3 "<skill-dir>/scripts/validate_markdown.py" "<document.md>"
   ```

9. PDF出力が必要な場合は、[PDF生成の導入要件](../docs/write-vscode-markdown-pdf-setup.md)を確認する。初回のみスキルディレクトリでロックファイルを使う`npm ci`と`npx playwright install chromium`を実行する。`npm install`は依存関係を意図的に更新する場合だけ使用する。Markdown検証の合格後、次のコマンドでPDFを生成する。

   **Windows PowerShell**

   ```powershell
   node "<skill-dir>\scripts\markdown_to_pdf.mjs" "<document.md>" "<output.pdf>"
   ```

   **macOS**

   ```bash
   node "<skill-dir>/scripts/markdown_to_pdf.mjs" "<document.md>" "<output.pdf>"
   ```

   PDF変換ではローカルの`markdown-it`、`mermaid`、`highlight.js`、Playwright Chromiumを使用する。外部CDNへ依存しない。出力先は`.pdf`に限り、既存ファイルを既定で上書きしない。意図的に置き換える場合だけ`--force`を追加する。横長の表や図が多い場合は`--landscape`を使用する。
10. 利用可能な場合は、結果をVS Codeでプレビューする。バリデーターの成功は構造上の検証結果であり、すべてのMermaid図が正しく描画されることの保証ではないと扱う。PDFを生成した場合は、PDF内の改ページ、表、図、画像、ヘッダー、ページ番号も確認する。
11. 見出し、リンク、アセットのパス、表、Mermaidを変更するたびに検証を再実行し、PDF納品が必要な場合はPDFも再生成する。

## 適切な粒度でレビューする

通常の作業は、メインエージェントと`validate_markdown.py`で完結させる。

複雑、重要、または大規模な文書では、ユーザーが不要とした場合やサブエージェントを利用できない場合を除き、機械検証の合格後にサブエージェントへ独立した内容レビューを依頼する。レビュー担当には、完成した成果物と、対象読者および目的を判断するために必要な最小限のタスク固有情報だけを渡す。期待する指摘、疑わしい不具合、修正案は伝えない。構成、明確さ、不足情報、図の選択、視覚的な読みやすさ、本文と図の整合性を評価させる。見出し番号、リンク、コードフェンス、禁止HTML、ファイルパスの再確認だけを目的にサブエージェントを使用しない。

メインエージェントが指摘を評価し、採用した修正をすべて反映してから、検証とプレビューを再実行する。PDFを納品する場合は、最終修正後にPDFを再生成する。

## 文書を納品する

次を確認する。

- H1タイトルと`## 目次`見出しに番号が付いていない。
- その他のH2からH4までの見出しに一貫した番号が付いている。
- 目次リンクがVS Codeで正しく機能する。
- 目次外の文書内アンカーと相対リンクの参照先が存在する。
- 生HTMLを使わなくても表を読み取れる。
- パイプ形式の表のヘッダー、区切り行、各データ行の列数が一致する。
- Mermaid図が適切な図種を使用し、本文を補う情報を示している。
- SVG画像がMarkdown文書のディレクトリ内の相対パスで参照され、参照先ファイルが存在し、外部リソース参照を含まない。
- バリデーターが正常に完了する。
- PDFが必要な場合、Mermaid図、表、画像、コードブロック、ページ番号が正しく描画されている。

通常は完成したMarkdownファイルと、そこから参照する外部SVGアセットを返す。PDF出力を依頼された場合はPDFも返す。確認できなかったプレビュー上またはPDF上の制約がある場合は明記する。
