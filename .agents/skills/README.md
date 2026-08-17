# Project Skills

このフォルダには、設計書レビューで使用する実装・検証済みのプロジェクト固有 Skill を格納しています。Codex は、[OpenAI 公式の配置規則](https://developers.openai.com/codex/build-skills)に従い、リポジトリルートの `.agents/skills/` から Skill を読み込みます。

レビュー全体を実行するときは `review-documents-orchestrator` を使用します。形式別の読み取り・Markdown 変換・編集は、次の各 Skill が担当します。

| Skill | 役割 |
|---|---|
| `xlsx-document` | XLSX の読み取り、Markdown 変換、編集、レビュー結果記入 |
| `docx-document` | DOCX の読み取り、Markdown 変換、編集 |
| `pptx-document` | PPTX の読み取り、Markdown 変換、編集 |
| `pdf-document` | PDF の読み取り、OCR、Markdown 変換、ページ編集 |
| `image-document` | 画像の読み取り、OCR、Markdown 変換、編集 |
| `convert-legacy-office` | Microsoft Office COM による旧形式変換 |
| `review-markdown-documents` | Markdown 文書のレビュー、結果検証、サマリー作成 |
| `review-documents-orchestrator` | 環境確認から成果物作成までの統合処理 |

各 Skill は `SKILL.md`、`agents/openai.yaml`、再利用可能な `scripts/`、必要時だけ読む `references/`、`tests/` で構成します。個別の Skill フォルダには補助的な `README.md` を置きません。

利用者のチェックリスト、参考資料、レビュー対象はこのフォルダへ置かず、`input/` 配下の対応するフォルダへ置いてください。
