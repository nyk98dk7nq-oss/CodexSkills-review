# Project Skills

このフォルダは、設計書レビューで使用するプロジェクト固有の Skill を格納する場所です。Codex は、[公式の配置規則](https://developers.openai.com/codex/build-skills)に従い、リポジトリルートの `.agents/skills/` から Skill を読み込みます。

Skill Creator で作成・検証する Skill は次の8つです。

- `xlsx-document`
- `docx-document`
- `pptx-document`
- `pdf-document`
- `image-document`
- `convert-legacy-office`
- `review-markdown-documents`
- `review-documents-orchestrator`

各 Skill は `.agents/skills/<skill-name>/` に置き、`SKILL.md`、`agents/openai.yaml`、必要な `scripts/`、`references/`、`assets/`、`tests/` だけで構成します。個別の Skill フォルダには追加の `README.md` を作成しません。

利用者のチェックリスト、参考資料、レビュー対象はこのフォルダへ置かず、`input/` 配下の対応するフォルダへ置いてください。
