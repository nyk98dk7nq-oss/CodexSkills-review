---
name: review-markdown-documents
description: Markdownへ変換済みのチェックリスト、参考資料、設計書を突き合わせ、根拠位置付きの判定データと不適合中心のレビューサマリーを作るSkill。設計書レビューのAI判断、結果JSONの検証、summary.md作成に使用する。
---

# Markdown設計書レビュー

入力ファイルを直接レビューせず、形式別Skillが作ったMarkdownだけを判断材料にする。スクリプトはレビュー材料の準備、結果の機械検証、サマリー生成を担当し、意味上の判断はCodexが行う。

## 実施手順

1. `references/result-schema.md` を読む。
2. 統合Skillが作ったmanifestからレビューbundleを作る。

   ```powershell
   py -3 .agents/skills/review-markdown-documents/scripts/review_markdown_documents.py prepare --manifest work/review-runs/<run_id>/manifest.json --output work/review-runs/<run_id>/review_bundle.json
   ```

3. bundleの`materials`にあるMarkdownを、checklist、reference、targetの順にすべて読む。
4. チェックリストの各行について、適用する参考資料と対象ファイルを確認する。ファイル名や文面だけで適用関係を推測しない。適用先を確定できなければ`要確認`にする。
5. `checklist_items × targets`の全組合せを1件ずつ判定し、bundleを別名の`results.json`として完成させる。
6. 結果を検証する。

   ```powershell
   py -3 .agents/skills/review-markdown-documents/scripts/review_markdown_documents.py validate --results work/review-runs/<run_id>/results.json
   ```

7. 統合レビューでは、検証エラーをすべて修正してから`review-documents-orchestrator`へ戻り、そのSkillの`finalize`を実行する。`finalize`が結果チェックリストと`summary.md`をステージングで完成させ、`output/reviews/<run_id>/`へ同時に確定する。この手順では、次の単独用`summary`コマンドを先に実行しない。

8. `review-documents-orchestrator`を使わず、このSkillだけでサマリーを作る場合に限り、次を実行する。この出力先を作成した後は、同じrun-idで統合Skillの`finalize`を実行できない。

   ```powershell
   py -3 .agents/skills/review-markdown-documents/scripts/review_markdown_documents.py summary --results work/review-runs/<run_id>/results.json --output output/reviews/<run_id>/summary.md
   ```

各出力は所定のrun-id別フォルダだけへ新規作成する。既存ファイル、`input/`配下、祖先にsymlinkがあるパス、リポジトリ外へ解決されるパスは拒否し、既存内容を上書きしない。

## 判定規則

- `result`には`適合`、`不適合`、`対象外`、`要確認`だけを使用する。
- チェックリストと参考資料に照らして要求を満たす証拠がある場合だけ`適合`とする。
- 要求との具体的な矛盾または不足を確認できた場合だけ`不適合`とする。
- 適用対象ではない根拠を確認できた場合だけ`対象外`とする。
- 参考資料不足、適用先不明、OCR不明瞭、図だけでは意味を確定できない場合は推測せず`要確認`とする。
- `comment`に判断理由と利用者が確認できる証拠位置を記載する。証拠位置は`evidence`にも同じ表記で列挙する。
- 証拠位置は、例として`work/markdown/<run_id>/targets/api.md:L42-L48`、`Sheet1!B12`、`page 3`、`slide 5`のように特定できる形にする。
- `不適合`には、根拠となる要求を満たすための具体的で実行可能な`improvement`を必ず記載する。単なる「修正してください」や根拠のない全面改訂にしない。
- `要確認`には、不足情報と確認先を`comment`へ明記する。断定的な改善案を作らない。
- OCR文字列は認識結果として扱い、図形、矢印、色、位置関係から仕様を推測しない。

## 完了条件

- すべてのチェック項目と対象ファイルの組合せが重複なく存在する。
- 全件に判定、コメント、証拠位置がある。
- すべての不適合に具体的な改善案がある。
- 検証コマンドが成功し、統合Skillの`finalize`または単独用`summary`で作った`summary.md`に4判定の件数、不適合全件、要確認全件が記載されている。
