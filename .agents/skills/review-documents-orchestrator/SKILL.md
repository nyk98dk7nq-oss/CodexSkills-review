---
name: review-documents-orchestrator
description: チェックリスト、参考資料、設計書、画像を使う一連の設計書レビューを統合するSkill。環境確認、旧Office変換、形式別Markdown化、OCR、AIレビュー、結果記入済みチェックリスト、不適合中心のsummary.md作成までを順序どおり実行するときに最初に使用する。
---

# 設計書レビュー統合

このリポジトリでレビューを始めるときは最初に本Skillを使う。入力を上書きせず、決定論的な変換処理とCodexによる意味上のレビューを分ける。

## 1. 事前確認

リポジトリルートで実行する。

```powershell
py -3 .agents/skills/review-documents-orchestrator/scripts/preflight.py --root .
```

Python 3.12以上、入力形式に必要なPythonライブラリ、Tesseractと言語データ、旧Office処理に必要なMicrosoft Office COMを検査する。失敗したらレビューを開始せず、表示された不足項目とルートREADMEのインストール手順を利用者へ案内する。

## 2. レビュー材料の準備

```powershell
py -3 .agents/skills/review-documents-orchestrator/scripts/orchestrate.py prepare --root .
```

必要なら`--run-id yyyyMMddhhmm`を指定する。コマンドは次を実施する。

1. `input/checklists/`、`input/references/`、`input/targets/`を確認する。
2. チェックリストまたは対象が空なら停止する。
3. 入力拡張子に対応する形式別Skillと、必要なlegacy/image Skillの`SKILL.md`を実行前に最後まで読む。
4. `.xls`、`.doc`、`.ppt`を`work/converted-office/`へ変換する。
5. 各形式Skillを使って全入力を`work/markdown/`へ変換する。
6. 画像と画像PDFは形式別SkillからOCRを実施する。
7. `work/review-runs/<run_id>/manifest.json`と`review_bundle.json`を作る。

出力されたbundleをコピーして同じフォルダに`results.json`を作る。

## 3. AIレビュー

`review-markdown-documents/SKILL.md`を最後まで読み、記載された規則に従う。bundleの全Markdownを読み、`items`の全組合せを埋める。

- 適用関係が明示されない内容を推測しない。
- 判定は`適合`、`不適合`、`対象外`、`要確認`だけにする。
- 全コメントに利用者が確認できる証拠位置を含める。
- 全不適合に根拠に沿った具体的で実行可能な改善案を作る。

## 4. 結果の確定

```powershell
py -3 .agents/skills/review-documents-orchestrator/scripts/orchestrate.py finalize --root . --run-id <run_id> --results work/review-runs/<run_id>/results.json
```

このコマンドは結果を検証してから、`output/reviews/<run_id>/`へ次を作る。

- 入力ごとの3列（`レビュー対象ファイル名`、`レビュー結果`、`レビューコメント`）を追加したチェックリストのコピー
- 4判定の件数、全不適合と改善案、要確認を記載した`summary.md`

`input/`の原本は上書きしない。同じ実行IDの`work/review-runs/`、`work/markdown/`、`work/images/`、`work/converted-office/`、`output/reviews/`のいずれかが既にある場合は、空でも再利用せず停止する。対象ファイルそのものを編集するのは、利用者が明示的に依頼した場合だけとし、`output/edited/`へ別名保存する。

`work/`、`output/`、入力フォルダまたは生成先の祖先にsymlinkがある場合や、解決先がリポジトリ外になる場合は処理を停止する。`--results`には当該runの`work/review-runs/<run_id>/results.json`をroot相対で指定する。

## 障害時

- 変換に失敗したファイルを飛ばして続行しない。
- 準備処理に失敗した場合は、その実行で新規作成したrun-id別の`work/`だけをロールバックする。確定処理に失敗した場合は、その実行で新規作成した`finalize-staging`だけを削除し、results修正後に同じrun-idで再試行する。
- OCR結果が不明瞭な場合は文字や図の意味を補完せず、該当項目を`要確認`にする。
- READMEにない外部変換サービスやLibreOfficeを使わない。
- 実行仕様は`references/orchestration-contract.md`を参照する。
