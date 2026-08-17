# Skill Creator へ渡す追加指示: 旧 Office 形式対応

以下の全文を、既存の `SKILL_CREATOR_PROMPT.md` と合わせて Codex の `$skill-creator` に渡してください。

---

あなたは Codex の Skill Creator です。既存の設計書レビュースキルへ、旧 Microsoft Office 形式 `.doc`、`.xls`、`.ppt` を安全に扱う機能を追加してください。

## 1. 基本方針

旧 Office 形式を直接解析することを主要経路にしないでください。旧形式は「前処理が必要な入力形式」として扱い、Windows 11 にインストール済みの Microsoft Office デスクトップアプリケーションを COM Automation で操作し、新形式または PDF へ正規化した後、既存の解析・レビュー処理へ渡してください。

LibreOffice は使用しないでください。LibreOffice の導入、検出、CLI 呼び出し、依存関係、フォールバック経路を実装しないでください。

対応する旧形式と正規化先は次の通りです。

| 入力 | 使用アプリ | 構造解析用の正規化先 | 視覚確認用 |
|---|---|---|---|
| `.doc` | Microsoft Word | `.docx` | PDF |
| `.xls` | Microsoft Excel | `.xlsx` | PDF |
| `.ppt` | Microsoft PowerPoint | `.pptx` | PDF |

PDF は補助的な視覚確認用とし、構造化抽出の正本にはしないでください。構造解析は `.docx`、`.xlsx`、`.pptx` に対する既存処理を再利用してください。

## 2. 前提環境

- OS: Windows 11
- Python: CPython 3.11 x64 を参照環境とする
- 旧形式を扱う場合は、対応する Microsoft Office デスクトップアプリがローカルにインストール済みであることを必須条件とする
- `.doc` には Microsoft Word、`.xls` には Microsoft Excel、`.ppt` には Microsoft PowerPoint が必要
- 必要なアプリが存在しない場合は、その旧形式だけを安全停止する
- `.docx`、`.xlsx`、`.pptx`、`.pdf` の通常処理は Office COM に依存させない
- Office アプリのインストール場所を固定パスでハードコードしない
- 端末固有の絶対パスを設定、manifest、ログ、成果物へ保存しない

## 3. 対応拡張子

レビュー対象と補助資料の対応形式へ次を含めてください。

- `.xlsx`
- `.xls`
- `.pptx`
- `.ppt`
- `.docx`
- `.doc`
- `.pdf`

拡張子の大文字・小文字は区別しないでください。

`.xlsm`、`.docm`、`.pptm`、暗号化・パスワード保護された Office ファイルは、別途明示要件がない限り未対応とし、安全停止してください。マクロ付き形式を通常形式へ暗黙変換しないでください。

## 4. 共通の正規化フロー

旧 Office ファイルは必ず次の順序で処理してください。

1. 入力ファイルの存在、拡張子、サイズ、SHA-256 を記録する
2. 原本がプロジェクトルート配下にあることを確認する
3. 原本を `.work/source/` 配下へ作業コピーする
4. コピー後に原本 SHA-256 が変化していないことを確認する
5. 対応する Office COM Application の利用可否を preflight で確認する
6. このスキル専用の Office Application インスタンスを起動する
7. 作業コピーを、可能な限り読み取り専用・更新抑止状態で開く
8. `.work/converted/` 配下へ新形式を新規保存する
9. 同じ作業コピーから `.work/converted/` 配下へ PDF を新規出力する
10. 文書・ブック・プレゼンテーションを閉じる
11. Office Application を終了する
12. COM オブジェクトを解放する
13. 正規化後ファイルと PDF を検証する
14. 検証成功後のみ、正規化後ファイルを既存の抽出処理へ渡す
15. 元の旧形式ファイルをレビュー対象の正本として manifest と根拠追跡に残す

原本を直接 Save、SaveAs、Export しないでください。必ず `.work/source/` のコピーだけを Office で開いてください。

## 5. `.doc` の正規化

Microsoft Word COM を使用してください。

- `Word.Application` をこのスキル専用インスタンスとして起動する
- UI を表示しない
- 不要な警告・ダイアログを抑止する
- マクロを実行しない
- 外部リンクの自動更新を行わない
- `.work/source/` のコピーを `ReadOnly=True` 相当で開く
- `.docx` は `SaveAs2` 等の Word 正式 API で新規保存する
- PDF は `ExportAsFixedFormat` 等の Word 正式 API で新規出力する
- 既存変換先を無条件上書きしない
- 例外時も文書を閉じ、Word Application を `Quit()` する

変換後 `.docx` は `python-docx` 等の既存抽出経路へ渡してください。

## 6. `.xls` の正規化

Microsoft Excel COM を使用してください。

- `Excel.Application` をこのスキル専用インスタンスとして起動する
- UI を表示しない
- `DisplayAlerts=False` 相当で不要な UI を抑止する
- `EnableEvents=False` 等を使用してイベント起因処理を抑止する
- マクロを実行しない
- 外部リンクを自動更新しない。`Workbooks.Open` では `UpdateLinks=0` 相当を指定する
- `.work/source/` のコピーを読み取り専用で開く
- `.xlsx` は Excel の正式な Open XML Workbook 形式で新規保存する
- PDF は `ExportAsFixedFormat` 等の Excel 正式 API で出力する
- 非表示シート、数式、名前定義、セル値、主要書式、シート順の保持状況を変換後に確認する
- `.xls` に存在する機能が `.xlsx` 変換で失われる可能性を診断として報告する
- 例外時も Workbook を閉じ、Excel Application を `Quit()` する

変換後 `.xlsx` は `openpyxl` 等の既存抽出経路へ渡してください。

Excel PDF は印刷設定・印刷範囲・改ページ・用紙設定の影響を受けるため、PDFに全セルが現れない可能性があります。したがって PDF を構造抽出の正本にせず、視覚確認用としてください。

## 7. `.ppt` の正規化

Microsoft PowerPoint COM を使用してください。

- `PowerPoint.Application` をこのスキル専用インスタンスとして起動する
- 可能な限り UI を表示しない
- マクロを実行しない
- 外部リンク、埋込みオブジェクト等を自動実行しない
- `.work/source/` のコピーを開く
- `.pptx` は PowerPoint の正式な Open XML Presentation 形式で新規保存する
- PDF は `SaveAs` または `ExportAsFixedFormat` 等の PowerPoint 正式 API で出力する
- スライド数、スライド順、タイトル、テキスト、ノート、表、画像等の主要構造を変換後に検証する
- 旧形式由来のアニメーション、OLE、SmartArt相当、埋込み要素等で失われた可能性がある領域を診断として残す
- 例外時も Presentation を閉じ、PowerPoint Application を `Quit()` する

変換後 `.pptx` は `python-pptx` 等の既存抽出経路へ渡してください。

## 8. COM 実装の共通安全条件

Python から Microsoft Office COM を扱うため `pywin32` を使用してください。

次を必須とします。

- 既存のユーザー Office セッションへ安易に接続しない
- 原則としてこのスキル専用の新しい COM Application インスタンスを生成する
- スキルが生成していない Word / Excel / PowerPoint プロセスを終了しない
- COM 初期化・終了を明示し、例外時にも cleanup を行う
- Office Application、Document/Workbook/Presentation 等の参照を残さない
- タイムアウトまたはハング検知を設ける
- UI ダイアログが必要になった場合は無限待機せず失敗として扱う
- 原本、作業コピー、変換先が同一実体でないことを確認する
- 変換先が既に存在する場合は実行 ID 等で衝突を回避する
- Office の Temporary/lock ファイルを成果物として扱わない
- Office プロセスをプロセス名だけで一括 kill しない

## 9. 変換後検証

COM API が成功を返しただけで正常変換とみなさないでください。

共通検証:

- 出力ファイルが存在する
- サイズが 0 より大きい
- 期待する拡張子・コンテナ形式である
- 既存の新形式用ライブラリで再オープンできる
- PDF が再オープンでき、ページ数が 1 以上である
- 原本 SHA-256 が変換前後で一致する

形式別の最低限検証:

- `.doc -> .docx`: 段落数、表数、主要テキスト量、可能なら画像数・セクション数
- `.xls -> .xlsx`: シート数、シート名、順序、使用セル範囲、主要セル値、数式件数
- `.ppt -> .pptx`: スライド数、順序、主要テキスト量、表・画像等の件数

完全一致を保証できない要素は黙って無視せず `diagnostics` に記録してください。

## 10. Manifest

旧形式ごとに、少なくとも次を記録してください。

```json
{
  "source": "input/targets/sample.xls",
  "source_format": "xls",
  "source_sha256": "...",
  "normalizer": "Microsoft Excel COM",
  "normalized_primary": ".work/converted/sample.xlsx",
  "normalized_pdf": ".work/converted/sample.pdf",
  "normalization_status": "success",
  "validation_status": "success"
}
```

端末固有絶対パス、Office のインストールパス、ユーザー名を保存しないでください。

## 11. キャッシュ

キャッシュキーには最低限次を含めてください。

- 原本 SHA-256
- 入力形式
- 正規化処理の版
- 変換先形式
- 検証処理の版

同一入力で検証済みキャッシュが存在する場合は、不要な Office COM 再変換を避けてください。ただしキャッシュファイルが欠落、破損、検証不能なら再変換してください。

Office のアプリバージョン差が変換結果に影響し得るため、キャッシュ metadata へ Office 製品名とバージョンを保存して構いません。ただし端末固有のインストールパスは保存しないでください。バージョン差をキャッシュ再利用条件に含めるかはテスト結果に基づき決定し、その方針を README に記載してください。

## 12. 異常系

次を明示的に扱ってください。

- 対応 Office アプリ未導入
- COM 登録不良
- ファイル破損
- パスワード保護
- 暗号化
- Protected View 等で自動処理できない
- 変換先保存失敗
- PDF 出力失敗
- Office ハング・タイムアウト
- COM 例外
- 正規化後ファイルが開けない
- 主要構造が著しく減少した
- 原本 SHA-256 が変化した

失敗時はレビューを開始せず、原因、対象ファイル、工程、再実行条件を `diagnostics` とログへ記録してください。設計書本文全体や機密セル値をログへ出さないでください。

## 13. 依存関係

`pywin32` を旧 Office 正規化機能に必要な依存として固定・監査してください。

Microsoft Word、Excel、PowerPoint 自体は Python ライブラリではなく、利用者環境へ別途導入されたデスクトップアプリです。スキルから Office を配布・自動インストールしないでください。Microsoft Office の利用条件・ライセンスは利用組織側の契約に従います。

`pywin32` とその依存について、ライセンス、商用利用、再配布条件、既知脆弱性、対応 Python/Windows architecture を確認してください。

LibreOffice、antiword、catdoc、wvWare、unoconv 等を標準経路として追加しないでください。

## 14. 推奨ファイル構成

既存 Skill に次の責務を追加してください。

```text
.agents/skills/review-design-documents/
├─ SKILL.md
├─ scripts/
│  ├─ preflight.py
│  ├─ normalize_legacy_office.py
│  ├─ office_com.py
│  ├─ validate_normalized_office.py
│  └─ extract_documents.py
└─ references/
   └─ legacy-office-normalization.md

.work/
├─ source/
├─ converted/
├─ extracted/
├─ index/
└─ cache/
```

`normalize_legacy_office.py` は拡張子に応じて Word / Excel / PowerPoint の adapter へ分岐する構成にしてください。1つの巨大な関数へ全処理を詰め込まず、共通 cleanup・manifest・cache・validation と、アプリ固有処理をテスト可能な単位へ分離してください。

## 15. テスト

### 15.1 Office 非依存の単体テスト

COM adapter をテストダブル化し、次を自動テストしてください。

- `.doc/.xls/.ppt` の正しい分岐
- 大文字拡張子
- 原本コピー
- SHA-256 不変
- manifest
- cache hit / miss
- 変換先衝突回避
- COM 例外時 cleanup 呼出し
- タイムアウト
- アプリ未導入時の安全停止
- 変換後検証失敗時にレビューへ進まない

### 15.2 Windows integration test

Microsoft Office が導入された Windows 11 環境で、実 COM を使用する opt-in integration test を作成してください。

- `.doc -> .docx + PDF`
- `.xls -> .xlsx + PDF`
- `.ppt -> .pptx + PDF`
- 日本語ファイル名
- 空白を含むパス
- 変換後再オープン
- 原本不変
- 例外後にスキル専用 Office Application が残留しない
- 利用者の既存 Office セッションを終了しない

### 15.3 E2E

チェック表 1 件と、旧形式 `.doc/.xls/.ppt` を含むレビュー対象を使い、次を確認してください。

1. Prepare が旧形式を検出する
2. 必要な Office アプリを preflight する
3. 新形式と PDF を生成する
4. 検証成功後に抽出する
5. 原本旧形式を evidence の source として保持する
6. Codex 判断用チャンクが正規化後データから作られる
7. Finalize でコピー済みチェック表へ結果を書き込む
8. 2回目は検証済みキャッシュを再利用する

## 16. 完了条件

次をすべて満たしてください。

- `.doc/.xls/.ppt` を直接解析せず、Microsoft Office COM で正規化する
- `.doc -> .docx + PDF`
- `.xls -> .xlsx + PDF`
- `.ppt -> .pptx + PDF`
- LibreOffice を一切使用しない
- 原本を一切変更しない
- 必要な Office アプリがない旧形式だけ安全停止する
- 新形式と PDF の双方を検証する
- Office COM の cleanup と timeout を実装・テストする
- ユーザーの既存 Office セッションを終了しない
- `pywin32` を固定・監査する
- manifest と cache を実装する
- Windows 11 実 Office integration test を用意する
- 旧形式を含む E2E を用意する
- 未検証・変換損失の可能性を `要確認` / `diagnostics` として追跡できる

---

以上を既存の設計書レビュースキルへ統合してください。旧形式対応のために既存の `.docx/.xlsx/.pptx/.pdf` 処理を COM 必須へ変更しないでください。