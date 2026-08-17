# Skill Creator へ渡す追加指示: 旧 Word `.doc` 対応

以下の全文を、既存の `SKILL_CREATOR_PROMPT.md` と合わせて Codex の `$skill-creator` に渡してください。

---

あなたは Codex の Skill Creator です。既存の設計書レビュースキルへ、旧 Microsoft Word 形式 `.doc` を安全に扱う機能を追加してください。

## 1. 目的

`.doc` を直接解析することを前提にしないでください。`.doc` は「前処理が必要な入力形式」として扱い、Windows 11 にインストール済みの Microsoft Word を COM Automation で操作して、作業用の `.docx` と PDF へ正規化してから既存の解析・レビュー処理へ渡してください。

LibreOffice は使用しないでください。LibreOffice の導入、検出、CLI 呼び出し、依存関係、フォールバック経路を実装しないでください。

## 2. 前提環境

- OS: Windows 11
- Python: CPython 3.11 x64 を参照環境とする
- `.doc` 対応時は Microsoft Word デスクトップ版がローカルにインストール済みであることを必須条件とする
- Word が利用できない環境では `.doc` のレビューを開始せず、安全停止する
- `.docx`、`.xlsx`、`.pptx`、`.pdf` の既存処理は Word COM に依存させない
- 端末固有の絶対パスを設定、manifest、ログ、成果物へ保存しない

## 3. 対応拡張子

既存の対応形式へ `.doc` を追加してください。

- `.xlsx`
- `.pptx`
- `.docx`
- `.doc`
- `.pdf`

拡張子の大文字・小文字は区別しないでください。補助資料にも `.doc` を許可してください。

`.xls`、`.ppt`、`.xlsm`、暗号化・パスワード保護された Office ファイルは、別途対応要件がない限り未対応のままとし、理由を示して安全停止してください。

## 4. `.doc` 正規化の基本方針

`.doc` はレビュー対象の原本として保持し、原本そのものを Word で上書き保存しないでください。

処理は必ず次の順序にしてください。

1. 入力ファイルの存在、拡張子、サイズ、SHA-256 を記録する
2. 原本がプロジェクトルート配下にあることを確認する
3. 原本を `.work/source/` 配下へ作業コピーする
4. コピー後に原本 SHA-256 が変化していないことを確認する
5. Microsoft Word の利用可否を preflight で確認する
6. Word COM で作業コピーを読み取り専用として開く
7. `.work/converted/` 配下へ `.docx` を新規保存する
8. 同じ作業コピーから `.work/converted/` 配下へ PDF を新規保存する
9. Word 文書を閉じる
10. Word Application を終了する
11. COM オブジェクトを解放する
12. 変換後の `.docx` と PDF を検証する
13. 検証成功後のみ、`.docx` を既存の Word 抽出処理へ渡す
14. PDF は視覚・ページ確認用の補助資料として扱う
15. 元の `.doc` をレビュー対象ファイルの正本として追跡情報へ残す

`.doc` をバイナリパーサーで直接解析する経路を主要実装にしないでください。antiword 等の外部コマンドを標準経路として導入しないでください。

## 5. Microsoft Word COM の利用

Python から Microsoft Word COM を扱うため、`pywin32` を `.doc` 対応機能の依存として使用してください。

次の安全条件を満たしてください。

- `Word.Application` はバックグラウンドで起動する
- `Visible = False`
- `DisplayAlerts = 0` 相当で不要な UI ダイアログを抑止する
- マクロを実行しない
- 自動リンク更新を行わない
- 原本ではなく `.work/source/` のコピーを開く
- `ReadOnly=True` を指定する
- 変換先は必ず原本と別ファイルにする
- `SaveAs2` / `ExportAsFixedFormat` 等、Microsoft Word が提供する正式な保存・PDF出力機能を利用する
- 既存ファイルを無条件で上書きしない
- 例外発生時でも `Close()` と `Quit()` を `finally` 相当で実行する
- COM オブジェクト参照を残さず解放する
- 処理終了後に、このスキルが起動した Word プロセスが残留していないことを可能な範囲で検証する

既にユーザーが開いている Word セッションや文書を終了しないでください。このスキル自身が生成した Word Application インスタンスだけを管理対象にしてください。

## 6. 変換先とファイル名

作業フォルダ内に次を用意してください。

```text
.work/
├─ source/
├─ converted/
├─ extracted/
├─ index/
└─ cache/
```

例えば入力が次の場合、

```text
input/targets/basic_design.doc
```

内部成果物は概ね次のようにしてください。

```text
.work/source/basic_design__<短いID>.doc
.work/converted/basic_design__<短いID>.docx
.work/converted/basic_design__<短いID>.pdf
```

同名ファイルの衝突を避けるため、入力相対パスまたは SHA-256 から安定した短い ID を生成してください。

作業ファイル名や manifest には端末固有の絶対パスを保存しないでください。

## 7. `.docx` 側の役割

変換後 `.docx` は構造解析用です。既存の `python-docx` ベースの抽出処理へ渡し、少なくとも次を取得してください。

- 見出し
- 段落
- 表
- ヘッダー
- フッター
- セクション情報
- 画像や図形等の存在診断
- 変更履歴、埋込みオブジェクト等の未対応領域の検出

変換元が `.doc` である場合、抽出された各チャンクには次を追跡できる情報を持たせてください。

- 原本相対パス
- 原本形式 `doc`
- 正規化済み `.docx` 相対パス
- 変換方式 `microsoft_word_com`
- 変換検証状態

Codex へ提示する根拠では、可能な限り原本 `.doc` のファイル名を主表示とし、内部で `.docx` に正規化したことを診断情報として保持してください。

## 8. PDF 側の役割

変換後 PDF は見た目・ページ構成確認用です。PDF を `.docx` の代わりの主抽出元にしないでください。

PDF では少なくとも次を確認できるようにしてください。

- PDF が開けること
- ページ数が 1 以上であること
- ページごとの文字量
- 画像・図・表らしき領域の有無
- ページ画像によるオンデマンド確認が可能であること

`.docx` 抽出結果と PDF の視覚情報に大きな差がある場合は、変換・抽出欠落の可能性を診断として残してください。

例:

- `.docx` 側の抽出画像件数が 0 だが PDF に図がある
- `.docx` 側の本文が極端に少ないが PDF は複数ページ存在する
- 表・図形・テキストボックス等の未抽出要素が疑われる

これらを自動的に `不適合` としないでください。レビュー根拠不足として `要確認` へつながる診断情報にしてください。

## 9. 変換後検証

Word COM の API 呼び出しが成功しただけで変換成功と判断しないでください。

`.docx` について最低限、次を確認してください。

- ファイルが存在する
- サイズが 0 より大きい
- ZIP/OOXML として基本構造が妥当である
- `python-docx` で開ける
- 段落、表、セクション等の抽出処理が異常終了しない

PDF について最低限、次を確認してください。

- ファイルが存在する
- サイズが 0 より大きい
- PDF として開ける
- 暗号化されていない、または意図しない暗号化状態でない
- ページ数が 1 以上である

変換結果が検証に失敗した場合はレビューへ進まず、診断コードと理由を残して停止してください。

## 10. manifest

`.doc` ごとに、少なくとも次の情報を manifest または同等の構造化データへ記録してください。

```json
{
  "source": "input/targets/basic_design.doc",
  "source_format": "doc",
  "source_sha256": "...",
  "normalization": {
    "converter": "microsoft_word_com",
    "docx": ".work/converted/basic_design__xxxx.docx",
    "pdf": ".work/converted/basic_design__xxxx.pdf",
    "status": "success"
  },
  "validation": {
    "docx": "success",
    "pdf": "success"
  }
}
```

端末固有の絶対パス、COM オブジェクト情報、ユーザー名を保存しないでください。

## 11. キャッシュ

同一 `.doc` を毎回 Word で再変換しないでください。

キャッシュキーには少なくとも次を含めてください。

- 原本 SHA-256
- 正規化処理のバージョン
- 主要な変換設定
- 必要であれば Microsoft Word の major version

キャッシュが有効で、`.docx` と PDF の検証済み成果物が揃っている場合は再利用してください。

ただし次の場合は再変換してください。

- 原本 SHA-256 が変わった
- 変換ロジックのバージョンが変わった
- 成果物のどちらかが欠落・破損している
- 変換設定が変わった

## 12. Word がない場合の挙動

preflight で Microsoft Word COM が利用できない場合、`.doc` を黙ってスキップしないでください。

次のような明示的な診断を返してください。

```text
DOC_NORMALIZATION_UNAVAILABLE

対象: input/targets/basic_design.doc
理由: Microsoft Word desktop / Word COM Automation を利用できません。
結果: .doc のレビューは開始していません。
対処: Microsoft Word が利用可能な Windows 11 環境で再実行するか、利用者側で .docx または PDF へ変換したファイルを入力してください。
```

LibreOffice を代替手段として案内・自動利用しないでください。

## 13. セキュリティと安全性

- 原本 `.doc` は変更禁止
- 原本 SHA-256 を変換前後で比較する
- マクロを実行しない
- 外部リンクを自動更新しない
- Word が出す確認ダイアログで無人処理が停止しないよう安全な設定を行う
- パスワード要求、破損修復確認、危険な埋込み要素等が発生した場合は自動承認せず停止する
- `.doc` から抽出した OLE、マクロ、埋込み実行ファイル等を実行しない
- 一時ファイルは `.work/` 配下に限定する
- 異常終了後も原本と既存成果物が不変であることを確認する
- ログへ設計書全文を出さない

## 14. 依存関係

`.doc` 対応を有効にする Windows 環境では `pywin32` を必要とします。

既存の依存管理方針に従い、直接依存・推移依存を固定し、ライセンス、無料利用、商用利用、再配布条件、既知脆弱性、Windows wheel の有無を監査してください。

`pywin32` を単に「任意の描画機能」扱いにせず、`.doc` を対応形式として掲げる構成では `.doc` 正規化機能に必要な依存として明確化してください。

ただし Microsoft Word 自体は Python パッケージではなく、ユーザー環境に別途必要なデスクトップアプリケーションです。README に次を明記してください。

- `.doc` 処理には Microsoft Word が必要
- Microsoft Word の利用条件・ライセンスは利用組織側の契約に従う
- スキルは Microsoft Word を配布しない
- LibreOffice は使用しない

## 15. 実装ファイル

既存構成へ、必要に応じて次の責務を追加してください。

```text
.agents/skills/review-design-documents/scripts/
├─ preflight.py
├─ normalize_legacy_word.py
├─ validate_normalized_word.py
└─ extract_documents.py
```

同じ責務を安全に統合できる場合はファイル名・数を変更して構いません。ただし次は独立してテスト可能にしてください。

- Word COM 利用可否判定
- `.doc` 作業コピー作成
- `.doc -> .docx` 変換
- `.doc -> PDF` 変換
- 変換結果検証
- Word COM cleanup
- manifest 生成
- キャッシュ判定

## 16. テスト要件

### 16.1 単体テスト

最低限、次を作成してください。

- `.doc` を対応拡張子として認識する
- `.DOC` も認識する
- `.doc` 原本の SHA-256 が変換前後で不変
- 同名 `.doc` の作業ファイルが衝突しない
- Word がない場合に preflight が安全停止する
- 変換後 `.docx` が破損している場合にレビューへ進まない
- 変換後 PDF が 0 ページまたは破損している場合にレビューへ進まない
- manifest に端末固有絶対パスが残らない
- キャッシュヒット時に Word COM を再起動しない
- 例外発生時にも Word cleanup が実行される

Word COM 自体を単体テストで毎回起動する必要はありません。COM 呼出しを抽象化し、fake/mock で成功・失敗・例外・cleanup を検証してください。

### 16.2 Windows integration test

Microsoft Word がインストールされた Windows 11 環境で、小さな安全な `.doc` fixture を使って実変換テストを追加してください。

確認項目:

1. `.doc` から `.docx` を生成できる
2. `.doc` から PDF を生成できる
3. `.docx` を `python-docx` で開ける
4. PDF を PDF ライブラリで開ける
5. ページ数が 1 以上
6. 原本 SHA-256 が不変
7. 作業コピー以外を書き換えていない
8. Word プロセスが残留しない
9. 変換後の本文・表について fixture の既知要素を抽出できる
10. 2 回目はキャッシュ利用で再変換を避ける

Word 未導入環境ではこの integration test を成功扱いでスキップせず、「環境要件未充足で未検証」であることをテストレポートへ明示してください。

### 16.3 E2E

既存の `Prepare -> Codex review -> Finalize` フローへ `.doc` を含めてください。

- チェック表 1 件
- `.doc` 対象 1 件以上
- 必要に応じて `.docx/.xlsx/.pptx/.pdf` も混在
- `.doc` は Prepare 内で正規化される
- Codex に渡す抽出チャンクは正規化済み `.docx` 由来だが、対象 ID は原本 `.doc` と対応する
- Finalize の Excel 結果列ではレビュー対象ファイル名として原本 `.doc` を記録する
- 正規化失敗時は Finalize でレビュー済み扱いにしない

## 17. README / SKILL.md への反映

README と SKILL.md に次を明記してください。

- `.doc` は対応対象
- `.doc` は直接解析せず Microsoft Word で `.docx` と PDF へ正規化してから扱う
- `.doc` 処理には Microsoft Word デスクトップ版が必要
- LibreOffice は使用しない
- `.docx` は構造抽出、PDF は視覚確認の役割
- 原本を変更しない
- Word が利用できない場合は `.doc` のみ安全停止する
- `.docx/.xlsx/.pptx/.pdf` の通常処理は Word がなくても利用できる

## 18. 完了条件

次をすべて満たすまで `.doc` 対応完了としないでください。

1. `.doc` が入力対象として認識される
2. 原本を上書きしない
3. `.doc -> .docx` が Word COM で実装されている
4. `.doc -> PDF` が Word COM で実装されている
5. LibreOffice を一切使用していない
6. 変換後 `.docx` と PDF を独立検証している
7. 変換失敗時はレビューを開始しない
8. Word 未導入時は明示的に安全停止する
9. Word COM cleanup が例外時も保証される
10. 原本 SHA-256 不変を自動テストしている
11. manifest に原本、正規化成果物、変換方式、検証状態を記録している
12. キャッシュが実装されている
13. `.doc` を含む Windows integration test が用意されている
14. `.doc` を含む E2E フローが用意されている
15. README、SKILL.md、依存関係資料が実装と一致している

既存の設計書レビュー機能を壊さず、`.doc` を安全な正規化経路として追加してください。

---
