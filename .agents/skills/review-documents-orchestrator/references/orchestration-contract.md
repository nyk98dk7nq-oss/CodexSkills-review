# 統合処理契約

## 入力

- `input/checklists/`: `.xlsx`、またはWindows版Microsoft Officeで変換する`.xls`だけを受け付ける。
- `input/references/`、`input/targets/`: `.xlsx`、`.docx`、`.pptx`、`.pdf`、PNG、JPEG、TIFF、BMP、WebPと、旧Office形式を受け付ける。
- `README.md`とドットで始まるファイルは入力として数えない。
- 未対応ファイルがあれば、そのファイルを飛ばさず実行全体を停止する。

## 形式別Skill CLI

```text
python <format_script> to-markdown INPUT OUTPUT --role ROLE --repo-root ROOT --images-dir DIR
python convert_legacy_office.py convert INPUT OUTPUT
python xlsx_document.py write-review INPUT OUTPUT --results RESULTS --checklist-path RELATIVE_PATH
```

画像とPDFには`--ocr --lang jpn+eng+jpn_vert`を追加する。XLSX、DOCX、PPTXの抽出画像はimage-documentで別途OCRし、元文書Markdownの`抽出画像OCR`節へ追記する。

抽出画像がimage-documentの未対応形式なら、その画像だけを黙って除外しない。元文書、抽出画像パス、拡張子と「OCR不可のため要確認。視覚情報を推測しない」を同節へ記録してレビューを続ける。対応画像のOCR処理自体が失敗した場合は実行全体を停止する。

## manifest

manifestの`root`は`.`とし、リポジトリの絶対パスを永続化しない。各入力要素は、リポジトリ相対の`path`と`original_path`、Markdownの相対パス、変換後形式を保持する。中間物は`work/markdown/<run_id>/`、`work/images/<run_id>/`へ分離する。旧Office形式だけ`intermediate_path`へ`work/converted-office/<run_id>/`の変換先を記録する。レビューSkillはmanifestの固定配置からリポジトリルートを導出し、ルート外を指すMarkdownパスを拒否する。

チェックリストはopenpyxlで直接読み、各シートの2行目以降の非空行を`checklist_items`にする。`check_item`は`列名/セル座標=値`を結合し、`row`には元のExcel行番号を保持する。

## 安全性

- 入力ファイルを保存先として渡さない。
- 中間ファイルは`work/`のrun-id別フォルダへ書き、確定結果は全ファイルの生成成功後に`output/reviews/<run_id>/`へ移す。
- 同じrun-idの`work/review-runs/`、`work/markdown/`、`work/images/`、`work/converted-office/`、`output/reviews/`のどれかがあれば、空でも開始前に停止する。
- Windowsで大文字小文字だけが異なるチェックリスト名も、結果ファイル名が衝突しないよう連番を付ける。
- 準備中の例外では、所有マーカーが今回の実行と一致する4つのrun-id別`work/`だけを削除する。既存の入力・出力は削除しない。
- 確定中の例外では、今回作成した`finalize-staging`だけを削除する。manifestとresultsを残し、同じrun-idで確定処理を再試行できるようにする。
- `input/`、`work/`、`output/`の構造ディレクトリと全生成先について祖先symlinkを拒否し、resolve後もリポジトリルート配下であることを確認する。
- manifestの`path`、`original_path`、`intermediate_path`、`markdown`、`checklist_file`はroot相対、役割別の所定prefix配下、かつリポジトリ内に限定する。絶対パスと`..`を拒否する。
- finalizeのresultsは`work/review-runs/<run_id>/results.json`だけをroot相対で受け付ける。
- 途中の変換、OCR、検証が1件でも失敗したら確定結果を作らない。
