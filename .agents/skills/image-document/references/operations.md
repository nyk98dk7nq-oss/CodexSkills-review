# 画像操作定義

`edit` は UTF-8 JSON の配列、または `operations` 配列を持つオブジェクトを受け取る。操作を上から順に適用する。`frames` は `"all"` または 1 始まりの番号配列とし、省略時は全フレームを対象にする。

主出力またはその既存祖先フォルダがシンボリックリンクの場合は処理を拒否する。リンク切れの出力シンボリックリンクも既存出力として扱う。編集済み画像は同一フォルダの一時ファイルへ完成させてから、未作成の主出力へ確定する。

## 切り抜き

```json
{"op": "crop", "frames": "all", "box": [10, 20, 810, 620]}
```

`box` は `[left, top, right, bottom]` のピクセル座標とし、各対象フレームの範囲内に収める。

## 回転

```json
{"op": "rotate", "frames": [1], "degrees": 90, "expand": true}
```

正の角度は反時計回りとする。`expand` の既定値は `true` とする。

## サイズ変更

```json
{"op": "resize", "width": 1200, "resample": "lanczos"}
```

- `width` または `height` の一方だけなら縦横比を維持する。
- 両方を指定した場合は指定寸法へ変形する。
- `resample` は `nearest`、`bilinear`、`bicubic`、`lanczos` とする。

## グレースケール

```json
{"op": "grayscale", "frames": "all"}
```

## モード・形式変換

```json
{"op": "convert", "mode": "RGB", "format": "JPEG"}
```

- `mode` は Pillow が扱える画像モードとする。
- `format` を指定する場合は出力拡張子と一致させる。
- 出力形式は出力ファイルの拡張子から決定する。
- 複数フレームの出力形式は TIFF または WebP に限定する。単一フレーム形式への暗黙のフレーム削除は行わない。

## 複数操作の例

```json
{
  "operations": [
    {"op": "crop", "box": [0, 0, 1000, 700]},
    {"op": "resize", "width": 800},
    {"op": "grayscale"}
  ]
}
```
