# POC08: Draw.io 画像参照方式の変更

## 目的

Excel → Draw.io 変換で、画像の扱いを比較する。

| 区分 | 方式 |
|---|---|
| ASIS | 画像を base64 に変換して Draw.io XML に埋め込む |
| TOBE | 画像を外部ファイルとして `resources/` に出力し、Draw.io から画像パスで参照する |

## 成果物

| ファイル | 内容 |
|---|---|
| `poc08_xlsx_external_images_to_drawio.py` | 外部画像パス参照方式の変換スクリプト |
| `poc08_intermediate.json` | 中間 JSON。画像は `external_path` を持つ |
| `poc08_output.drawio` | Draw.io 出力 |
| `resources/` | Draw.io から参照する画像ファイル |

## 出力例

Draw.io の画像セルは base64 ではなく、以下のような外部パス参照になる。

```xml
<mxCell
  id="image-pic-1"
  value=""
  style="shape=image;html=1;imageAspect=0;aspect=fixed;image=resources/image2.png;"
  vertex="1"
  parent="1">
</mxCell>
```

## 注意点

- `.drawio` 単体では画像を内包しないため、`resources/` フォルダとのセット管理が必要。
- Draw.io / diagrams.net の実行環境によっては、ローカル相対パス画像が表示されない可能性がある。
- 画像を確実に1ファイルで持ち運ぶ用途では base64 方式が有利。
- 差分確認、ファイルサイズ削減、画像差し替えのしやすさでは外部参照方式が有利。
