# POC08: Draw.io 画像方式の確認

## 目的

Excel → Draw.io 変換で外部画像参照を試したが、Draw.io はローカル相対パスを画像ソースとして扱えない。
そのため、Excel は base64 埋め込み方式を採用する方針に戻す。

| 区分 | 方式 |
|---|---|
| Excel 方針 | 画像を base64 に変換して Draw.io XML に埋め込む |
| 外部参照検証 | `resources/` や相対パスでは Draw.io 表示不可。URLなら表示可能 |
| Word 追加検証 | Word `.docx` を Draw.io に変換し、画像は base64 で埋め込む |

## 成果物

| ファイル | 内容 |
|---|---|
| `poc08_xlsx_external_images_to_drawio.py` | Excel 外部画像URL参照方式の検証スクリプト |
| `poc08_docx_to_drawio.py` | Word → Draw.io 変換の検証スクリプト |
| `poc08_word_intermediate.json` | Word 解析結果の中間 JSON |
| `poc08_word_output.dio.xml` | Word → Draw.io 出力。画像は base64 埋め込み |
| `poc08_intermediate.json` | Excel 外部参照検証用の中間 JSON |

## 実行方法

Draw.io は `resources/image2.png` のようなローカル相対パスを画像ソースとして受け付けない。
そのため、画像ファイルは外部管理のまま `resources/` に置き、Draw.io XML には URL を出力する。

```bash
python3 doc/20.InternalDesign/00.POC/excel→drowio/08/poc08_xlsx_external_images_to_drawio.py \
  --input "doc/30.test/00.pocTestData/テストデータ2.xlsx" \
  --drawio doc/20.InternalDesign/00.POC/excel→drowio/08/poc08_output.dio.xml \
  --image-base-url http://localhost:8008/resources
```

ローカル確認時は、08フォルダで簡易HTTPサーバーを起動する。

```bash
cd doc/20.InternalDesign/00.POC/excel→drowio/08
python3 -m http.server 8008
```

## 出力例

Draw.io の画像セルは base64 ではなく、以下のような外部URL参照になる。

```xml
<mxCell
  id="image-pic-1"
  value=""
  style="shape=image;html=1;imageAspect=0;aspect=fixed;image=http://localhost:8008/resources/image2.png;"
  vertex="1"
  parent="1">
</mxCell>
```

## 注意点

- `.drawio` 単体では画像を内包しないため、`resources/` フォルダとのセット管理が必要。
- Draw.io / diagrams.net は相対パス画像を不正な画像ソースとして扱うため、外部参照方式では URL 化が必要。
- 画像を確実に1ファイルで持ち運ぶ用途では base64 方式が有利。
- 差分確認、ファイルサイズ削減、画像差し替えのしやすさでは外部参照方式が有利。

## Word → Draw.io 検証

対象:

```text
doc/30.test/00.pocTestData/ペット プロフィール - スペアミント (1).docx
```

実行:

```bash
python3 doc/20.InternalDesign/00.POC/excel→drowio/08/poc08_docx_to_drawio.py \
  --input "doc/30.test/00.pocTestData/ペット プロフィール - スペアミント (1).docx" \
  --json doc/20.InternalDesign/00.POC/excel→drowio/08/poc08_word_intermediate.json \
  --drawio doc/20.InternalDesign/00.POC/excel→drowio/08/poc08_word_output.dio.xml
```

確認結果:

- Word 内の画像は base64 `data:image/...` として Draw.io に埋め込む。
- 先頭のペット写真とプロフィール文は、Draw.io 上で左画像・右テキストに簡易配置する。
- Word 図形テキスト `オブジェクトA` / `オブジェクトB` は Draw.io の角丸矩形と矢印に変換する。
- Word の完全なページレイアウト再現ではなく、構造確認用の POC とする。
