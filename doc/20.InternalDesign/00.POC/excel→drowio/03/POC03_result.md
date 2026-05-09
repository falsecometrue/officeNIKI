# POC03 結果まとめ

## 1. 実施内容
- `poc03_html_to_ai_readable.py` を 03 フォルダに作成。
- `doc/30.test/testData/Excel_sampleData/シート1.html` を入力として実行。
- 出力ファイル:
  - `poc03_output.md`
  - `poc03_summary.json`
  - `poc03_drawio_hints.xml`

## 2. POC で確認した事項
- Excel 由来の HTML 出力 (`<table class="waffle">`) からテーブル構造を抽出し、Markdown に変換できた。
- 画像リソース `resources/drawing0.png` を検出し、埋め込みオブジェクトとして保持した。
- JavaScript の `posObj(...)` 呼び出しを解析し、表示位置ヒント (`x=98, y=16`) を取得した。
- 表内に記述されたオブジェクト情報から、Draw.io 向け図形候補（`オブジェクトA`, `オブジェクトB`）を生成した。

## 3. 生成ファイル内容
- `poc03_output.md`
  - HTML 解析結果の Markdown 出力
  - テーブル全体、画像リソース、埋め込みオブジェクト、JS 位置情報、Draw.io ヒント抽出結果を含む
- `poc03_summary.json`
  - 抽出したテーブルデータ、画像、オーバーレイ、posObj 情報を JSON 形式で保持
- `poc03_drawio_hints.xml`
  - Draw.io XML のスケルトンを出力
  - 表データから抽出した図形候補を簡易的に配置する形式

## 4. 検証結果と考察
-  **テーブルは直接変換可能** であることを確認。
-  **画像はそのまま流用** し、AI リーダブル化の前段階で保持できる。
-  **OCR/意味解釈は今回は保留** とし、GUI 補正フローで対応する方針に合致している。
-  **オブジェクトヒントは表情報から十分に生成できるが、最終補正は人間による方針が妥当** である。

## 5. 追加対応候補
- `poc03_drawio_hints.xml` を元に、実際の draw.io 形式の形状・矢印の自動配置を強化する。
- `resources/drawing0.png` の画像単体を Markdown に埋め込む処理を追加する。
- `posObj` のセル基準オフセットを実際のセル位置に変換し、より正確な配置情報を出力する。

