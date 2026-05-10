# office to mdrow

VS CodeのエクスプローラーでOfficeファイルを右クリックし、変換ファイルを同一フォルダに出力する拡張機能プロジェクト。

## 対応コマンド

| 入力 | 右クリックメニュー | 出力 |
|---|---|---|
| `.xlsx` | Excel → Draw.io に変換 | 同一フォルダの `.drawio` |
| `.docx` | Word → Markdown に変換 | 同一フォルダの `.md` と `resources/` |

## 開発実行

```bash
npm install
npm run compile
```

VS Codeでこのフォルダを開き、拡張機能ホストを起動する。

Pythonは `officeToMdrow.pythonPath` 設定で変更できる。デフォルトは `python3`。

## 構成

| パス | 内容 |
|---|---|
| `src/extension.ts` | VS Code拡張機能のエントリポイント |
| `scripts/convert.py` | 拡張機能から呼び出す変換ラッパー |
| `scripts/xlsx_to_drawio.py` | Excel→Draw.io変換処理 |
| `scripts/docx_to_md.py` | Word→Markdown変換処理 |

## 出力方針

- 中間JSONは最終成果物として残さない。
- Word画像は `resources/` 配下に外部ファイルとして出力する。
- Excel画像はDraw.io XML内へbase64で埋め込む。
