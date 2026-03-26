# セットアップ手順

## 1. Tesseract 本体のインストール

### Windows
1. 下記からインストーラをダウンロード
   https://github.com/UB-Mannheim/tesseract/wiki
2. インストール時に **「Japanese」** にチェックを入れる（重要）
3. インストール先はデフォルトのまま OK
   `C:\Program Files\Tesseract-OCR\`

### Mac
```bash
brew install tesseract
brew install tesseract-lang
```

---

## 2. Python ライブラリのインストール

```bash
pip install pytesseract pillow openpyxl
```

---

## 3. 使い方

1. `inputs/` フォルダに画像（jpg / png など）を入れる
2. 下記コマンドで実行
   ```bash
   python ocr_tesseract.py
   ```
3. 同じフォルダに `output.xlsx` が生成される

---

## よくあるエラー

| エラーメッセージ | 原因 | 対処 |
|---|---|---|
| `tesseract is not installed or not in path` | Tesseract が見つからない | `tesseract_cmd` のパスを確認 |
| `Failed loading language 'jpn'` | 日本語データ未インストール | Tesseract を再インストールし「Japanese」を選択 |
| `ModuleNotFoundError` | ライブラリ未インストール | `pip install pytesseract pillow openpyxl` を実行 |
