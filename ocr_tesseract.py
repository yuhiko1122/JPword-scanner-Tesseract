import pytesseract
from PIL import Image
import openpyxl
from datetime import datetime
from pathlib import Path

# =============================================
# Windows の場合: Tesseract のインストール先を指定
# Mac で brew install した場合はこの行をコメントアウト
# =============================================
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

INPUT_DIR = Path("inputs")
OUTPUT_FILE = Path("output.xlsx")


def extract_text(image_path: Path) -> str:
    """画像からテキストを抽出する"""
    image = Image.open(image_path)
    # lang="jpn" で日本語 / "jpn+eng" で日英混在に対応
    text = pytesseract.image_to_string(image, lang="jpn")
    return text.strip()


def load_or_create_workbook():
    """既存の Excel を開くか、なければ新規作成する"""
    if OUTPUT_FILE.exists():
        wb = openpyxl.load_workbook(OUTPUT_FILE)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "OCR結果"
        ws.append(["処理日時", "ファイル名", "抽出テキスト"])  # ヘッダー行

        # ヘッダーの列幅を整える
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 60
    return wb, ws


def main():
    # inputs フォルダがなければ作成して終了
    if not INPUT_DIR.exists():
        INPUT_DIR.mkdir()
        print(f"'{INPUT_DIR}' フォルダを作成しました。画像を入れてから再実行してください。")
        return

    image_extensions = {".jpg", ".jpeg", ".png", ".bmp", ".tiff"}
    image_files = [
        f for f in INPUT_DIR.iterdir() if f.suffix.lower() in image_extensions
    ]

    if not image_files:
        print(f"'{INPUT_DIR}' に画像ファイルが見つかりません。")
        return

    wb, ws = load_or_create_workbook()

    success_count = 0
    error_count = 0

    for image_path in sorted(image_files):
        print(f"処理中: {image_path.name} ... ", end="", flush=True)
        try:
            text = extract_text(image_path)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ws.append([timestamp, image_path.name, text])
            print("完了")
            success_count += 1
        except Exception as e:
            print(f"エラー: {e}")
            error_count += 1

    wb.save(OUTPUT_FILE)
    print(f"\n--- 完了 ---")
    print(f"成功: {success_count} 件 / エラー: {error_count} 件")
    print(f"結果を '{OUTPUT_FILE}' に保存しました。")


if __name__ == "__main__":
    main()
