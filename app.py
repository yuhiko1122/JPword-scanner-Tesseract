import pytesseract
from PIL import Image
import openpyxl
from datetime import datetime
from io import BytesIO
import streamlit as st

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

st.set_page_config(page_title="手書きOCR", page_icon="📝", layout="centered")
st.title("📝 手書き文字 → テキスト変換")
st.caption("画像をアップロードすると文字を読み取ります")

uploaded_files = st.file_uploader(
    "画像を選択（複数可）",
    type=["jpg", "jpeg", "png", "bmp", "tiff"],
    accept_multiple_files=True,
)

if uploaded_files:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "OCR結果"
    ws.append(["処理日時", "ファイル名", "抽出テキスト"])
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 60

    for uploaded_file in uploaded_files:
        st.divider()
        st.subheader(uploaded_file.name)

        image = Image.open(uploaded_file)
        st.image(image, use_container_width=True)

        with st.spinner("読み取り中..."):
            try:
                text = pytesseract.image_to_string(image, lang="jpn").strip()
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                ws.append([timestamp, uploaded_file.name, text])

                st.text_area("抽出テキスト", value=text if text else "（テキストを検出できませんでした）", height=150)
            except Exception as e:
                st.error(f"エラー: {e}")

    st.divider()

    # Excel をメモリ上に保存してダウンロードボタンを表示
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="📥 Excelをダウンロード",
        data=buffer,
        file_name=f"ocr_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
