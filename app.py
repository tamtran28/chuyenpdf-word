import streamlit as st
from pdf2docx import Converter
import tempfile
import os

st.set_page_config(page_title="PDF → Word giữ bố cục", layout="wide")
st.title("📄 Chuyển PDF → Word (Giữ bố cục gốc – chạy được trên Streamlit Cloud)")


uploaded = st.file_uploader("📤 Chọn file PDF", type="pdf")

if uploaded:
    st.success("Đã tải PDF thành công!")

    # Lưu PDF vào file tạm
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    temp_pdf.write(uploaded.read())
    temp_pdf.close()

    if st.button("🔄 Chuyển sang Word (giữ bố cục)"):
        with st.spinner("Đang chuyển đổi PDF → Word..."):

            # Tạo file docx tạm
            output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name

            # Chuyển đổi bằng pdf2docx
            cv = Converter(temp_pdf.name)
            cv.convert(output_path, start=0, end=None)  # convert toàn bộ
            cv.close()

        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Tải file Word",
                data=f,
                file_name="converted_layout.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        # Xóa file tạm
        os.unlink(temp_pdf.name)
