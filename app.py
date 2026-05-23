import os
import re
import tempfile
import streamlit as st
import pypandoc

st.set_page_config(page_title="Markdown to Word", page_icon="📝")
st.title("Chuyển đổi Markdown sang Word")

# Hàm tự động tải Pandoc nếu máy chủ (hoặc máy cá nhân) chưa có
@st.cache_resource
def setup_pandoc():
    try:
        pypandoc.get_pandoc_version()
    except OSError:
        with st.spinner("Đang khởi tạo Pandoc cho máy chủ (chỉ mất vài giây ở lần chạy đầu tiên)..."):
            pypandoc.download_pandoc()

# Chạy kiểm tra/tải Pandoc ngay khi mở web
setup_pandoc()

st.write("Tải lên file Markdown (.md, .txt) để nhận lại file định dạng .docx")

uploaded_file = st.file_uploader("Chọn file Markdown", type=["md", "txt"])

if uploaded_file is not None:
    original_name = os.path.splitext(uploaded_file.name)[0]
    
    with st.spinner('Đang xử lý...'):
        try:
            # Đọc và xử lý nội dung
            text = uploaded_file.read().decode("utf-8")
            text = re.sub(r"(?<!\n)\n(?!\n)", "  \n", text)

            # Tạo file tạm
            with tempfile.NamedTemporaryFile(delete=False, suffix=".md") as temp_md:
                temp_md.write(text.encode("utf-8"))
                temp_md_path = temp_md.name

            output_docx_path = temp_md_path.replace(".md", ".docx")

            # Gọi pypandoc để chuyển đổi thay vì dùng subprocess
            pypandoc.convert_file(
                temp_md_path,
                'docx',
                outputfile=output_docx_path,
                extra_args=["-f", "markdown+tex_math_dollars+tex_math_single_backslash"]
            )

            # Đọc file Word xuất ra
            with open(output_docx_path, "rb") as f:
                docx_data = f.read()

            st.success("Chuyển đổi thành công!")
            
            st.download_button(
                label="📥 Tải file Word về máy",
                data=docx_data,
                file_name=f"{original_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {e}")
        finally:
            # Dọn dẹp file tạm
            if 'temp_md_path' in locals() and os.path.exists(temp_md_path):
                os.remove(temp_md_path)
            if 'output_docx_path' in locals() and os.path.exists(output_docx_path):
                os.remove(output_docx_path)