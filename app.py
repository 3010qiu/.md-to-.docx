import os
import re
import shutil
import tempfile
import subprocess
import streamlit as st
from pathlib import Path

# Hàm tìm kiếm Pandoc trên máy tính (lấy từ code gốc của bạn)
def find_pandoc():
    paths = [
        shutil.which("pandoc"),
        r"C:\Program Files\Pandoc\pandoc.exe",
        os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\Pandoc\pandoc.exe"),
        "pandoc" # Mặc định cho Linux/Mac hoặc server Web
    ]

    for p in paths:
        if not p:
            continue
        path = Path(p)
        # Nếu là file .exe tồn tại hoặc là lệnh hệ thống ("pandoc")
        if path.exists() or p == "pandoc":
            return p
    return None

st.set_page_config(page_title="Markdown to Word", page_icon="📝")
st.title("Chuyển đổi Markdown sang Word")

# Kiểm tra Pandoc trước khi chạy
pandoc_path = find_pandoc()

if not pandoc_path:
    st.error("❌ Không tìm thấy Pandoc trên máy. Vui lòng cài đặt tại: https://pandoc.org/installing.html")
    st.stop() # Dừng chạy code bên dưới nếu không có Pandoc

st.write("Tải lên file Markdown (.md, .txt) để nhận lại file định dạng .docx")

uploaded_file = st.file_uploader("Chọn file Markdown", type=["md", "txt"])

if uploaded_file is not None:
    original_name = os.path.splitext(uploaded_file.name)[0]
    
    with st.spinner('Đang xử lý...'):
        try:
            text = uploaded_file.read().decode("utf-8")
            text = re.sub(r"(?<!\n)\n(?!\n)", "  \n", text)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".md") as temp_md:
                temp_md.write(text.encode("utf-8"))
                temp_md_path = temp_md.name

            output_docx_path = temp_md_path.replace(".md", ".docx")

            # Sử dụng pandoc_path đã tìm được thay vì gọi chữ "pandoc" cứng ngắc
            cmd = [
                pandoc_path,
                temp_md_path,
                "-f", "markdown+tex_math_dollars+tex_math_single_backslash",
                "-o", output_docx_path
            ]
            subprocess.run(cmd, check=True)

            with open(output_docx_path, "rb") as f:
                docx_data = f.read()

            st.success("Chuyển đổi thành công!")
            
            st.download_button(
                label="📥 Tải file Word về máy",
                data=docx_data,
                file_name=f"{original_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except subprocess.CalledProcessError as e:
            st.error(f"Lỗi Pandoc: Không chuyển được file.\n{e}")
        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {e}")
        finally:
            if 'temp_md_path' in locals() and os.path.exists(temp_md_path):
                os.remove(temp_md_path)
            if 'output_docx_path' in locals() and os.path.exists(output_docx_path):
                os.remove(output_docx_path)