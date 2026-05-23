import os
import re
import tempfile
import subprocess
import streamlit as st

st.set_page_config(page_title="Markdown to Word", page_icon="📝")
st.title("Chuyển đổi Markdown sang Word")
st.write("Tải lên file Markdown (.md, .txt) để nhận lại file định dạng .docx")

# Tạo giao diện upload file
uploaded_file = st.file_uploader("Chọn file Markdown", type=["md", "txt"])

if uploaded_file is not None:
    # Lấy tên file gốc để làm tên file tải xuống
    original_name = os.path.splitext(uploaded_file.name)[0]
    
    with st.spinner('Đang xử lý...'):
        try:
            # 1. Đọc nội dung file
            text = uploaded_file.read().decode("utf-8")

            # Giữ xuống dòng đơn trong Markdown khi chuyển sang Word
            text = re.sub(r"(?<!\n)\n(?!\n)", "  \n", text)

            # 2. Tạo file tạm an toàn cho môi trường web (tránh trùng lặp giữa các user)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".md") as temp_md:
                temp_md.write(text.encode("utf-8"))
                temp_md_path = temp_md.name

            output_docx_path = temp_md_path.replace(".md", ".docx")

            # 3. Gọi Pandoc (trên server web, Pandoc sẽ nằm sẵn trong hệ thống)
            cmd = [
                "pandoc",
                temp_md_path,
                "-f", "markdown+tex_math_dollars+tex_math_single_backslash",
                "-o", output_docx_path
            ]
            subprocess.run(cmd, check=True)

            # 4. Đọc file Word đầu ra để chuẩn bị cho nút tải xuống
            with open(output_docx_path, "rb") as f:
                docx_data = f.read()

            st.success("Chuyển đổi thành công!")
            
            # Nút tải file
            st.download_button(
                label="📥 Tải file Word về máy",
                data=docx_data,
                file_name=f"{original_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except subprocess.CalledProcessError as e:
            st.error(f"Lỗi Pandoc: Không chuyển được file. Vui lòng kiểm tra lại nội dung.\n{e}")
        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {e}")
        finally:
            # 5. Dọn dẹp file tạm để không làm đầy bộ nhớ server
            if 'temp_md_path' in locals() and os.path.exists(temp_md_path):
                os.remove(temp_md_path)
            if 'output_docx_path' in locals() and os.path.exists(output_docx_path):
                os.remove(output_docx_path)