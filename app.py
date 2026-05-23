import os
import re
import zipfile
import tempfile
import streamlit as st
import pypandoc
from pathlib import Path

st.set_page_config(page_title="Markdown to Word (Hỗ trợ Ảnh)", page_icon="📝")
st.title("Chuyển đổi ZIP (Markdown + Ảnh) sang Word")

@st.cache_resource
def setup_pandoc():
    try:
        pypandoc.get_pandoc_version()
    except OSError:
        with st.spinner("Đang khởi tạo Pandoc cho máy chủ..."):
            pypandoc.download_pandoc()

setup_pandoc()

st.write("Tải lên file **.zip** chứa file Markdown và các thư mục hình ảnh kèm theo.")

uploaded_zip = st.file_uploader("Chọn file .zip", type=["zip"])

if uploaded_zip is not None:
    original_name = os.path.splitext(uploaded_zip.name)[0]
    
    with st.spinner('Đang giải nén và xử lý hình ảnh...'):
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_dir_path = Path(temp_dir)
                
                # 1. Lưu và giải nén ZIP
                zip_path = temp_dir_path / "uploaded.zip"
                zip_path.write_bytes(uploaded_zip.read())
                
                extract_dir = temp_dir_path / "extracted"
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_dir)
                
                # 2. Tìm file Markdown ở bất kỳ đâu trong ZIP
                md_files = [f for f in extract_dir.rglob("*.md") if "__MACOSX" not in f.parts]
                
                if not md_files:
                    st.error("❌ Không tìm thấy file .md nào trong file ZIP của bạn.")
                    st.stop()
                
                main_md = md_files[0]
                
                # QUAN TRỌNG: Xác định chính xác thư mục đang chứa file Markdown
                md_folder = main_md.parent 
                
                # 3. Chuẩn hóa nội dung
                try:
                    text = main_md.read_text(encoding="utf-8")
                except UnicodeDecodeError:
                    text = main_md.read_text(encoding="utf-8-sig")
                
                text = re.sub(r"(?<!\n)\n(?!\n)", "  \n", text)
                main_md.write_text(text, encoding="utf-8")
                
                output_docx_path = temp_dir_path / "output.docx"
                
                # 4. Tìm file template.docx ở MỌI NGÓC NGÁCH trong file nén
                template_files = list(extract_dir.rglob("template.docx"))
                
                # 5. Cấu hình Pandoc
                # Dùng thư mục md_folder làm gốc để tìm ảnh (nhờ vậy ![](hinh1.png) mới hoạt động)
                pandoc_args = [
                    "-f", "markdown+tex_math_dollars+tex_math_single_backslash",
                    f"--resource-path={str(md_folder)}"
                ]
                
                # Nếu tìm thấy file template, báo cho Pandoc biết
                if template_files:
                    template_path = template_files[0]
                    pandoc_args.append(f"--reference-doc={str(template_path)}")
                
                # 6. Chuyển đổi
                pypandoc.convert_file(
                    str(main_md),
                    'docx',
                    outputfile=str(output_docx_path),
                    extra_args=pandoc_args
                )
                
                # 7. Tải file Word đã hoàn thiện
                docx_data = output_docx_path.read_bytes()

            st.success("Chuyển đổi thành công! Ảnh và Template đã được áp dụng.")
            
            st.download_button(
                label="📥 Tải file Word về máy",
                data=docx_data,
                file_name=f"{original_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {e}")