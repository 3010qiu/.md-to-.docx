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
            # Tạo một thư mục tạm thời, nó sẽ tự động bị xóa sau khi ra khỏi block 'with'
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_dir_path = Path(temp_dir)
                
                # 1. Lưu file ZIP người dùng tải lên
                zip_path = temp_dir_path / "uploaded.zip"
                zip_path.write_bytes(uploaded_zip.read())
                
                # 2. Giải nén ZIP
                extract_dir = temp_dir_path / "extracted"
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_dir)
                
                # 3. Tìm file Markdown đầu tiên trong thư mục giải nén (bỏ qua thư mục rác của Mac)
                md_files = [f for f in extract_dir.rglob("*.md") if "__MACOSX" not in f.parts]
                
                if not md_files:
                    st.error("❌ Không tìm thấy file .md nào trong file ZIP của bạn.")
                    st.stop()
                
                main_md = md_files[0]
                
                # 4. Đọc và chuẩn hóa nội dung (thử utf-8 trước, nếu lỗi chuyển sang utf-8-sig)
                try:
                    text = main_md.read_text(encoding="utf-8")
                except UnicodeDecodeError:
                    text = main_md.read_text(encoding="utf-8-sig")
                
                text = re.sub(r"(?<!\n)\n(?!\n)", "  \n", text)
                main_md.write_text(text, encoding="utf-8")
                
                output_docx_path = temp_dir_path / "output.docx"
                
                # CHỈNH SỬA Ở ĐÂY: Tìm file template.docx trong file ZIP tải lên
                template_path = extract_dir / "template.docx"
                
                # Cấu hình mặc định của Pandoc
                pandoc_args = [
                    "-f", "markdown+tex_math_dollars+tex_math_single_backslash",
                    f"--resource-path={extract_dir}" 
                ]
                
                # Nếu phát hiện có file template, lệnh cho Pandoc áp dụng format đó
                if template_path.exists():
                    pandoc_args.append(f"--reference-doc={str(template_path)}")
                
                # 5. Gọi pypandoc
                pypandoc.convert_file(
                    str(main_md),
                    'docx',
                    outputfile=str(output_docx_path),
                    extra_args=pandoc_args
                )
                
                # 5. Gọi pypandoc, truyền thêm --resource-path để Pandoc biết đường dẫn tìm ảnh
                pypandoc.convert_file(
                    str(main_md),
                    'docx',
                    outputfile=str(output_docx_path),
                    extra_args=[
                        "-f", "markdown+tex_math_dollars+tex_math_single_backslash",
                        f"--resource-path={extract_dir}" 
                    ]
                )
                
                # 6. Đọc file Word xuất ra để tải về
                docx_data = output_docx_path.read_bytes()

            st.success("Chuyển đổi thành công! Ảnh đã được nhúng vào Word.")
            
            st.download_button(
                label="📥 Tải file Word về máy",
                data=docx_data,
                file_name=f"{original_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"Đã xảy ra lỗi: {e}")