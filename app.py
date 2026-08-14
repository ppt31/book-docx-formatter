import os
import sys
from pathlib import Path
import tempfile
import streamlit as st
from PIL import Image

# Ensure project root is in sys.path
BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from src.config import Config
from src.cover_extractor import extract_book_cover
from src.book_parser import parse_book_content
from src.docx_builder import build_docx_document
from main import create_default_logo_if_missing


def main():
    st.set_page_config(
        page_title="Book Formatter - Word (.docx)",
        page_icon="📚",
        layout="wide"
    )

    st.title("📚 Book DOCX Formatter")
    st.markdown("""
    Upload your **Book Cover** (Image/PDF/EPUB) and your **Translated Book (.docx file)**.
    The system automatically adds the **Cover Page**, **2"x2" Top-Right Logo**, **93% Transparent Center Watermark**, 
    and **Bottom-Center Page Numbers**.
    """)

    Config.ensure_directories()
    create_default_logo_if_missing(Config.LOGO_PATH)

    st.sidebar.header("⚙️ Configuration Settings")
    
    top_right_logo_size = st.sidebar.number_input(
        "Top-Right Logo Size (Inches)", 
        min_value=0.5, 
        max_value=5.0, 
        value=float(Config.TOP_RIGHT_LOGO_WIDTH_INCHES), 
        step=0.1
    )

    watermark_transparency = st.sidebar.slider(
        "Watermark Transparency (%)", 
        min_value=50, 
        max_value=99, 
        value=int(Config.WATERMARK_TRANSPARENCY_PERCENT),
        help="93% transparency means 7% subtle opacity centered behind page text."
    )

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. Upload Cover Image (or PDF/EPUB)")
        uploaded_cover = st.file_uploader(
            "Select Cover Image or PDF/EPUB to screenshot cover", 
            type=["jpg", "jpeg", "png", "pdf", "epub"]
        )

    with col2:
        st.subheader("2. Upload Translated Book File (.docx)")
        uploaded_docx = st.file_uploader(
            "Select Translated Book Content Document (.docx)", 
            type=["docx", "epub", "pdf"]
        )

    st.divider()
    st.subheader("3. Logo Image for Watermark & Header")
    uploaded_logo = st.file_uploader(
        "Upload Custom Logo PNG/JPG (Optional - Defaults to mm ENGLISH BOOKS logo)", 
        type=["png", "jpg", "jpeg"]
    )
    
    if uploaded_logo is not None:
        logo_img = Image.open(uploaded_logo)
        st.image(logo_img, caption="Custom Uploaded Logo", width=140)
    else:
        st.image(str(Config.LOGO_PATH), caption="Default Logo (mm ENGLISH BOOKS)", width=140)

    # Process when files are provided
    if uploaded_docx is not None or uploaded_cover is not None:
        st.divider()
        st.subheader("4. Cover Preview & Document Generation")
        
        with st.spinner("Processing cover and book content..."):
            # Determine book content file
            main_book_file = uploaded_docx if uploaded_docx is not None else uploaded_cover
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=Path(main_book_file.name).suffix) as tmp_book:
                tmp_book.write(main_book_file.getbuffer())
                tmp_book_path = tmp_book.name

            # Determine cover image file
            cover_path = None
            if uploaded_cover is not None:
                cover_ext = Path(uploaded_cover.name).suffix.lower()
                if cover_ext in [".png", ".jpg", ".jpeg"]:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=cover_ext) as tmp_c:
                        tmp_c.write(uploaded_cover.getbuffer())
                        cover_path = tmp_c.name
                else:
                    # PDF or EPUB uploaded as cover source
                    with tempfile.NamedTemporaryFile(delete=False, suffix=cover_ext) as tmp_c_src:
                        tmp_c_src.write(uploaded_cover.getbuffer())
                        cover_path = extract_book_cover(tmp_c_src.name)
            else:
                # Extract cover from main book file
                cover_path = extract_book_cover(tmp_book_path)

            # Save uploaded logo if provided
            if uploaded_logo is not None:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_logo:
                    tmp_logo.write(uploaded_logo.getbuffer())
                    current_logo_path = tmp_logo.name
            else:
                current_logo_path = str(Config.LOGO_PATH)

            # Display Preview & Generate Button
            c_prev1, c_prev2 = st.columns([1, 2])
            with c_prev1:
                st.markdown("#### Selected Cover Image")
                st.image(cover_path, use_container_width=True)

            with c_prev2:
                st.markdown("#### Ready to Format & Build Word Document")
                output_filename = f"{Path(main_book_file.name).stem}.docx"
                
                st.info(f"""
                - **Output Filename**: `{output_filename}`
                - **Cover Page**: Included on Page 1.
                - **Top-Right Logo**: {top_right_logo_size}" x {top_right_logo_size}" size in header.
                - **Center Watermark**: {watermark_transparency}% transparent watermark logo.
                - **Page Numbers**: Centered on bottom footer line.
                """)

                if st.button("🚀 Generate Formatted Word Document (.docx)", type="primary"):
                    with st.spinner("Generating Microsoft Word Document..."):
                        content_elements = parse_book_content(tmp_book_path)
                        output_path = Config.OUTPUT_FOLDER / output_filename

                        build_docx_document(
                            cover_image_path=cover_path,
                            content_elements=content_elements,
                            logo_path=current_logo_path,
                            output_docx_path=str(output_path),
                            top_right_logo_width=top_right_logo_size,
                            top_right_logo_height=top_right_logo_size,
                            transparency_percent=watermark_transparency
                        )

                        with open(output_path, "rb") as f:
                            docx_bytes = f.read()

                        st.success(f"Successfully formatted `{output_filename}`!")
                        st.download_button(
                            label="📥 Download Formatted .docx File",
                            data=docx_bytes,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )


if __name__ == "__main__":
    main()
