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
from src.book_parser import parse_book_content, elements_to_text, text_to_elements
from src.docx_builder import build_docx_document, format_existing_docx
from src.pdf_exporter import convert_docx_to_pdf
from main import create_default_logo_if_missing


def main():
    st.set_page_config(
        page_title="Book Formatter - Word (.docx) & PDF Exporter",
        page_icon="📚",
        layout="wide"
    )

    st.title("📚 Book Formatter & PDF Exporter")
    st.markdown("""
    Upload your **Book Cover** and **Book File (.docx / .pdf / .epub)**.
    Format your book with **Letter size (8.5" x 11")**, **2"x2" top-right logo**, **93% transparent center watermark**, 
    and **bottom-center page numbers**, while **preserving all embedded photos, images, and tables**!
    """)

    Config.ensure_directories()
    create_default_logo_if_missing(Config.LOGO_PATH)

    # --- SIDEBAR CONFIGURATION ---
    st.sidebar.header("⚙️ Document Settings")
    
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

    st.sidebar.markdown("---")
    st.sidebar.info("""
    **Layout Specifications:**
    - Fixed Letter Size (8.5" x 11")
    - 2"x2" Top-Right Logo (0.5" from top/right, behind text)
    - 93% Center Watermark (behind text)
    - Dynamic Page Numbers (bottom-center)
    - Embedded Photos & Tables: 100% Preserved
    """)

    # --- UPLOAD SECTION ---
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("1. Upload Book Cover Image (or PDF/EPUB)")
        uploaded_cover = st.file_uploader(
            "Select Cover Image or PDF/EPUB to screenshot cover", 
            type=["jpg", "jpeg", "png", "pdf", "epub"]
        )

    with col2:
        st.subheader("2. Upload Book Content Document (.docx)")
        uploaded_docx = st.file_uploader(
            "Select Book Document (.docx / .epub / .pdf)", 
            type=["docx", "epub", "pdf"]
        )

    st.divider()

    # --- LOGO SECTION ---
    with st.expander("🖼️ Logo Settings (Defaults to mm ENGLISH BOOKS logo)", expanded=False):
        uploaded_logo = st.file_uploader(
            "Upload Custom Logo (Optional)", 
            type=["png", "jpg", "jpeg"]
        )
        if uploaded_logo is not None:
            logo_img = Image.open(uploaded_logo)
            st.image(logo_img, caption="Custom Uploaded Logo", width=140)
        else:
            st.image(str(Config.LOGO_PATH), caption="Current Logo (assets/logo.png)", width=140)

    # --- PROCESSING WORKFLOW ---
    main_book_file = uploaded_docx if uploaded_docx is not None else uploaded_cover

    if main_book_file is not None:
        file_ext = Path(main_book_file.name).suffix.lower()
        file_key = f"book_content_{main_book_file.name}"
        
        # Save temp copy of uploaded book file
        with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp_book:
            tmp_book.write(main_book_file.getbuffer())
            tmp_book_path = tmp_book.name

        # Extract Cover Image
        cover_path = None
        if uploaded_cover is not None:
            cover_ext = Path(uploaded_cover.name).suffix.lower()
            if cover_ext in [".png", ".jpg", ".jpeg"]:
                with tempfile.NamedTemporaryFile(delete=False, suffix=cover_ext) as tmp_c:
                    tmp_c.write(uploaded_cover.getbuffer())
                    cover_path = tmp_c.name
            else:
                with tempfile.NamedTemporaryFile(delete=False, suffix=cover_ext) as tmp_c_src:
                    tmp_c_src.write(uploaded_cover.getbuffer())
                    cover_path = extract_book_cover(tmp_c_src.name)
        else:
            cover_path = extract_book_cover(tmp_book_path)

        # Parse text for optional text editor
        if file_key not in st.session_state:
            initial_elements = parse_book_content(tmp_book_path)
            st.session_state[file_key] = elements_to_text(initial_elements)

        # Logo path
        if uploaded_logo is not None:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_logo:
                tmp_logo.write(uploaded_logo.getbuffer())
                current_logo_path = tmp_logo.name
        else:
            current_logo_path = str(Config.LOGO_PATH)

        # Layout: Cover on Left, Options & In-browser Editor on Right
        st.subheader("3. Book Preview & Content Editor")
        c_left, c_right = st.columns([1, 2])

        with c_left:
            st.markdown("#### 📖 Book Cover Preview")
            if cover_path and Path(cover_path).exists():
                st.image(cover_path, use_container_width=True)
            else:
                st.info("No cover preview available")

        with c_right:
            if file_ext == ".docx":
                st.success("✅ **Original DOCX Detected**: Direct Formatting Mode is active — all embedded photos, diagrams, tables, and layouts inside your document will be **100% preserved**.")
                
                edit_mode = st.radio(
                    "Formatting Mode:",
                    [
                        "Preserve Full Original DOCX (Keeps all embedded photos, tables & layout)",
                        "Rebuild from In-Browser Text Editor"
                    ],
                    index=0
                )
            else:
                edit_mode = "Rebuild from In-Browser Text Editor"

            with st.expander("✏️ View / Edit Text Content in Browser", expanded=(edit_mode == "Rebuild from In-Browser Text Editor")):
                st.caption("You can modify text, sentences, or chapter titles below.")
                edited_text = st.text_area(
                    "Book Text Content",
                    value=st.session_state[file_key],
                    height=350
                )
                st.session_state[file_key] = edited_text

                if st.button("🔄 Reset to Original Text"):
                    initial_elements = parse_book_content(tmp_book_path)
                    st.session_state[file_key] = elements_to_text(initial_elements)
                    st.rerun()

        # --- 4. EXPORT ACTIONS (DOCX & PDF) ---
        st.divider()
        st.subheader("4. Generate & Export Formatted Files")

        output_filename = f"{Path(main_book_file.name).stem}.docx"
        pdf_filename = f"{Path(main_book_file.name).stem}.pdf"

        col_export1, col_export2 = st.columns(2)

        with col_export1:
            st.markdown("#### 📝 Export Word Document (.docx)")
            if st.button("🚀 Build & Download Word Document (.docx)", type="primary", use_container_width=True):
                with st.spinner("Formatting Word document (preserving all photos & layout)..."):
                    output_path = Config.OUTPUT_FOLDER / output_filename

                    if file_ext == ".docx" and edit_mode.startswith("Preserve Full Original"):
                        # Direct formatting: 100% preserves all embedded photos & tables
                        format_existing_docx(
                            input_docx_path=tmp_book_path,
                            cover_image_path=cover_path,
                            logo_path=current_logo_path,
                            output_docx_path=str(output_path),
                            top_right_logo_width=top_right_logo_size,
                            top_right_logo_height=top_right_logo_size,
                            transparency_percent=watermark_transparency
                        )
                    else:
                        # Rebuild from content elements
                        content_elements = text_to_elements(st.session_state[file_key])
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

                    st.session_state[f"docx_ready_{file_key}"] = docx_bytes
                    st.success(f"Generated `{output_filename}` with all photos preserved!")

            if f"docx_ready_{file_key}" in st.session_state:
                st.download_button(
                    label="📥 Download .docx Document",
                    data=st.session_state[f"docx_ready_{file_key}"],
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

        with col_export2:
            st.markdown("#### 📄 Export PDF Document (.pdf)")
            if st.button("📑 Build & Export PDF Document (.pdf)", type="secondary", use_container_width=True):
                with st.spinner("Generating document and exporting to PDF (with all photos & watermark)..."):
                    output_path = Config.OUTPUT_FOLDER / output_filename

                    if file_ext == ".docx" and edit_mode.startswith("Preserve Full Original"):
                        format_existing_docx(
                            input_docx_path=tmp_book_path,
                            cover_image_path=cover_path,
                            logo_path=current_logo_path,
                            output_docx_path=str(output_path),
                            top_right_logo_width=top_right_logo_size,
                            top_right_logo_height=top_right_logo_size,
                            transparency_percent=watermark_transparency
                        )
                    else:
                        content_elements = text_to_elements(st.session_state[file_key])
                        build_docx_document(
                            cover_image_path=cover_path,
                            content_elements=content_elements,
                            logo_path=current_logo_path,
                            output_docx_path=str(output_path),
                            top_right_logo_width=top_right_logo_size,
                            top_right_logo_height=top_right_logo_size,
                            transparency_percent=watermark_transparency
                        )

                    # Export to PDF via Word COM / docx2pdf
                    pdf_path = Config.OUTPUT_FOLDER / pdf_filename
                    convert_docx_to_pdf(str(output_path), str(pdf_path))

                    with open(pdf_path, "rb") as f:
                        pdf_bytes = f.read()

                    st.session_state[f"pdf_ready_{file_key}"] = pdf_bytes
                    st.success(f"Exported `{pdf_filename}` with all photos & watermarks intact!")

            if f"pdf_ready_{file_key}" in st.session_state:
                st.download_button(
                    label="📄 Download .pdf Document",
                    data=st.session_state[f"pdf_ready_{file_key}"],
                    file_name=pdf_filename,
                    mime="application/pdf",
                    use_container_width=True
                )


if __name__ == "__main__":
    main()
