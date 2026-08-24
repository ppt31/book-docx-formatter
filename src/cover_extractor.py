import os
import io
import zipfile
from pathlib import Path
from PIL import Image
try:
    import fitz  # PyMuPDF
except ImportError:
    import pymupdf as fitz
import ebooklib
from ebooklib import epub
import docx


def extract_cover_from_pdf(pdf_path: str, output_image_path: str = None) -> str:
    """
    Screenshots / renders page 0 of a PDF file as high-quality cover image.
    """
    pdf_path = Path(pdf_path)
    if output_image_path is None:
        cache_dir = pdf_path.parent / ".cache"
        cache_dir.mkdir(parents=True, exist_ok=True)
        output_image_path = cache_dir / f"cover_{pdf_path.stem}.png"

    doc = fitz.open(str(pdf_path))
    if len(doc) == 0:
        raise ValueError(f"PDF document '{pdf_path}' is empty.")

    page = doc[0]  # Page 0 is the front cover
    mat = fitz.Matrix(2.0, 2.0)
    pix = page.get_pixmap(matrix=mat)
    
    pix.save(str(output_image_path))
    doc.close()
    return str(output_image_path)


def extract_cover_from_epub(epub_path: str, output_image_path: str = None) -> str:
    """
    Extracts cover image from EPUB using ebooklib with zipfile fallback.
    """
    epub_path = Path(epub_path)
    if output_image_path is None:
        cache_dir = epub_path.parent / ".cache"
        cache_dir.mkdir(parents=True, exist_ok=True)
        output_image_path = cache_dir / f"cover_{epub_path.stem}.png"

    cover_bytes = None

    # Strategy 1: Try ebooklib read
    try:
        book = epub.read_epub(str(epub_path))
        for item in book.get_items():
            if item.get_type() == ebooklib.ITEM_COVER:
                cover_bytes = item.get_content()
                break
            elif item.get_type() == ebooklib.ITEM_IMAGE:
                name_lower = (item.get_name() or "").lower()
                if "cover" in name_lower:
                    cover_bytes = item.get_content()
                    break
    except Exception:
        pass

    # Strategy 2: Zipfile direct search
    if cover_bytes is None:
        try:
            with zipfile.ZipFile(str(epub_path), 'r') as z:
                cover_names = [f for f in z.namelist() if any(k in f.lower() for k in ["cover", "front"])]
                image_names = [f for f in z.namelist() if f.lower().endswith((".png", ".jpg", ".jpeg"))]
                
                target_file = cover_names[0] if cover_names else (image_names[0] if image_names else None)
                if target_file:
                    cover_bytes = z.read(target_file)
        except Exception:
            pass

    if cover_bytes is not None:
        img = Image.open(io.BytesIO(cover_bytes))
        img.save(str(output_image_path))
        return str(output_image_path)

    # Strategy 3: Create fallback cover image
    img = Image.new("RGB", (600, 800), color=(240, 240, 240))
    img.save(str(output_image_path))
    return str(output_image_path)


def extract_cover_from_docx(docx_path: str, output_image_path: str = None) -> str:
    """
    Extracts cover image for a DOCX file.
    Checks:
    1. Direct image cover files in same folder ({stem}.jpg, {stem}.png, {stem}_cover.png)
    2. Matching PDF or EPUB book in same folder ({stem}.pdf, {stem}.epub)
    3. Embedded picture inside DOCX
    """
    docx_path = Path(docx_path)
    if output_image_path is None:
        cache_dir = docx_path.parent / ".cache"
        cache_dir.mkdir(parents=True, exist_ok=True)
        output_image_path = cache_dir / f"cover_{docx_path.stem}.png"

    parent_dir = docx_path.parent
    stem = docx_path.stem

    # 1. Check for standalone cover image with matching name in same folder
    possible_cover_images = [
        parent_dir / f"{stem}.jpg",
        parent_dir / f"{stem}.png",
        parent_dir / f"{stem}.jpeg",
        parent_dir / f"{stem}_cover.jpg",
        parent_dir / f"{stem}_cover.png",
        parent_dir / f"{stem}_cover.jpeg",
    ]
    for img_p in possible_cover_images:
        if img_p.exists():
            img = Image.open(img_p)
            img.save(str(output_image_path))
            return str(output_image_path)

    # 2. Check for matching PDF/EPUB book in same folder
    pdf_match = parent_dir / f"{stem}.pdf"
    epub_match = parent_dir / f"{stem}.epub"
    if pdf_match.exists():
        return extract_cover_from_pdf(str(pdf_match), str(output_image_path))
    elif epub_match.exists():
        return extract_cover_from_epub(str(epub_match), str(output_image_path))

    # 3. Check for embedded picture inside DOCX
    try:
        doc = docx.Document(str(docx_path))
        for rel in doc.part.rels.values():
            if "image" in rel.target_ref:
                img_bytes = rel.target_part.blob
                img = Image.open(io.BytesIO(img_bytes))
                img.save(str(output_image_path))
                return str(output_image_path)
    except Exception:
        pass

    # Fallback placeholder cover
    img = Image.new("RGB", (600, 800), color=(240, 240, 240))
    img.save(str(output_image_path))
    return str(output_image_path)


def extract_book_cover(book_path: str, output_image_path: str = None) -> str:
    """
    Extracts/screenshots cover from PDF, EPUB, or DOCX file automatically based on extension.
    """
    ext = Path(book_path).suffix.lower()
    if ext == ".pdf":
        return extract_cover_from_pdf(book_path, output_image_path)
    elif ext == ".epub":
        return extract_cover_from_epub(book_path, output_image_path)
    elif ext == ".docx":
        return extract_cover_from_docx(book_path, output_image_path)
    else:
        raise ValueError(f"Unsupported book format '{ext}'. Expected .pdf, .epub, or .docx.")
