import os
import zipfile
from pathlib import Path
from bs4 import BeautifulSoup
import ebooklib
from ebooklib import epub
import fitz  # PyMuPDF
import docx


def parse_epub_content(epub_path: str) -> list[dict]:
    """
    Parses EPUB chapters into structured content elements (headings, paragraphs) with fallback.
    """
    content_elements = []

    # Strategy 1: Try ebooklib
    try:
        book = epub.read_epub(str(epub_path))
        for item in book.get_items():
            if item.get_type() == ebooklib.ITEM_DOCUMENT:
                html_content = item.get_content().decode("utf-8", errors="ignore")
                soup = BeautifulSoup(html_content, "html.parser")
                
                for elem in soup.find_all(['h1', 'h2', 'h3', 'h4', 'p']):
                    text = elem.get_text().strip()
                    if not text:
                        continue
                    tag_name = elem.name
                    if tag_name in ['h1', 'h2', 'h3', 'h4']:
                        content_elements.append({"type": "heading", "level": int(tag_name[1]), "text": text})
                    else:
                        content_elements.append({"type": "paragraph", "text": text})
        if content_elements:
            return content_elements
    except Exception:
        pass

    # Strategy 2: Direct Zipfile parsing
    try:
        with zipfile.ZipFile(str(epub_path), 'r') as z:
            html_files = [f for f in z.namelist() if f.endswith(('.xhtml', '.html', '.htm'))]
            for hf in html_files:
                html_content = z.read(hf).decode("utf-8", errors="ignore")
                soup = BeautifulSoup(html_content, "html.parser")
                for elem in soup.find_all(['h1', 'h2', 'h3', 'h4', 'p']):
                    text = elem.get_text().strip()
                    if not text:
                        continue
                    tag_name = elem.name
                    if tag_name in ['h1', 'h2', 'h3', 'h4']:
                        content_elements.append({"type": "heading", "level": int(tag_name[1]), "text": text})
                    else:
                        content_elements.append({"type": "paragraph", "text": text})
    except Exception as e:
        print(f"[!] Warning parsing EPUB via zip: {e}")

    return content_elements


def parse_pdf_content(pdf_path: str) -> list[dict]:
    """
    Parses PDF pages (excluding page 0 cover) into text paragraphs.
    """
    doc = fitz.open(str(pdf_path))
    content_elements = []

    # Skip page 0 (cover page)
    for page_num in range(1, len(doc)):
        page = doc[page_num]
        text = page.get_text("text")
        lines = [line.strip() for line in text.split("\n") if line.strip()]
        
        for line in lines:
            if len(line) < 50 and not line.endswith("."):
                content_elements.append({"type": "heading", "level": 2, "text": line})
            else:
                content_elements.append({"type": "paragraph", "text": line})

    doc.close()
    return content_elements


def parse_docx_content(docx_path: str) -> list[dict]:
    """
    Parses existing Word .docx file into structured paragraphs and headings.
    """
    doc = docx.Document(str(docx_path))
    content_elements = []

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        
        style_name = p.style.name.lower()
        if "heading" in style_name or "title" in style_name:
            content_elements.append({"type": "heading", "level": 2, "text": text})
        else:
            content_elements.append({"type": "paragraph", "text": text})

    return content_elements


def parse_book_content(book_path: str) -> list[dict]:
    """
    Parses book content based on file extension (.epub, .pdf, or .docx).
    """
    ext = Path(book_path).suffix.lower()
    if ext == ".epub":
        return parse_epub_content(book_path)
    elif ext == ".pdf":
        return parse_pdf_content(book_path)
    elif ext == ".docx":
        return parse_docx_content(book_path)
    else:
        raise ValueError(f"Unsupported file format '{ext}'")
