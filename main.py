import os
import sys
import argparse
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

# Add project root to sys.path
BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from src.config import Config
from src.cover_extractor import extract_book_cover
from src.book_parser import parse_book_content
from src.docx_builder import build_docx_document


def create_default_logo_if_missing(logo_path: Path):
    """
    Generates a sample logo if user hasn't placed logo.png in assets folder yet.
    """
    if logo_path.exists():
        return

    logo_path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGBA", (400, 400), color=(0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    draw.ellipse([20, 20, 380, 380], fill=(41, 128, 185, 255), outline=(255, 255, 255, 255), width=8)
    draw.text((130, 160), "MM BOOKS\nLOGO", fill=(255, 255, 255, 255), font_size=36)

    img.save(logo_path, format="PNG")
    print(f"[*] Created default sample logo at: {logo_path}")


def process_book_file(book_file_path: Path, logo_path: Path, output_dir: Path, cover_file_path: Path = None) -> Path:
    """
    Orchestrates the conversion of a book file into a formatted DOCX document.
    Can take an explicit cover image/file or extract cover automatically.
    """
    print(f"\n==========================================")
    print(f"[>] Processing book file: {book_file_path.name}")
    print(f"==========================================")

    # 1. Extract / Screenshot cover page
    if cover_file_path and cover_file_path.is_file():
        cover_ext = cover_file_path.suffix.lower()
        if cover_ext in [".png", ".jpg", ".jpeg"]:
            cover_path = str(cover_file_path)
        else:
            print("[1/4] Extracting/Screenshotting cover from specified cover file...")
            cover_path = extract_book_cover(str(cover_file_path))
    else:
        print("[1/4] Extracting/Screenshotting book cover page...")
        cover_path = extract_book_cover(str(book_file_path))
    
    print(f"      Cover image path: {cover_path}")

    # 2. Parse text content
    print("[2/4] Parsing text content & chapters...")
    content_elements = parse_book_content(str(book_file_path))
    print(f"      Extracted {len(content_elements)} content elements.")

    # 3. Output file path (using exact book stem name)
    output_docx_name = f"{book_file_path.stem}.docx"
    output_docx_path = output_dir / output_docx_name

    # 4. Generate formatted DOCX
    print("[3/4] Building Word (.docx) file with Cover, 2\" Logo, 93% Watermark & Page Numbers...")
    final_docx = build_docx_document(
        cover_image_path=cover_path,
        content_elements=content_elements,
        logo_path=str(logo_path),
        output_docx_path=str(output_docx_path),
        top_right_logo_width=Config.TOP_RIGHT_LOGO_WIDTH_INCHES,
        top_right_logo_height=Config.TOP_RIGHT_LOGO_HEIGHT_INCHES,
        transparency_percent=Config.WATERMARK_TRANSPARENCY_PERCENT
    )

    print(f"[4/4] SUCCESS! Output file generated:\n      {final_docx}\n")
    return Path(final_docx)


def main():
    parser = argparse.ArgumentParser(description="Format EPUB, PDF or DOCX Books into Word (.docx) with Cover, Logo, Watermark & Page Numbers")
    parser.add_argument("--input", "-i", type=str, help="Input EPUB, PDF or DOCX file path or directory")
    parser.add_argument("--cover", "-c", type=str, help="Explicit Cover image or EPUB/PDF file for cover")
    parser.add_argument("--logo", "-l", type=str, help="Logo image path (defaults to assets/logo.png)")
    parser.add_argument("--output", "-o", type=str, help="Output directory (defaults to output_docx/)")
    
    args = parser.parse_args()

    Config.ensure_directories()

    logo_path = Path(args.logo) if args.logo else Config.LOGO_PATH
    create_default_logo_if_missing(logo_path)

    input_path = Path(args.input) if args.input else Config.INPUT_FOLDER
    output_dir = Path(args.output) if args.output else Config.OUTPUT_FOLDER
    cover_path = Path(args.cover) if args.cover else None

    book_files = []
    if input_path.is_file():
        book_files.append(input_path)
    elif input_path.is_dir():
        for ext in ("*.epub", "*.pdf", "*.docx"):
            book_files.extend(list(input_path.glob(ext)))

    if not book_files:
        print(f"[!] No .epub, .pdf, or .docx files found in: {input_path}")
        print(f"[!] Please place your book files in '{Config.INPUT_FOLDER}' or pass --input <file_path>.")
        return

    print(f"[*] Found {len(book_files)} book file(s) to process.")
    for book_file in book_files:
        if input_path == output_dir and book_file.suffix == ".docx":
            continue
        try:
            process_book_file(book_file, logo_path, output_dir, cover_path)
        except Exception as e:
            print(f"[X] Error processing '{book_file.name}': {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    main()
