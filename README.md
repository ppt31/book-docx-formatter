# 📚 Book DOCX Formatter (EPUB / PDF / DOCX to Word)

A Python application for converting **EPUB**, **PDF**, or **DOCX** books (English or bilingual English-Burmese) into professionally formatted Microsoft Word (`.docx`) documents.

---

## ✨ Features

- 📸 **Automatic Cover Screenshot/Extraction**: Extracts cover page image from EPUB/PDF page 0 or standalone cover images.
- 📐 **Letter Page Dimensions (8.5" x 11.0")**: Formats document pages to standard Letter size.
- 🏷️ **2" x 2" Top-Right Logo**: Positions header logo **0.5" from top/right corner** floating **behind text**.
- 💧 **93% Transparent Center Watermark**: Pre-processes logo to 93% transparency (7% opacity) centered behind document text.
- 🔢 **Bottom-Center Page Numbers**: Embeds dynamic Microsoft Word page numbers centered at the bottom line.
- 🌐 **Dual Interface**:
  - **Web UI** (`streamlit run app.py`): Drag & drop interface with cover preview.
  - **Command Line (CLI)** (`python main.py`): Batch processing script.

---

## 🚀 Quick Start

### 1. Setup Virtual Environment (`venv`)
Run the setup script or manually install requirements:
```cmd
setup_env.bat
```
*(Or in PowerShell)*:
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 2. Launch Web Interface (Recommended)
```powershell
streamlit run app.py
```
*(Or double-click `run_web.bat`)*

### 3. Launch Command Line Batch Formatter
Place your `.epub`, `.pdf`, or `.docx` files in `input_books/` and run:
```powershell
python main.py
```

---

## ⚙️ Configuration (`.env`)

Edit `.env` to customize default settings:
```env
INPUT_FOLDER=./input_books
OUTPUT_FOLDER=./output_docx
LOGO_PATH=./assets/logo.png

TOP_RIGHT_LOGO_WIDTH_INCHES=2.0
TOP_RIGHT_LOGO_HEIGHT_INCHES=2.0
WATERMARK_TRANSPARENCY_PERCENT=93
PAGE_NUMBER_POSITION=center_bottom
```

---

## 📂 Project Structure

```
book_docx_formatter/
├── .env                    # Environment configuration
├── .env.example            # Environment template
├── requirements.txt        # Python dependencies
├── setup_env.bat           # Environment setup script
├── run_cli.bat / .ps1      # CLI launch scripts
├── run_web.bat / .ps1      # Web UI launch scripts
├── assets/
│   └── logo.png            # Default logo image
├── input_books/            # Input folder for EPUB, PDF & DOCX books
├── output_docx/           # Output folder for generated formatted .docx files
├── src/
│   ├── config.py           # Configuration loader
│   ├── cover_extractor.py  # Cover extraction & rendering
│   ├── watermark.py        # 93% transparency PIL processor
│   ├── book_parser.py      # EPUB/PDF/DOCX content parser
│   └── docx_builder.py     # Microsoft Word document builder
├── main.py                 # CLI entry point
└── app.py                  # Streamlit Web UI
```

---

## 📜 License
MIT License
