import os
from pathlib import Path
import pythoncom
import win32com.client


def convert_docx_to_pdf(docx_path: str, pdf_path: str = None) -> str:
    """
    Converts a DOCX file to a PDF file using Microsoft Word COM Automation.
    Ensures exact 1:1 fidelity with Word formatting, watermarks, headers, and Burmese fonts.
    """
    docx_file = Path(docx_path).resolve()
    if not docx_file.exists():
        raise FileNotFoundError(f"DOCX file not found: {docx_path}")

    if pdf_path is None:
        pdf_file = docx_file.with_suffix(".pdf")
    else:
        pdf_file = Path(pdf_path).resolve()

    pdf_file.parent.mkdir(parents=True, exist_ok=True)

    # Initialize COM library for current thread (required in web threads like Streamlit)
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False

        # Open DOCX in Word
        doc = word.Documents.Open(str(docx_file))

        # 17 = wdExportFormatPDF
        # Export as PDF with high quality
        doc.ExportAsFixedFormat(
            OutputFileName=str(pdf_file),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,  # wdExportOptimizeForPrint
            CreateBookmarks=1,  # wdExportCreateHeadingBookmarks
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )
        return str(pdf_file)
    finally:
        if doc is not None:
            doc.Close(SaveChanges=False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()
