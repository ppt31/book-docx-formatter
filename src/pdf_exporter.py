import os
from pathlib import Path


def convert_docx_to_pdf(docx_path: str, pdf_path: str = None) -> str:
    """
    Converts a DOCX file to a PDF file using Microsoft Word COM Automation with docx2pdf fallback.
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

    # Strategy 1: Word COM Automation
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False

            doc = word.Documents.Open(str(docx_file))

            # 17 = wdExportFormatPDF
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
    except Exception as e_com:
        # Strategy 2: Fallback to docx2pdf package
        try:
            from docx2pdf import convert
            convert(str(docx_file), str(pdf_file))
            return str(pdf_file)
        except Exception as e_d2p:
            raise RuntimeError(
                f"Failed to export PDF via Word COM ({e_com}) and docx2pdf ({e_d2p}). "
                "Please make sure Microsoft Word is installed on Windows."
            )
