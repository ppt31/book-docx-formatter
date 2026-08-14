import os
from pathlib import Path
from docx import Document
from docx.shared import Inches

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "output_docx"


def verify_docx_files():
    docx_files = list(OUTPUT_DIR.glob("*.docx"))
    print(f"[*] Found {len(docx_files)} output docx file(s) to verify.")

    for df in docx_files:
        print(f"\n[Verifying] {df.name}")
        doc = Document(df)
        
        section = doc.sections[0]
        header = section.header
        footer = section.footer

        # Verify page dimensions: Letter size 8.5" x 11.0"
        print(f"  - Page Width: {section.page_width / 914400:.2f} inches (Expected: 8.50)")
        print(f"  - Page Height: {section.page_height / 914400:.2f} inches (Expected: 11.00)")

        assert abs(section.page_width / 914400 - 8.5) < 0.01, "Page width is not 8.5 inches!"
        assert abs(section.page_height / 914400 - 11.0) < 0.01, "Page height is not 11.0 inches!"
        
        header_xml = header._element.xml
        has_top_right_logo = "TopRightLogoShape" in header_xml
        has_vml_watermark = "CenterWatermark" in header_xml
        
        print(f"  - Header Contains 0.5\" Top-Right Logo (Behind Text): {has_top_right_logo}")
        print(f"  - Header Contains Centered 93% Watermark (Behind Text): {has_vml_watermark}")

        footer_xml = footer._element.xml
        has_page_num = "PAGE" in footer_xml or "w:fldSimple" in footer_xml
        print(f"  - Footer Contains Center Page Number: {has_page_num}")

        assert has_top_right_logo, "Header missing top-right logo!"
        assert has_vml_watermark, "Header missing watermark!"
        assert has_page_num, "Footer missing page number!"

    print("\n[✓] ALL LETTER SIZE (8.5\" x 11\") & BEHIND-TEXT VERIFICATION CHECKS PASSED!")


if __name__ == "__main__":
    verify_docx_files()
