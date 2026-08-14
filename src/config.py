import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env file from base directory
BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")


class Config:
    BASE_DIR = BASE_DIR
    INPUT_FOLDER = Path(os.getenv("INPUT_FOLDER", BASE_DIR / "input_books")).resolve()
    OUTPUT_FOLDER = Path(os.getenv("OUTPUT_FOLDER", BASE_DIR / "output_docx")).resolve()
    LOGO_PATH = Path(os.getenv("LOGO_PATH", BASE_DIR / "assets" / "logo.png")).resolve()

    TOP_RIGHT_LOGO_WIDTH_INCHES = float(os.getenv("TOP_RIGHT_LOGO_WIDTH_INCHES", "2.0"))
    TOP_RIGHT_LOGO_HEIGHT_INCHES = float(os.getenv("TOP_RIGHT_LOGO_HEIGHT_INCHES", "2.0"))

    WATERMARK_TRANSPARENCY_PERCENT = float(os.getenv("WATERMARK_TRANSPARENCY_PERCENT", "93"))
    WATERMARK_SCALE_PERCENT = float(os.getenv("WATERMARK_SCALE_PERCENT", "60"))

    PAGE_NUMBER_POSITION = os.getenv("PAGE_NUMBER_POSITION", "center_bottom")

    @classmethod
    def ensure_directories(cls):
        cls.INPUT_FOLDER.mkdir(parents=True, exist_ok=True)
        cls.OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
        cls.LOGO_PATH.parent.mkdir(parents=True, exist_ok=True)
