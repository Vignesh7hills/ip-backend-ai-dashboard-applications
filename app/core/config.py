import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent.parent

class Settings:
    APP_NAME: str = "Vittus Fintech Backend"
    APP_VERSION: str = "1.0.0"
    DEBUG: bool = os.getenv("DEBUG", "false").lower() == "true"

    HOST: str = os.getenv("HOST", "0.0.0.0")
    PORT: int = int(os.getenv("PORT", "8000"))

    MAX_FILE_SIZE_MB: int = int(os.getenv("MAX_FILE_SIZE_MB", "50"))
    ALLOWED_EXTENSIONS: set = {".pdf", ".xlsx", ".xls", ".csv", ".xml"}

    TEMP_DIR: Path = BASE_DIR / "tmp"
    LOG_LEVEL: str = os.getenv("LOG_LEVEL", "INFO")

    COMPANY_NAME: str = os.getenv("COMPANY_NAME", "")
    FISCAL_YEAR_START_MONTH: int = 4  # April

settings = Settings()
settings.TEMP_DIR.mkdir(parents=True, exist_ok=True)
