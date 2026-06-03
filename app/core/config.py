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
    FISCAL_YEAR_START_MONTH: int = 4

    # ── PostgreSQL ────────────────────────────────────────────────────────────
    # Force asyncpg driver — handles plain postgresql:// URLs (e.g. from Heroku/env)
    _db_url: str = os.getenv(
        "DATABASE_URL",
        "postgresql+asyncpg://postgres:password@localhost:5432/diginnovators",
    )
    if _db_url.startswith("postgresql://"):
        _db_url = _db_url.replace("postgresql://", "postgresql+asyncpg://", 1)
    DATABASE_URL: str = _db_url

    # ── JWT ───────────────────────────────────────────────────────────────────
    JWT_SECRET: str = os.getenv("JWT_SECRET", "change-this-secret-in-production")
    JWT_ALGORITHM: str = "HS256"
    JWT_EXPIRE_MINUTES: int = int(os.getenv("JWT_EXPIRE_MINUTES", "60"))

    # ── Frontend ──────────────────────────────────────────────────────────────
    FRONTEND_URL: str = os.getenv("FRONTEND_URL", "http://localhost:3000")

settings = Settings()
settings.TEMP_DIR.mkdir(parents=True, exist_ok=True)
