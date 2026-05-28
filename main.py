"""
Vittus Fintech Backend — FastAPI Entry Point

Endpoints:
  POST /api/loan-reporting-process   → Loan Reporting Excel
  POST /api/trial-balance-process    → Trial Balance Excel
  POST /api/tds-returns-process      → TDS Returns Excel
  GET  /health                        → Health check
  GET  /docs                          → Swagger UI
  GET  /redoc                         → ReDoc
"""
from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.api.endpoints.loan_reporting import router as loan_router
from app.api.endpoints.trial_balance import router as tb_router
from app.api.endpoints.tds_returns import router as tds_router
from app.core.config import settings
from app.core.logger import get_logger

logger = get_logger("main")

app = FastAPI(
    title=settings.APP_NAME,
    version=settings.APP_VERSION,
    description=(
        "Financial file processing backend. "
        "Accepts uploaded files and returns processed Excel reports. "
        "All calculations, parsing, validation, and generation happen server-side."
    ),
    docs_url="/docs",
    redoc_url="/redoc",
    openapi_url="/openapi.json",
)

# ── CORS ──────────────────────────────────────────────────────────────────────
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],          # Tighten in production
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
    expose_headers=[
        "Content-Disposition",
        "X-Records-Processed",
        "X-Warnings-Count",
        "X-Errors-Count",
        "X-Duration-Ms",
    ],
)

# ── Routers ───────────────────────────────────────────────────────────────────
app.include_router(loan_router, prefix="/api")
app.include_router(tb_router, prefix="/api")
app.include_router(tds_router, prefix="/api")


# ── Health ────────────────────────────────────────────────────────────────────
@app.get("/health", tags=["System"])
async def health_check():
    return {
        "status": "ok",
        "service": settings.APP_NAME,
        "version": settings.APP_VERSION,
    }


# ── Global error handler ──────────────────────────────────────────────────────
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    logger.exception("Unhandled exception: %s", exc)
    return JSONResponse(
        status_code=500,
        content={"success": False, "error": "internal_error", "message": str(exc)},
    )


# ── Startup ───────────────────────────────────────────────────────────────────
@app.on_event("startup")
async def startup():
    logger.info("=== %s v%s starting ===", settings.APP_NAME, settings.APP_VERSION)
    logger.info("Docs available at /docs")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host=settings.HOST,
        port=settings.PORT,
        reload=settings.DEBUG,
        log_level=settings.LOG_LEVEL.lower(),
    )
