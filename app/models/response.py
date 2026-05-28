from pydantic import BaseModel
from typing import Optional, List, Any


class ProcessingLog(BaseModel):
    level: str          # INFO | WARNING | ERROR
    message: str


class ProcessResponse(BaseModel):
    success: bool
    module: str
    records_processed: int
    warnings: List[str] = []
    errors: List[str] = []
    logs: List[ProcessingLog] = []
    message: str = ""


class ValidationResult(BaseModel):
    is_valid: bool
    errors: List[str] = []
    warnings: List[str] = []


class ErrorResponse(BaseModel):
    success: bool = False
    error: str
    details: Optional[Any] = None
