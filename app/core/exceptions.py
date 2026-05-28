class FinanceBackendError(Exception):
    """Base exception for all finance backend errors."""
    def __init__(self, message: str, details: dict = None):
        super().__init__(message)
        self.message = message
        self.details = details or {}


class FileParseError(FinanceBackendError):
    """Raised when a file cannot be parsed."""


class UnsupportedFileTypeError(FinanceBackendError):
    """Raised when an unsupported file type is uploaded."""


class ValidationError(FinanceBackendError):
    """Raised when data fails validation rules."""


class CalculationError(FinanceBackendError):
    """Raised when a financial calculation fails."""


class EmptyFileError(FinanceBackendError):
    """Raised when a file has no usable content."""


class MalformedDataError(FinanceBackendError):
    """Raised when file data is structurally broken."""
