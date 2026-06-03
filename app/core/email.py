from fastapi_mail import FastMail, MessageSchema, ConnectionConfig, MessageType
from app.core.config import settings
from app.core.logger import get_logger

logger = get_logger("email")

_conf = ConnectionConfig(
    MAIL_USERNAME=settings.MAIL_USERNAME,
    MAIL_PASSWORD=settings.MAIL_PASSWORD,
    MAIL_FROM=settings.MAIL_FROM,
    MAIL_PORT=settings.MAIL_PORT,
    MAIL_SERVER=settings.MAIL_SERVER,
    MAIL_STARTTLS=settings.MAIL_STARTTLS,
    MAIL_SSL_TLS=settings.MAIL_SSL_TLS,
    USE_CREDENTIALS=True,
    VALIDATE_CERTS=True,
)

_mailer = FastMail(_conf)


async def send_verification_email(email: str, token: str) -> None:
    url = f"{settings.FRONTEND_URL}/verify-email?token={token}"
    body = f"""
    <html><body>
    <h2>Welcome to Diginnovators!</h2>
    <p>Please verify your email address by clicking the button below:</p>
    <a href="{url}" style="
        display:inline-block;padding:12px 24px;
        background:#6c47ff;color:#fff;
        border-radius:8px;text-decoration:none;font-weight:600;">
      Verify Email
    </a>
    <p style="margin-top:16px;color:#666;">
      This link expires in 24 hours. If you did not create an account, ignore this email.
    </p>
    </body></html>
    """
    message = MessageSchema(
        subject="Verify your Diginnovators account",
        recipients=[email],
        body=body,
        subtype=MessageType.html,
    )
    try:
        await _mailer.send_message(message)
    except Exception as exc:
        logger.error("Failed to send verification email to %s: %s", email, exc)


async def send_reset_email(email: str, token: str) -> None:
    url = f"{settings.FRONTEND_URL}/reset-password?token={token}"
    body = f"""
    <html><body>
    <h2>Reset Your Password</h2>
    <p>Click the button below to set a new password for your Diginnovators account:</p>
    <a href="{url}" style="
        display:inline-block;padding:12px 24px;
        background:#6c47ff;color:#fff;
        border-radius:8px;text-decoration:none;font-weight:600;">
      Reset Password
    </a>
    <p style="margin-top:16px;color:#666;">
      This link expires in 1 hour. If you didn't request a reset, ignore this email.
    </p>
    </body></html>
    """
    message = MessageSchema(
        subject="Reset your Diginnovators password",
        recipients=[email],
        body=body,
        subtype=MessageType.html,
    )
    try:
        await _mailer.send_message(message)
    except Exception as exc:
        logger.error("Failed to send reset email to %s: %s", email, exc)
