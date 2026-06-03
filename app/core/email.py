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


async def send_verification_email(email: str, otp: str) -> None:
    body = f"""
    <html><body style="font-family:Arial,sans-serif;background:#f4f4f4;padding:20px;">
    <div style="max-width:480px;margin:auto;background:#fff;border-radius:12px;padding:32px;">
      <h2 style="color:#6c47ff;">Welcome to Diginnovators!</h2>
      <p style="color:#444;">Use the OTP below to verify your email address:</p>
      <div style="text-align:center;margin:24px 0;">
        <span style="font-size:36px;font-weight:700;letter-spacing:10px;color:#6c47ff;">{otp}</span>
      </div>
      <p style="color:#888;font-size:13px;">This OTP expires in 10 minutes. Do not share it with anyone.</p>
    </div>
    </body></html>
    """
    message = MessageSchema(
        subject="Your Diginnovators verification OTP",
        recipients=[email],
        body=body,
        subtype=MessageType.html,
    )
    try:
        await _mailer.send_message(message)
    except Exception as exc:
        logger.error("Failed to send verification OTP to %s: %s", email, exc)


async def send_reset_email(email: str, otp: str) -> None:
    body = f"""
    <html><body style="font-family:Arial,sans-serif;background:#f4f4f4;padding:20px;">
    <div style="max-width:480px;margin:auto;background:#fff;border-radius:12px;padding:32px;">
      <h2 style="color:#6c47ff;">Reset Your Password</h2>
      <p style="color:#444;">Use the OTP below to reset your Diginnovators password:</p>
      <div style="text-align:center;margin:24px 0;">
        <span style="font-size:36px;font-weight:700;letter-spacing:10px;color:#6c47ff;">{otp}</span>
      </div>
      <p style="color:#888;font-size:13px;">This OTP expires in 10 minutes. If you didn't request this, ignore this email.</p>
    </div>
    </body></html>
    """
    message = MessageSchema(
        subject="Your Diginnovators password reset OTP",
        recipients=[email],
        body=body,
        subtype=MessageType.html,
    )
    try:
        await _mailer.send_message(message)
    except Exception as exc:
        logger.error("Failed to send reset OTP to %s: %s", email, exc)
