"""
Auth endpoints:
  POST /auth/register          — Create account, send verification email
  POST /auth/login             — Login, get JWT
  POST /auth/forgot-password   — Send password reset email
  POST /auth/reset-password    — Set new password via reset token
  POST /auth/verify-email      — Verify email with token
  POST /auth/resend-verification — Resend verification email
"""
from datetime import datetime, timedelta, timezone

from fastapi import APIRouter, Depends, HTTPException, status, BackgroundTasks
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from app.db.database import get_db
from app.models.user import User, UserStatus
from app.schemas.auth import (
    LoginRequest,
    LoginResponse,
    RegisterRequest,
    ForgotPasswordRequest,
    ResetPasswordRequest,
    VerifyEmailRequest,
    MessageResponse,
)
from app.core.security import (
    hash_password,
    verify_password,
    create_access_token,
    generate_token,
)
from app.core.email import send_verification_email, send_reset_email
from app.core.config import settings

router = APIRouter(prefix="/auth", tags=["Authentication"])


# ── Register ──────────────────────────────────────────────────────────────────
@router.post(
    "/register",
    response_model=MessageResponse,
    status_code=status.HTTP_201_CREATED,
)
async def register(
    body: RegisterRequest,
    background_tasks: BackgroundTasks,
    db: AsyncSession = Depends(get_db),
):
    result = await db.execute(select(User).where(User.email == body.email))
    if result.scalar_one_or_none():
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail={"success": False, "message": "An account with this email already exists"},
        )

    email_configured = bool(settings.MAIL_USERNAME and settings.MAIL_USERNAME != "your-email@gmail.com")
    verification_token = generate_token() if email_configured else None

    user = User(
        first_name=body.first_name,
        last_name=body.last_name,
        email=body.email,
        phone=body.phone,
        password_hash=hash_password(body.password),
        verification_token=verification_token,
        is_verified=not email_configured,
        status=UserStatus.active if not email_configured else UserStatus.pending,
    )
    db.add(user)
    await db.commit()

    if email_configured:
        background_tasks.add_task(send_verification_email, body.email, verification_token)
        message = "Account created. Please check your email to verify your account."
    else:
        message = "Account created successfully. You can now log in."

    return {"success": True, "message": message}


# ── Login ─────────────────────────────────────────────────────────────────────
@router.post("/login", response_model=LoginResponse)
async def login(
    body: LoginRequest,
    db: AsyncSession = Depends(get_db),
):
    result = await db.execute(select(User).where(User.email == body.email))
    user = result.scalar_one_or_none()

    if not user or not verify_password(body.password, user.password_hash):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail={"success": False, "message": "Invalid email or password"},
        )

    if not user.is_verified:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail={
                "success": False,
                "message": "Email not verified. Please check your inbox.",
                "code": "email_not_verified",
            },
        )

    if user.status == UserStatus.inactive:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail={"success": False, "message": "Your account has been deactivated. Contact support."},
        )

    token = create_access_token({"sub": str(user.id), "role": user.role.value})

    return {
        "success": True,
        "message": "Login successful",
        "token": token,
        "user": {
            "id": user.id,
            "first_name": user.first_name,
            "last_name": user.last_name,
            "email": user.email,
            "role": user.role.value,
            "status": user.status.value,
        },
    }


# ── Forgot Password ───────────────────────────────────────────────────────────
@router.post("/forgot-password", response_model=MessageResponse)
async def forgot_password(
    body: ForgotPasswordRequest,
    background_tasks: BackgroundTasks,
    db: AsyncSession = Depends(get_db),
):
    result = await db.execute(select(User).where(User.email == body.email))
    user = result.scalar_one_or_none()

    # Always return success to prevent email enumeration
    if user:
        reset_token = generate_token()
        user.reset_token = reset_token
        user.reset_token_expiry = datetime.now(timezone.utc) + timedelta(hours=1)
        await db.commit()
        background_tasks.add_task(send_reset_email, body.email, reset_token)

    return {
        "success": True,
        "message": "If an account exists with that email, a reset link has been sent.",
    }


# ── Reset Password ────────────────────────────────────────────────────────────
@router.post("/reset-password", response_model=MessageResponse)
async def reset_password(
    body: ResetPasswordRequest,
    db: AsyncSession = Depends(get_db),
):
    result = await db.execute(select(User).where(User.reset_token == body.token))
    user = result.scalar_one_or_none()

    if not user:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"success": False, "message": "Invalid or expired reset link"},
        )

    if user.reset_token_expiry and user.reset_token_expiry < datetime.now(timezone.utc):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"success": False, "message": "Reset link has expired. Please request a new one."},
        )

    user.password_hash = hash_password(body.new_password)
    user.reset_token = None
    user.reset_token_expiry = None
    await db.commit()

    return {"success": True, "message": "Password has been reset successfully. You can now log in."}


# ── Verify Email ──────────────────────────────────────────────────────────────
@router.post("/verify-email", response_model=MessageResponse)
async def verify_email(
    body: VerifyEmailRequest,
    db: AsyncSession = Depends(get_db),
):
    result = await db.execute(
        select(User).where(User.verification_token == body.token)
    )
    user = result.scalar_one_or_none()

    if not user:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"success": False, "message": "Invalid or already used verification link"},
        )

    user.is_verified = True
    user.status = UserStatus.active
    user.verification_token = None
    await db.commit()

    return {"success": True, "message": "Email verified successfully. You can now log in."}


# ── Resend Verification ───────────────────────────────────────────────────────
@router.post("/resend-verification", response_model=MessageResponse)
async def resend_verification(
    body: ForgotPasswordRequest,
    background_tasks: BackgroundTasks,
    db: AsyncSession = Depends(get_db),
):
    result = await db.execute(select(User).where(User.email == body.email))
    user = result.scalar_one_or_none()

    if user and not user.is_verified:
        token = generate_token()
        user.verification_token = token
        await db.commit()
        background_tasks.add_task(send_verification_email, body.email, token)

    return {
        "success": True,
        "message": "If an unverified account exists, a new verification email has been sent.",
    }
