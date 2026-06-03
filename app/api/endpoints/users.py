"""
User management endpoints (admin-protected):
  GET    /users            — List users with search + pagination
  POST   /users            — Create a user
  GET    /users/{id}       — Get single user
  PUT    /users/{id}       — Update user
  DELETE /users/{id}       — Delete user
"""
from fastapi import APIRouter, Depends, HTTPException, Query, status
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, func, or_

from app.db.database import get_db
from app.models.user import User
from app.schemas.user import UserCreate, UserUpdate, UserResponse, UserListResponse
from app.schemas.auth import MessageResponse
from app.api.dependencies import get_current_user
from app.core.security import hash_password, generate_token
from app.core.logger import get_logger

logger = get_logger("users")

router = APIRouter(prefix="/users", tags=["User Management"])


# ── List users ────────────────────────────────────────────────────────────────
@router.get("", response_model=UserListResponse)
async def list_users(
    search: str = Query(default="", description="Search by name, email, or phone"),
    page: int = Query(default=1, ge=1),
    page_size: int = Query(default=10, ge=1, le=100),
    db: AsyncSession = Depends(get_db),
    _: User = Depends(get_current_user),
):
    query = select(User)

    if search:
        term = f"%{search}%"
        query = query.where(
            or_(
                User.first_name.ilike(term),
                User.last_name.ilike(term),
                User.email.ilike(term),
                User.phone.ilike(term),
            )
        )

    count_result = await db.execute(select(func.count()).select_from(query.subquery()))
    total = count_result.scalar_one()

    query = query.order_by(User.created_at.desc())
    query = query.offset((page - 1) * page_size).limit(page_size)
    result = await db.execute(query)
    users = result.scalars().all()

    return {
        "success": True,
        "total": total,
        "page": page,
        "page_size": page_size,
        "data": users,
    }


# ── Create user ───────────────────────────────────────────────────────────────
@router.post("", response_model=UserResponse, status_code=status.HTTP_201_CREATED)
async def create_user(
    body: UserCreate,
    db: AsyncSession = Depends(get_db),
    _: User = Depends(get_current_user),
):
    result = await db.execute(select(User).where(User.email == body.email))
    if result.scalar_one_or_none():
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail={"success": False, "message": "A user with this email already exists"},
        )

    user = User(
        first_name=body.first_name,
        last_name=body.last_name,
        email=body.email,
        phone=body.phone,
        password_hash=hash_password(body.password),
        role=body.role,
        status=body.status,
        is_verified=True,
        verification_token=None,
    )
    db.add(user)
    await db.commit()
    await db.refresh(user)
    logger.info("Admin created user %s (%s)", user.email, user.id)
    return user


# ── Get single user ───────────────────────────────────────────────────────────
@router.get("/{user_id}", response_model=UserResponse)
async def get_user(
    user_id: int,
    db: AsyncSession = Depends(get_db),
    _: User = Depends(get_current_user),
):
    result = await db.execute(select(User).where(User.id == user_id))
    user = result.scalar_one_or_none()
    if not user:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail={"success": False, "message": "User not found"},
        )
    return user


# ── Update user ───────────────────────────────────────────────────────────────
@router.put("/{user_id}", response_model=UserResponse)
async def update_user(
    user_id: int,
    body: UserUpdate,
    db: AsyncSession = Depends(get_db),
    _: User = Depends(get_current_user),
):
    result = await db.execute(select(User).where(User.id == user_id))
    user = result.scalar_one_or_none()
    if not user:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail={"success": False, "message": "User not found"},
        )

    # Email uniqueness check if changing email
    if body.email and body.email != user.email:
        existing = await db.execute(select(User).where(User.email == body.email))
        if existing.scalar_one_or_none():
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail={"success": False, "message": "Email is already taken by another user"},
            )

    updated_fields = body.model_dump(exclude_none=True)
    if "password" in updated_fields:
        user.password_hash = hash_password(updated_fields.pop("password"))

    for field, value in updated_fields.items():
        setattr(user, field, value)

    await db.commit()
    await db.refresh(user)
    logger.info("User %s updated", user_id)
    return user


# ── Delete user ───────────────────────────────────────────────────────────────
@router.delete("/{user_id}", response_model=MessageResponse)
async def delete_user(
    user_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    if current_user.id == user_id:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={"success": False, "message": "You cannot delete your own account"},
        )

    result = await db.execute(select(User).where(User.id == user_id))
    user = result.scalar_one_or_none()
    if not user:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail={"success": False, "message": "User not found"},
        )

    await db.delete(user)
    await db.commit()
    logger.info("User %s deleted by %s", user_id, current_user.id)
    return {"success": True, "message": "User deleted successfully"}
