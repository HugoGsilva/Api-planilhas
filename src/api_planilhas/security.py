"""Security helpers for the API."""

from __future__ import annotations

import secrets

from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPBasic, HTTPBasicCredentials

from .config import get_settings

http_basic = HTTPBasic()


def require_basic_auth(
    credentials: HTTPBasicCredentials = Depends(http_basic),
) -> None:
    try:
        settings = get_settings()
    except RuntimeError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Erro de configuracao da aplicacao",
        ) from exc

    user_ok = secrets.compare_digest(credentials.username, settings.basic_user)
    password_ok = secrets.compare_digest(
        credentials.password,
        settings.basic_password,
    )

    if not (user_ok and password_ok):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid authentication credentials",
            headers={"WWW-Authenticate": "Basic"},
        )
