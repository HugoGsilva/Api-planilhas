"""Configuracao da aplicacao."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from urllib.parse import urlparse


DEFAULT_DIRECTD_BASE_URL = "https://apiv3.directd.com.br"
DEFAULT_DIRECTD_TIMEOUT_SECONDS = 20.0


@dataclass(frozen=True)
class AppSettings:
    directd_token: str = field(repr=False)
    basic_user: str
    basic_password: str = field(repr=False)
    directd_base_url: str = DEFAULT_DIRECTD_BASE_URL
    directd_timeout_seconds: float = DEFAULT_DIRECTD_TIMEOUT_SECONDS


def _required(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"{name} is required")
    return value


def _optional_positive_float(name: str, default: float) -> float:
    value = os.getenv(name, "").strip()
    if not value:
        return default
    try:
        parsed = float(value)
    except ValueError as exc:
        raise RuntimeError(f"{name} must be a positive number") from exc
    if parsed <= 0:
        raise RuntimeError(f"{name} must be a positive number")
    return parsed


def _https_url(name: str, default: str) -> str:
    value = os.getenv(name, default).strip().rstrip("/")
    parsed = urlparse(value)
    if parsed.scheme != "https" or not parsed.netloc:
        raise RuntimeError(f"{name} must be an https URL")
    return value


def get_settings() -> AppSettings:
    return AppSettings(
        directd_token=_required("DIRECTD_TOKEN"),
        basic_user=_required("APP_BASIC_USER"),
        basic_password=_required("APP_BASIC_PASSWORD"),
        directd_base_url=_https_url("DIRECTD_BASE_URL", DEFAULT_DIRECTD_BASE_URL),
        directd_timeout_seconds=_optional_positive_float(
            "DIRECTD_TIMEOUT_SECONDS",
            DEFAULT_DIRECTD_TIMEOUT_SECONDS,
        ),
    )
