"""Configuracao da aplicacao."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urlparse


DEFAULT_DIRECTD_BASE_URL = "https://apiv3.directd.com.br"
DEFAULT_DIRECTD_TIMEOUT_SECONDS = 20.0
DEFAULT_JOB_RETENTION_HOURS = 48.0
DEFAULT_UPLOAD_MAX_MB = 100
DEFAULT_DIRECTD_BATCH_DELAY_SECONDS = 0.0
DEFAULT_JOB_STORAGE_DIR = Path("storage/jobs")


@dataclass(frozen=True)
class AppSettings:
    directd_token: str = field(repr=False)
    basic_user: str
    basic_password: str = field(repr=False)
    directd_base_url: str = DEFAULT_DIRECTD_BASE_URL
    directd_timeout_seconds: float = DEFAULT_DIRECTD_TIMEOUT_SECONDS
    job_storage_dir: Path = DEFAULT_JOB_STORAGE_DIR
    job_retention_hours: float = DEFAULT_JOB_RETENTION_HOURS
    upload_max_mb: int = DEFAULT_UPLOAD_MAX_MB
    directd_batch_delay_seconds: float = DEFAULT_DIRECTD_BATCH_DELAY_SECONDS


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


def _optional_positive_int(name: str, default: int) -> int:
    value = os.getenv(name, "").strip()
    if not value:
        return default
    try:
        parsed = int(value)
    except ValueError as exc:
        raise RuntimeError(f"{name} must be a positive integer") from exc
    if parsed <= 0:
        raise RuntimeError(f"{name} must be a positive integer")
    return parsed


def _optional_non_negative_float(name: str, default: float) -> float:
    value = os.getenv(name, "").strip()
    if not value:
        return default
    try:
        parsed = float(value)
    except ValueError as exc:
        raise RuntimeError(f"{name} must be a non-negative number") from exc
    if parsed < 0:
        raise RuntimeError(f"{name} must be a non-negative number")
    return parsed


def _optional_path(name: str, default: Path) -> Path:
    value = os.getenv(name, "").strip()
    return Path(value) if value else default


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
        job_storage_dir=_optional_path("JOB_STORAGE_DIR", DEFAULT_JOB_STORAGE_DIR),
        job_retention_hours=_optional_positive_float(
            "JOB_RETENTION_HOURS",
            DEFAULT_JOB_RETENTION_HOURS,
        ),
        upload_max_mb=_optional_positive_int("UPLOAD_MAX_MB", DEFAULT_UPLOAD_MAX_MB),
        directd_batch_delay_seconds=_optional_non_negative_float(
            "DIRECTD_BATCH_DELAY_SECONDS",
            DEFAULT_DIRECTD_BATCH_DELAY_SECONDS,
        ),
    )
