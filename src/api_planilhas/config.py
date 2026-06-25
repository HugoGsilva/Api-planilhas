"""Configuracao da aplicacao."""

from __future__ import annotations

import math
import os
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from urllib.parse import urlparse


DEFAULT_DIRECTD_BASE_URL = "https://apiv3.directd.com.br"
DEFAULT_DIRECTD_TIMEOUT_SECONDS = 20.0
DEFAULT_JOB_RETENTION_HOURS = 168.0
DEFAULT_UPLOAD_MAX_MB = 100
DEFAULT_DIRECTD_BATCH_DELAY_SECONDS = 0.0
DEFAULT_JOB_STORAGE_DIR = Path("storage/jobs")
DEFAULT_CNPJ_QUERY_UNIT_PRICE_BRL = Decimal("0.16")
# Fixed contracted price per CadastroPessoaJuridica query. Keep this value in
# one place so future price changes are easy to update.
CNPJ_QUERY_UNIT_PRICE_SOURCE = (
    "Preco fixo configurado para CadastroPessoaJuridica: R$ 0,16 por CNPJ."
)


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
    cnpj_query_unit_price_brl: Decimal = DEFAULT_CNPJ_QUERY_UNIT_PRICE_BRL


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
    if not math.isfinite(parsed) or parsed <= 0:
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
    if not math.isfinite(parsed) or parsed < 0:
        raise RuntimeError(f"{name} must be a non-negative number")
    return parsed


def _optional_non_negative_decimal(name: str, default: Decimal) -> Decimal:
    value = os.getenv(name, "").strip().replace(",", ".")
    if not value:
        return default
    try:
        parsed = Decimal(value)
    except InvalidOperation as exc:
        raise RuntimeError(f"{name} must be a non-negative decimal") from exc
    if parsed < 0:
        raise RuntimeError(f"{name} must be a non-negative decimal")
    return parsed.quantize(Decimal("0.01"))


def _optional_path(name: str, default: Path) -> Path:
    value = os.getenv(name, "").strip()
    if not value:
        return default
    path = Path(value)
    if path.is_absolute() or path.drive or path.root or ".." in path.parts:
        raise RuntimeError(f"{name} must be a relative path without parent traversal")
    return path


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
        cnpj_query_unit_price_brl=_optional_non_negative_decimal(
            "CNPJ_QUERY_UNIT_PRICE_BRL",
            DEFAULT_CNPJ_QUERY_UNIT_PRICE_BRL,
        ),
    )
