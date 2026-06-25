"""DirectD HTTP client utilities."""

from __future__ import annotations

import json
import socket
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from api_planilhas.config import AppSettings


class DirectDError(RuntimeError):
    """Raised when DirectD communication or payload format fails."""


def build_cnpj_url(cnpj: str, settings: AppSettings) -> str:
    query = urlencode({"CNPJ": cnpj, "TOKEN": settings.directd_token})
    base_url = settings.directd_base_url.rstrip("/")
    return f"{base_url}/api/CadastroPessoaJuridica?{query}"


def build_cpf_url(cpf: str, settings: AppSettings) -> str:
    query = urlencode({"CPF": cpf, "TOKEN": settings.directd_token})
    base_url = settings.directd_base_url.rstrip("/")
    return f"{base_url}/api/CadastroPessoaFisica?{query}"


def _fetch_json(url: str, settings: AppSettings) -> dict[str, Any]:
    request = Request(
        url,
        headers={"Accept": "application/json"},
    )
    try:
        with urlopen(request, timeout=settings.directd_timeout_seconds) as response:
            raw_payload = response.read()
    except HTTPError as exc:
        raise DirectDError(f"DirectD retornou HTTP {exc.code}") from exc
    except URLError as exc:
        if isinstance(exc.reason, (TimeoutError, socket.timeout)):
            raise DirectDError("Tempo limite ao consultar DirectD") from exc
        raise DirectDError("Falha de comunicacao com DirectD") from exc
    except TimeoutError as exc:
        raise DirectDError("Tempo limite ao consultar DirectD") from exc

    try:
        payload = json.loads(raw_payload.decode("utf-8-sig"))
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise DirectDError("DirectD retornou payload JSON invalido") from exc

    if not isinstance(payload, dict):
        raise DirectDError("Resposta DirectD invalida")
    if not isinstance(payload.get("retorno"), dict):
        raise DirectDError("Resposta DirectD sem retorno esperado")

    return payload


def fetch_cnpj(cnpj: str, settings: AppSettings) -> dict[str, Any]:
    return _fetch_json(build_cnpj_url(cnpj, settings), settings)


def fetch_cpf(cpf: str, settings: AppSettings) -> dict[str, Any]:
    return _fetch_json(build_cpf_url(cpf, settings), settings)
