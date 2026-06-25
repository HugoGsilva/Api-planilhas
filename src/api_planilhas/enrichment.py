from __future__ import annotations

import copy
from collections.abc import Callable
from typing import Any

from api_planilhas.config import AppSettings
from api_planilhas.directd import fetch_cnpj, fetch_cpf


FetchCnpj = Callable[[str, AppSettings], dict[str, Any]]
FetchCpf = Callable[[str, AppSettings], dict[str, Any]]


def _as_dict(value: Any) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


def _dict_items(value: Any) -> list[dict[str, Any]]:
    if not isinstance(value, list):
        return []
    return [item for item in value if isinstance(item, dict)]


def _digits(value: Any) -> str:
    return "".join(char for char in str(value) if char.isdigit())


def _cpf_document(value: Any) -> str:
    digits = _digits(value)
    return str(value) if len(digits) == 11 else ""


def extract_cpf_telefones(payload: Any) -> list[str]:
    retorno = _as_dict(_as_dict(payload).get("retorno"))
    telefones = []
    for item in _dict_items(retorno.get("telefones")):
        telefone = item.get("telefoneComDDD")
        if telefone not in (None, ""):
            telefones.append(str(telefone))
    return telefones


def enrich_cnpj_payload(
    cnpj: str,
    settings: AppSettings,
    fetcher: FetchCnpj = fetch_cnpj,
    cpf_fetcher: FetchCpf = fetch_cpf,
    cpf_cache: dict[str, list[str]] | None = None,
) -> dict[str, Any]:
    payload = fetcher(cnpj, settings)
    enriched = copy.deepcopy(payload)
    retorno = _as_dict(enriched.get("retorno"))
    socios = retorno.get("socios")
    if not isinstance(socios, list):
        return enriched

    cache = cpf_cache if cpf_cache is not None else {}
    for socio in socios:
        if not isinstance(socio, dict):
            continue
        cpf = _cpf_document(socio.get("documento"))
        if cpf == "":
            socio["telefonesSocio"] = []
            continue
        if cpf not in cache:
            cache[cpf] = extract_cpf_telefones(cpf_fetcher(cpf, settings))
        socio["telefonesSocio"] = cache[cpf]

    return enriched
