from __future__ import annotations


def normalize_cnpj(cnpj: str) -> str:
    return "".join(ch for ch in (cnpj or "") if ch.isdigit())


def validate_cnpj(cnpj: str) -> str:
    value = normalize_cnpj(cnpj)
    if len(value) != 14:
        raise ValueError("CNPJ deve conter 14 digitos")
    return value
