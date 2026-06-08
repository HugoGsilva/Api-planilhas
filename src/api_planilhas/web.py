"""Application factory and small web utilities."""

from __future__ import annotations

from pathlib import Path

from fastapi import Depends, FastAPI, Form, HTTPException, status
from fastapi.responses import HTMLResponse, Response

from api_planilhas.config import AppSettings, get_settings
from api_planilhas.converter import build_xlsx_bytes, extract_row
from api_planilhas.directd import DirectDError, fetch_cnpj
from .security import require_basic_auth


TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "index.html"


def normalize_cnpj(cnpj: str) -> str:
    return "".join(ch for ch in (cnpj or "") if ch.isdigit())


def validate_cnpj(cnpj: str) -> str:
    value = normalize_cnpj(cnpj)
    if len(value) != 14:
        raise ValueError("CNPJ deve conter 14 digitos")
    return value


def create_app() -> FastAPI:
    app = FastAPI()

    @app.get("/", response_class=HTMLResponse, dependencies=[Depends(require_basic_auth)])
    async def homepage() -> str:
        return TEMPLATE_PATH.read_text(encoding="utf-8")

    @app.post("/api/planilha")
    def gerar_planilha(
        cnpj: str = Form(...),
        settings: AppSettings = Depends(get_settings),
        _: None = Depends(require_basic_auth),
    ) -> Response:
        try:
            normalized = validate_cnpj(cnpj)
        except ValueError as exc:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=str(exc),
            ) from exc

        try:
            payload = fetch_cnpj(normalized, settings)
        except DirectDError as exc:
            raise HTTPException(
                status_code=status.HTTP_502_BAD_GATEWAY,
                detail=str(exc),
            ) from exc

        content = build_xlsx_bytes([extract_row(payload)])
        filename = f"importacao_oportunidade_{normalized}.xlsx"
        return Response(
            content=content,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    return app


app = create_app()
