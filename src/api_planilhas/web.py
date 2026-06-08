"""Application factory and small web utilities."""

from __future__ import annotations

from pathlib import Path

from fastapi import (
    BackgroundTasks,
    Depends,
    FastAPI,
    Form,
    HTTPException,
    Request,
    UploadFile,
    status,
)
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response

from api_planilhas.batch_processor import process_job
from api_planilhas.cnpj import normalize_cnpj, validate_cnpj
from api_planilhas.config import AppSettings, get_settings
from api_planilhas.converter import build_xlsx_bytes, extract_row
from api_planilhas.directd import DirectDError, fetch_cnpj
from api_planilhas.jobs import JobNotFoundError, JobStore
from api_planilhas.xlsx_reader import (
    InvalidXlsxError,
    MissingCnpjColumnError,
    extract_cnpj_values,
)
from .security import require_basic_auth


TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "index.html"
UPLOAD_CHUNK_SIZE = 1024 * 1024


def _get_job_store(app: FastAPI, settings: AppSettings) -> JobStore:
    store = getattr(app.state, "job_store", None)
    if store is not None and store.storage_dir == settings.job_storage_dir:
        return store

    store = JobStore(settings.job_storage_dir)
    store.initialize()
    store.cleanup_expired(settings.job_retention_hours)
    app.state.job_store = store
    return store


async def _read_upload_with_limit(
    file: UploadFile,
    max_bytes: int,
    chunk_size: int = UPLOAD_CHUNK_SIZE,
) -> bytes:
    content = bytearray()
    while True:
        chunk = await file.read(chunk_size)
        if not chunk:
            return bytes(content)
        if len(content) + len(chunk) > max_bytes:
            raise HTTPException(
                status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                detail="Arquivo acima do limite permitido",
            )
        content.extend(chunk)


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

    @app.post("/api/lotes", dependencies=[Depends(require_basic_auth)])
    async def criar_lote(
        request: Request,
        background_tasks: BackgroundTasks,
        file: UploadFile,
        settings: AppSettings = Depends(get_settings),
    ) -> JSONResponse:
        max_bytes = settings.upload_max_mb * 1024 * 1024
        content = await _read_upload_with_limit(file, max_bytes)

        try:
            cnpjs = extract_cnpj_values(content)
        except MissingCnpjColumnError as exc:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=str(exc),
            ) from exc
        except InvalidXlsxError as exc:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=str(exc),
            ) from exc

        store = _get_job_store(request.app, settings)
        job = store.create_job(cnpjs)
        background_tasks.add_task(process_job, job.job_id, store, settings)
        return JSONResponse(
            status_code=status.HTTP_202_ACCEPTED,
            content={"job_id": job.job_id},
        )

    @app.get("/api/lotes/{job_id}", dependencies=[Depends(require_basic_auth)])
    def consultar_lote(
        request: Request,
        job_id: str,
        settings: AppSettings = Depends(get_settings),
    ) -> dict:
        store = _get_job_store(request.app, settings)
        try:
            job = store.get_job(job_id)
        except JobNotFoundError as exc:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Job nao encontrado",
            ) from exc
        return {
            "job_id": job.job_id,
            "status": job.status,
            "total": job.total,
            "processed": job.processed,
            "success": job.success,
            "errors": job.errors,
            "download_ready": job.download_ready,
        }

    @app.get("/api/lotes/{job_id}/download", dependencies=[Depends(require_basic_auth)])
    def baixar_lote(
        request: Request,
        job_id: str,
        settings: AppSettings = Depends(get_settings),
    ) -> FileResponse:
        store = _get_job_store(request.app, settings)
        try:
            job = store.get_job(job_id)
        except JobNotFoundError as exc:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Job nao encontrado",
            ) from exc
        if not job.download_ready:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="Planilha ainda nao esta pronta",
            )
        return FileResponse(
            store.output_path(job_id),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=f"importacao_oportunidades_lote_{job_id}.xlsx",
        )

    return app


app = create_app()
