from __future__ import annotations

import time
from collections.abc import Callable
from typing import Any

from api_planilhas.cnpj import normalize_cnpj, validate_cnpj
from api_planilhas.config import AppSettings
from api_planilhas.converter import extract_row, write_xlsx
from api_planilhas.directd import DirectDError, fetch_cnpj
from api_planilhas.jobs import JobStore


FetchCnpj = Callable[[str, AppSettings], dict[str, Any]]
Sleeper = Callable[[float], None]


def _error_cnpj(raw_cnpj: str) -> str:
    return normalize_cnpj(raw_cnpj) or str(raw_cnpj)


def process_job(
    job_id: str,
    store: JobStore,
    settings: AppSettings,
    fetcher: FetchCnpj = fetch_cnpj,
    sleeper: Sleeper = time.sleep,
) -> None:
    rows: list[list[Any]] = []
    store.mark_processing(job_id)

    try:
        cnpjs = store.get_input_cnpjs(job_id)
        delay = settings.directd_batch_delay_seconds

        for index, raw_cnpj in enumerate(cnpjs):
            try:
                cnpj = validate_cnpj(raw_cnpj)
            except ValueError as exc:
                store.record_error(job_id, _error_cnpj(raw_cnpj), str(exc))
            else:
                try:
                    payload = fetcher(cnpj, settings)
                except DirectDError as exc:
                    store.record_error(job_id, cnpj, str(exc))
                else:
                    rows.append(extract_row(payload))
                    store.record_success(job_id)

            if delay > 0 and index < len(cnpjs) - 1:
                sleeper(delay)

        write_xlsx(store.output_path(job_id), rows)
        store.mark_completed(job_id)
    except Exception as exc:
        store.mark_failed(job_id, str(exc))
