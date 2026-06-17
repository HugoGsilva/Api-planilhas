from __future__ import annotations

import time
from collections.abc import Callable
from typing import Any

from api_planilhas.cnpj import normalize_cnpj, validate_cnpj
from api_planilhas.config import AppSettings
from api_planilhas.converter import extract_rows, write_xlsx
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
    store.mark_processing(job_id)

    try:
        store.sync_progress_from_artifacts(job_id)
        cnpjs = store.get_input_cnpjs(job_id)
        completed = store.completed_cnpjs(job_id)
        delay = settings.directd_batch_delay_seconds
        fetched_once = False

        for raw_cnpj in cnpjs:
            try:
                cnpj = validate_cnpj(raw_cnpj)
            except ValueError as exc:
                error_cnpj = _error_cnpj(raw_cnpj)
                if error_cnpj not in completed:
                    store.record_cnpj_error(job_id, error_cnpj, str(exc))
                    completed.add(error_cnpj)
            else:
                if cnpj in completed:
                    continue
                try:
                    if fetched_once and delay > 0:
                        sleeper(delay)
                    fetched_once = True
                    payload = fetcher(cnpj, settings)
                except DirectDError as exc:
                    store.record_cnpj_error(job_id, cnpj, str(exc))
                else:
                    store.record_success_payload(job_id, cnpj, payload)
                completed.add(cnpj)

        rows: list[list[Any]] = []
        for payload in store.get_success_payloads(job_id):
            rows.extend(extract_rows(payload))
        write_xlsx(store.output_path(job_id), rows)
        store.mark_completed(job_id)
    except Exception as exc:
        store.mark_failed(job_id, str(exc))
