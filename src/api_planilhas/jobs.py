from __future__ import annotations

import json
import re
import shutil
import sqlite3
import uuid
from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from decimal import Decimal
from pathlib import Path
from typing import Any


_JOB_ID_RE = re.compile(r"^[A-Za-z0-9_-]+$")


class _ClosingConnection(sqlite3.Connection):
    def __exit__(self, exc_type: Any, exc: Any, tb: Any) -> bool:
        try:
            super().__exit__(exc_type, exc, tb)
        finally:
            self.close()
        return False


@dataclass(frozen=True)
class JobSnapshot:
    job_id: str
    status: str
    total: int
    processed: int
    success: int
    errors: list[dict[str, str]]
    download_ready: bool
    created_at: str
    finished_at: str | None
    unit_price_brl: str
    cost_total_brl: str


class JobNotFoundError(KeyError):
    pass


class JobInputError(ValueError):
    pass


class JobStore:
    def __init__(self, storage_dir: Path):
        self.storage_dir = Path(storage_dir)
        self.db_path = self.storage_dir / "jobs.sqlite3"

    def initialize(self) -> None:
        self.storage_dir.mkdir(parents=True, exist_ok=True)
        with self._connect() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS jobs (
                    job_id TEXT PRIMARY KEY,
                    status TEXT NOT NULL,
                    total INTEGER NOT NULL,
                    processed INTEGER NOT NULL,
                    success INTEGER NOT NULL,
                    errors_json TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    finished_at TEXT,
                    unit_price_brl TEXT NOT NULL DEFAULT '0.00',
                    cost_total_brl TEXT NOT NULL DEFAULT '0.00'
                )
                """
            )
            self._ensure_column(connection, "unit_price_brl", "TEXT NOT NULL DEFAULT '0.00'")
            self._ensure_column(connection, "cost_total_brl", "TEXT NOT NULL DEFAULT '0.00'")

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path, factory=_ClosingConnection)
        connection.row_factory = sqlite3.Row
        return connection

    def _safe_job_id(self, job_id: str) -> str:
        if not job_id or not _JOB_ID_RE.fullmatch(job_id):
            raise JobNotFoundError(job_id)
        return job_id

    def _job_dir(self, job_id: str) -> Path:
        return self.storage_dir / self._safe_job_id(job_id)

    def input_path(self, job_id: str) -> Path:
        return self._job_dir(job_id) / "input.json"

    def output_path(self, job_id: str) -> Path:
        return self._job_dir(job_id) / "resultado.xlsx"

    def result_path(self, job_id: str, cnpj: str) -> Path:
        return self._job_dir(job_id) / "results" / f"{self._artifact_name(cnpj)}.json"

    def error_path(self, job_id: str, cnpj: str) -> Path:
        return self._job_dir(job_id) / "errors" / f"{self._artifact_name(cnpj)}.json"

    def create_job(
        self,
        cnpjs: list[str],
        status: str = "queued",
        unit_price_brl: Decimal = Decimal("0.00"),
    ) -> JobSnapshot:
        job_id = uuid.uuid4().hex
        job_dir = self._job_dir(job_id)
        job_dir.mkdir(parents=True, exist_ok=False)
        (job_dir / "results").mkdir()
        (job_dir / "errors").mkdir()
        self.input_path(job_id).write_text(
            json.dumps(cnpjs, ensure_ascii=True),
            encoding="utf-8",
        )
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO jobs (
                    job_id,
                    status,
                    total,
                    processed,
                    success,
                    errors_json,
                    created_at,
                    finished_at,
                    unit_price_brl,
                    cost_total_brl
                )
                VALUES (?, ?, ?, 0, 0, '[]', ?, NULL, ?, ?)
                """,
                (
                    job_id,
                    status,
                    len(cnpjs),
                    self._now(),
                    self._money_text(unit_price_brl),
                    self._money_text(unit_price_brl * len(cnpjs)),
                ),
            )
        return self.get_job(job_id)

    def get_input_cnpjs(self, job_id: str) -> list[str]:
        self._load_row(job_id)
        path = self.input_path(job_id)
        if not path.exists():
            raise JobNotFoundError(job_id)
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise JobInputError(f"input.json invalido para job {job_id}") from exc
        if not isinstance(data, list):
            raise JobInputError(f"input.json invalido para job {job_id}")
        return [str(item) for item in data]

    def record_success_payload(
        self,
        job_id: str,
        cnpj: str,
        payload: dict[str, Any],
    ) -> None:
        was_completed = self._artifact_name(cnpj) in self.completed_cnpjs(job_id)
        self._write_json(self.result_path(job_id, cnpj), payload)
        if not was_completed:
            self.record_success(job_id)

    def record_cnpj_error(self, job_id: str, cnpj: str, message: str) -> None:
        was_completed = self._artifact_name(cnpj) in self.completed_cnpjs(job_id)
        error = {"cnpj": cnpj, "message": message}
        self._write_json(self.error_path(job_id, cnpj), error)
        if not was_completed:
            self._append_error(job_id, error, increment_processed=True)

    def completed_cnpjs(self, job_id: str) -> set[str]:
        self._load_row(job_id)
        completed: set[str] = set()
        for directory in (self._job_dir(job_id) / "results", self._job_dir(job_id) / "errors"):
            if directory.exists():
                completed.update(path.stem for path in directory.glob("*.json"))
        return completed

    def get_success_payloads(self, job_id: str) -> list[dict[str, Any]]:
        payloads: list[dict[str, Any]] = []
        for raw_cnpj in self.get_input_cnpjs(job_id):
            path = self.result_path(job_id, raw_cnpj)
            if not path.exists():
                path = self.result_path(job_id, self._artifact_name(raw_cnpj))
            if path.exists():
                payload = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(payload, dict):
                    payloads.append(payload)
        return payloads

    def sync_progress_from_artifacts(self, job_id: str) -> None:
        self._load_row(job_id)
        result_dir = self._job_dir(job_id) / "results"
        error_dir = self._job_dir(job_id) / "errors"
        success_count = len(list(result_dir.glob("*.json"))) if result_dir.exists() else 0
        errors: list[dict[str, str]] = []
        if error_dir.exists():
            for path in sorted(error_dir.glob("*.json")):
                try:
                    data = json.loads(path.read_text(encoding="utf-8"))
                except json.JSONDecodeError:
                    continue
                if isinstance(data, dict):
                    errors.append(
                        {
                            "cnpj": str(data.get("cnpj", path.stem)),
                            "message": str(data.get("message", "")),
                        }
                    )
        with self._connect() as connection:
            connection.execute(
                """
                UPDATE jobs
                SET processed = ?,
                    success = ?,
                    errors_json = ?
                WHERE job_id = ?
                """,
                (
                    success_count + len(errors),
                    success_count,
                    json.dumps(errors, ensure_ascii=True),
                    self._safe_job_id(job_id),
                ),
            )

    def get_job(self, job_id: str) -> JobSnapshot:
        row = self._load_row(job_id)
        return self._snapshot_from_row(row)

    def list_jobs(self, limit: int = 50) -> list[JobSnapshot]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT *
                FROM jobs
                ORDER BY created_at DESC
                LIMIT ?
                """,
                (limit,),
            ).fetchall()
        return [self._snapshot_from_row(row) for row in rows]

    def list_incomplete_jobs(self) -> list[JobSnapshot]:
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT *
                FROM jobs
                WHERE status IN ('queued', 'processing')
                ORDER BY created_at DESC
                """
            ).fetchall()
        return [self._snapshot_from_row(row) for row in rows]

    def total_history_cost_brl(self, limit: int = 50) -> str:
        jobs = self.list_jobs(limit=limit)
        total = sum((Decimal(job.cost_total_brl) for job in jobs), Decimal("0.00"))
        return self._money_text(total)

    def mark_queued(self, job_id: str) -> None:
        self._execute_existing(
            "UPDATE jobs SET status = 'queued' WHERE job_id = ?",
            (self._safe_job_id(job_id),),
            job_id,
        )

    def mark_processing(self, job_id: str) -> None:
        self._execute_existing(
            "UPDATE jobs SET status = 'processing' WHERE job_id = ?",
            (self._safe_job_id(job_id),),
            job_id,
        )

    def record_success(self, job_id: str) -> None:
        self._execute_existing(
            """
            UPDATE jobs
            SET processed = processed + 1,
                success = success + 1
            WHERE job_id = ?
            """,
            (self._safe_job_id(job_id),),
            job_id,
        )

    def record_error(self, job_id: str, cnpj: str, message: str) -> None:
        self._append_error(
            job_id,
            {"cnpj": cnpj, "message": message},
            increment_processed=True,
        )

    def mark_completed(self, job_id: str) -> None:
        self._execute_existing(
            """
            UPDATE jobs
            SET status = 'completed',
                finished_at = ?
            WHERE job_id = ?
            """,
            (self._now(), self._safe_job_id(job_id)),
            job_id,
        )

    def mark_failed(self, job_id: str, message: str) -> None:
        self._append_error(
            job_id,
            {"cnpj": "", "message": message},
            status="failed",
            finished_at=self._now(),
        )

    def cleanup_expired(self, retention_hours: float) -> None:
        cutoff = datetime.now(UTC) - timedelta(hours=retention_hours)
        cutoff_text = cutoff.isoformat()
        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT job_id
                FROM jobs
                WHERE COALESCE(finished_at, created_at) < ?
                """,
                (cutoff_text,),
            ).fetchall()
        for row in rows:
            job_id = row["job_id"]
            job_dir = self._job_dir(job_id)
            if job_dir.exists():
                shutil.rmtree(job_dir)
            with self._connect() as connection:
                connection.execute(
                    "DELETE FROM jobs WHERE job_id = ?",
                    (job_id,),
                )

    def _append_error(
        self,
        job_id: str,
        error: dict[str, str],
        increment_processed: bool = False,
        status: str | None = None,
        finished_at: str | None = None,
    ) -> None:
        safe_job_id = self._safe_job_id(job_id)
        with self._connect() as connection:
            connection.execute("BEGIN IMMEDIATE")
            row = connection.execute(
                "SELECT errors_json FROM jobs WHERE job_id = ?",
                (safe_job_id,),
            ).fetchone()
            if row is None:
                raise JobNotFoundError(job_id)
            errors = self._decode_errors(row["errors_json"])
            errors.append(error)
            errors_json = json.dumps(errors, ensure_ascii=True)
            if status is None:
                connection.execute(
                    """
                    UPDATE jobs
                    SET processed = processed + ?,
                        errors_json = ?
                    WHERE job_id = ?
                    """,
                    (1 if increment_processed else 0, errors_json, safe_job_id),
                )
                return
            connection.execute(
                """
                UPDATE jobs
                SET status = ?,
                    errors_json = ?,
                    finished_at = ?
                WHERE job_id = ?
                """,
                (status, errors_json, finished_at, safe_job_id),
            )

    def _load_row(self, job_id: str) -> sqlite3.Row:
        safe_job_id = self._safe_job_id(job_id)
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM jobs WHERE job_id = ?",
                (safe_job_id,),
            ).fetchone()
        if row is None:
            raise JobNotFoundError(job_id)
        return row

    def _execute_existing(
        self,
        sql: str,
        parameters: tuple[Any, ...],
        job_id: str,
    ) -> None:
        with self._connect() as connection:
            cursor = connection.execute(sql, parameters)
            if cursor.rowcount == 0:
                raise JobNotFoundError(job_id)

    def _snapshot_from_row(self, row: sqlite3.Row) -> JobSnapshot:
        status = row["status"]
        job_id = row["job_id"]
        return JobSnapshot(
            job_id=job_id,
            status=status,
            total=row["total"],
            processed=row["processed"],
            success=row["success"],
            errors=self._decode_errors(row["errors_json"]),
            download_ready=status == "completed" and self.output_path(job_id).exists(),
            created_at=row["created_at"],
            finished_at=row["finished_at"],
            unit_price_brl=row["unit_price_brl"],
            cost_total_brl=row["cost_total_brl"],
        )

    def _decode_errors(self, errors_json: str) -> list[dict[str, str]]:
        data = json.loads(errors_json)
        if not isinstance(data, list):
            return []
        errors: list[dict[str, str]] = []
        for item in data:
            if isinstance(item, dict):
                errors.append(
                    {
                        "cnpj": str(item.get("cnpj", "")),
                        "message": str(item.get("message", "")),
                    }
                )
        return errors

    def _artifact_name(self, cnpj: str) -> str:
        digits = re.sub(r"\D+", "", str(cnpj))
        if digits:
            return digits
        value = re.sub(r"[^A-Za-z0-9_-]+", "", str(cnpj))
        return value or "sem_cnpj"

    def _write_json(self, path: Path, payload: dict[str, Any]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = path.with_suffix(".tmp")
        temp_path.write_text(
            json.dumps(payload, ensure_ascii=True),
            encoding="utf-8",
        )
        temp_path.replace(path)

    def _now(self) -> str:
        return datetime.now(UTC).isoformat()

    def _ensure_column(
        self,
        connection: sqlite3.Connection,
        column: str,
        definition: str,
    ) -> None:
        columns = {
            row["name"]
            for row in connection.execute("PRAGMA table_info(jobs)").fetchall()
        }
        if column not in columns:
            connection.execute(f"ALTER TABLE jobs ADD COLUMN {column} {definition}")

    def _money_text(self, value: Decimal) -> str:
        return str(value.quantize(Decimal("0.01")))
