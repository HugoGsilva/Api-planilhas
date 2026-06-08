from __future__ import annotations

import json
import re
import shutil
import sqlite3
import uuid
from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
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


class JobNotFoundError(KeyError):
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
                    finished_at TEXT
                )
                """
            )

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

    def create_job(self, cnpjs: list[str]) -> JobSnapshot:
        job_id = uuid.uuid4().hex
        job_dir = self._job_dir(job_id)
        job_dir.mkdir(parents=True, exist_ok=False)
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
                    finished_at
                )
                VALUES (?, 'queued', ?, 0, 0, '[]', ?, NULL)
                """,
                (job_id, len(cnpjs), self._now()),
            )
        return self.get_job(job_id)

    def get_input_cnpjs(self, job_id: str) -> list[str]:
        self._load_row(job_id)
        path = self.input_path(job_id)
        if not path.exists():
            raise JobNotFoundError(job_id)
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, list):
            return []
        return [str(item) for item in data]

    def get_job(self, job_id: str) -> JobSnapshot:
        row = self._load_row(job_id)
        return self._snapshot_from_row(row)

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
        row = self._load_row(job_id)
        errors = self._decode_errors(row["errors_json"])
        errors.append({"cnpj": cnpj, "message": message})
        self._execute_existing(
            """
            UPDATE jobs
            SET processed = processed + 1,
                errors_json = ?
            WHERE job_id = ?
            """,
            (json.dumps(errors, ensure_ascii=True), self._safe_job_id(job_id)),
            job_id,
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
        row = self._load_row(job_id)
        errors = self._decode_errors(row["errors_json"])
        errors.append({"cnpj": "", "message": message})
        self._execute_existing(
            """
            UPDATE jobs
            SET status = 'failed',
                errors_json = ?,
                finished_at = ?
            WHERE job_id = ?
            """,
            (
                json.dumps(errors, ensure_ascii=True),
                self._now(),
                self._safe_job_id(job_id),
            ),
            job_id,
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
            job_ids = [row["job_id"] for row in rows]
            connection.executemany(
                "DELETE FROM jobs WHERE job_id = ?",
                [(job_id,) for job_id in job_ids],
            )

        for job_id in job_ids:
            shutil.rmtree(self._job_dir(job_id), ignore_errors=True)

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

    def _now(self) -> str:
        return datetime.now(UTC).isoformat()
