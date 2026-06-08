# Lote CNPJ com Fila Assincrona Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Permitir upload de planilhas XLSX com coluna `CNPJ`, processar consultas DirectD em fila local e entregar o XLSX final para download posterior.

**Architecture:** A aplicacao FastAPI continuara monolitica, com uma fila local em SQLite e arquivos em `storage/jobs`. Um worker em background dentro do mesmo processo processara um job por vez, consultando a DirectD sequencialmente e gravando progresso, erros e planilha final.

**Tech Stack:** Python 3.14, FastAPI, SQLite stdlib, `zipfile`/`xml.etree.ElementTree` stdlib para leitura XLSX, gerador XLSX existente em `api_planilhas.converter`, unittest.

---

## File Structure

- Modify `src/api_planilhas/config.py`: adicionar configuracoes de lote (`JOB_STORAGE_DIR`, `JOB_RETENTION_HOURS`, `UPLOAD_MAX_MB`, `DIRECTD_BATCH_DELAY_SECONDS`) com validacao.
- Create `src/api_planilhas/cnpj.py`: centralizar `normalize_cnpj` e `validate_cnpj` para evitar dependencias circulares entre web e processamento.
- Create `src/api_planilhas/xlsx_reader.py`: ler a primeira aba de XLSX, localizar coluna `CNPJ` e extrair valores.
- Create `src/api_planilhas/jobs.py`: persistencia SQLite, modelos simples de job/status/erro e limpeza de retencao.
- Create `src/api_planilhas/batch_processor.py`: processar um job, chamar DirectD sequencialmente, gerar XLSX final e registrar erros.
- Modify `src/api_planilhas/web.py`: endpoints `POST /api/lotes`, `GET /api/lotes/{job_id}`, `GET /api/lotes/{job_id}/download` e inicializacao da fila.
- Modify `src/api_planilhas/templates/index.html`: UI de upload, status, erros e download do lote, mantendo fluxo individual.
- Create `tests/test_xlsx_reader.py`: cobertura de extracao de CNPJ e erro obrigatorio.
- Create `tests/test_jobs.py`: cobertura de SQLite store, status e cleanup.
- Create `tests/test_batch_processor.py`: cobertura do processamento com sucessos/falhas.
- Modify `tests/test_web.py`: cobertura dos endpoints de lote e UI.

---

### Task 1: Configuracao de lote

**Files:**
- Modify: `src/api_planilhas/config.py`
- Modify: `tests/test_web.py`

- [ ] **Step 1: Write failing tests for batch config defaults**

Add these imports/expectations in `tests/test_web.py`:

```python
from pathlib import Path
from api_planilhas.config import (
    DEFAULT_DIRECTD_BASE_URL,
    DEFAULT_DIRECTD_TIMEOUT_SECONDS,
    DEFAULT_JOB_RETENTION_HOURS,
    DEFAULT_UPLOAD_MAX_MB,
    get_settings,
)
```

Add this test in `ConfigAndSecurityTest`:

```python
    def test_get_settings_reads_batch_defaults(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"

        settings = get_settings()

        self.assertEqual(settings.job_storage_dir, Path("storage/jobs"))
        self.assertEqual(settings.job_retention_hours, DEFAULT_JOB_RETENTION_HOURS)
        self.assertEqual(settings.upload_max_mb, DEFAULT_UPLOAD_MAX_MB)
        self.assertEqual(settings.directd_batch_delay_seconds, 0.0)
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_web.ConfigAndSecurityTest.test_get_settings_reads_batch_defaults -v
```

Expected: FAIL because `DEFAULT_JOB_RETENTION_HOURS` or new `AppSettings` fields are missing.

- [ ] **Step 3: Implement batch settings**

Update `src/api_planilhas/config.py`:

```python
from pathlib import Path
from urllib.parse import urlparse

DEFAULT_JOB_RETENTION_HOURS = 48.0
DEFAULT_UPLOAD_MAX_MB = 100
DEFAULT_DIRECTD_BATCH_DELAY_SECONDS = 0.0
DEFAULT_JOB_STORAGE_DIR = Path("storage/jobs")


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
```

Add helpers:

```python
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
    if parsed < 0:
        raise RuntimeError(f"{name} must be a non-negative number")
    return parsed


def _optional_path(name: str, default: Path) -> Path:
    value = os.getenv(name, "").strip()
    return Path(value) if value else default
```

Update `get_settings()`:

```python
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
```

- [ ] **Step 4: Add validation tests for env overrides**

Add tests:

```python
    def test_get_settings_reads_batch_env_overrides(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"
        os.environ["JOB_STORAGE_DIR"] = "custom/jobs"
        os.environ["JOB_RETENTION_HOURS"] = "72"
        os.environ["UPLOAD_MAX_MB"] = "250"
        os.environ["DIRECTD_BATCH_DELAY_SECONDS"] = "0.5"

        settings = get_settings()

        self.assertEqual(settings.job_storage_dir, Path("custom/jobs"))
        self.assertEqual(settings.job_retention_hours, 72.0)
        self.assertEqual(settings.upload_max_mb, 250)
        self.assertEqual(settings.directd_batch_delay_seconds, 0.5)

    def test_get_settings_rejects_invalid_upload_max_mb(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"
        os.environ["UPLOAD_MAX_MB"] = "0"

        with self.assertRaises(RuntimeError) as ctx:
            get_settings()

        self.assertIn("UPLOAD_MAX_MB", str(ctx.exception))
```

- [ ] **Step 5: Run config tests**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_web.ConfigAndSecurityTest -v
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add src/api_planilhas/config.py tests/test_web.py
git commit -m "Adiciona configuracao de lote"
```

---

### Task 2: Utilitario de CNPJ e leitor XLSX para coluna CNPJ

**Files:**
- Create: `src/api_planilhas/cnpj.py`
- Modify: `src/api_planilhas/web.py`
- Create: `src/api_planilhas/xlsx_reader.py`
- Create: `tests/test_xlsx_reader.py`

- [ ] **Step 1: Move CNPJ helpers into dedicated module**

Create `src/api_planilhas/cnpj.py`:

```python
from __future__ import annotations


def normalize_cnpj(cnpj: str) -> str:
    return "".join(ch for ch in (cnpj or "") if ch.isdigit())


def validate_cnpj(cnpj: str) -> str:
    value = normalize_cnpj(cnpj)
    if len(value) != 14:
        raise ValueError("CNPJ deve conter 14 digitos")
    return value
```

Update `src/api_planilhas/web.py`:

```python
from api_planilhas.cnpj import normalize_cnpj, validate_cnpj
```

Remove the existing local `normalize_cnpj()` and `validate_cnpj()` definitions from `web.py`.

- [ ] **Step 2: Run existing CNPJ tests**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_web.ConfigAndSecurityTest.test_normalize_cnpj_removes_all_non_digits tests.test_web.ConfigAndSecurityTest.test_validate_cnpj_accepts_14_digits tests.test_web.ConfigAndSecurityTest.test_validate_cnpj_rejects_non_14_digits -v
```

Expected: PASS.

- [ ] **Step 3: Write failing XLSX reader tests**

Create `tests/test_xlsx_reader.py`:

```python
import sys
import unittest
import zipfile
from io import BytesIO
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from api_planilhas.xlsx_reader import (
    MISSING_CNPJ_COLUMN_MESSAGE,
    MissingCnpjColumnError,
    extract_cnpj_values,
)


def build_minimal_xlsx(rows):
    def cell_ref(row_index, col_index):
        return f"{chr(64 + col_index)}{row_index}"

    row_xml = []
    for row_index, row in enumerate(rows, start=1):
        cells = []
        for col_index, value in enumerate(row, start=1):
            ref = cell_ref(row_index, col_index)
            cells.append(
                f'<c r="{ref}" t="inlineStr"><is><t>{value}</t></is></c>'
            )
        row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    sheet = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        '</worksheet>'
    )
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("xl/worksheets/sheet1.xml", sheet)
    return buffer.getvalue()


class XlsxReaderTest(unittest.TestCase):
    def test_extracts_cnpj_column_values(self):
        content = build_minimal_xlsx(
            [
                ["Nome", " CNPJ "],
                ["Empresa A", "12.345.678/0001-90"],
                ["Empresa B", "11222333000144"],
                ["Empresa C", ""],
            ]
        )

        values = extract_cnpj_values(content)

        self.assertEqual(values, ["12.345.678/0001-90", "11222333000144"])

    def test_rejects_spreadsheet_without_cnpj_column(self):
        content = build_minimal_xlsx([["Nome", "Documento"], ["Empresa A", "123"]])

        with self.assertRaises(MissingCnpjColumnError) as ctx:
            extract_cnpj_values(content)

        self.assertEqual(str(ctx.exception), MISSING_CNPJ_COLUMN_MESSAGE)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 4: Run test to verify it fails**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_xlsx_reader -v
```

Expected: FAIL because `api_planilhas.xlsx_reader` does not exist.

- [ ] **Step 5: Implement XLSX reader**

Create `src/api_planilhas/xlsx_reader.py`:

```python
from __future__ import annotations

import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO

MISSING_CNPJ_COLUMN_MESSAGE = (
    "nao tem cnpj na sua planilha adicione uma coluna CNPJ e os cnpj em baixo"
)

SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


class MissingCnpjColumnError(ValueError):
    pass


class InvalidXlsxError(ValueError):
    pass


def _cell_column(cell_ref: str) -> int:
    letters = "".join(char for char in cell_ref if char.isalpha())
    result = 0
    for char in letters:
        result = result * 26 + (ord(char.upper()) - 64)
    return result


def _cell_text(cell: ET.Element) -> str:
    texts = [
        node.text or ""
        for node in cell.findall(f".//{{{SPREADSHEET_NS}}}t")
    ]
    return "".join(texts).strip()


def _read_rows(content: bytes) -> list[list[str]]:
    try:
        with zipfile.ZipFile(BytesIO(content)) as archive:
            sheet_xml = archive.read("xl/worksheets/sheet1.xml")
    except (KeyError, zipfile.BadZipFile) as exc:
        raise InvalidXlsxError("Planilha XLSX invalida") from exc

    root = ET.fromstring(sheet_xml)
    rows: list[list[str]] = []
    for row in root.findall(f".//{{{SPREADSHEET_NS}}}row"):
        values_by_column: dict[int, str] = {}
        max_column = 0
        for cell in row.findall(f"./{{{SPREADSHEET_NS}}}c"):
            ref = cell.attrib.get("r", "")
            column = _cell_column(ref)
            if column <= 0:
                continue
            values_by_column[column] = _cell_text(cell)
            max_column = max(max_column, column)
        rows.append([values_by_column.get(index, "") for index in range(1, max_column + 1)])
    return rows


def extract_cnpj_values(content: bytes) -> list[str]:
    rows = _read_rows(content)
    if not rows:
        raise MissingCnpjColumnError(MISSING_CNPJ_COLUMN_MESSAGE)

    header = [value.strip().upper() for value in rows[0]]
    try:
        cnpj_index = header.index("CNPJ")
    except ValueError as exc:
        raise MissingCnpjColumnError(MISSING_CNPJ_COLUMN_MESSAGE) from exc

    values: list[str] = []
    for row in rows[1:]:
        if cnpj_index >= len(row):
            continue
        value = row[cnpj_index].strip()
        if value:
            values.append(value)
    return values
```

- [ ] **Step 6: Run XLSX reader tests**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_xlsx_reader -v
```

Expected: PASS.

- [ ] **Step 7: Commit**

```powershell
git add src/api_planilhas/cnpj.py src/api_planilhas/web.py src/api_planilhas/xlsx_reader.py tests/test_xlsx_reader.py
git commit -m "Adiciona leitura de CNPJ em XLSX"
```

---

### Task 3: Store SQLite de jobs

**Files:**
- Create: `src/api_planilhas/jobs.py`
- Create: `tests/test_jobs.py`
- Modify: `.gitignore`

- [ ] **Step 1: Write failing job store tests**

Create `tests/test_jobs.py`:

```python
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from api_planilhas.jobs import JobStore


class JobStoreTest(unittest.TestCase):
    def test_creates_and_reads_job(self):
        with TemporaryDirectory() as temp_dir:
            store = JobStore(Path(temp_dir))
            store.initialize()

            job = store.create_job(["12345678000190", "11222333000144"])
            loaded = store.get_job(job.job_id)

            self.assertEqual(loaded.job_id, job.job_id)
            self.assertEqual(loaded.status, "queued")
            self.assertEqual(loaded.total, 2)
            self.assertEqual(loaded.processed, 0)
            self.assertEqual(loaded.success, 0)
            self.assertEqual(loaded.errors, [])

    def test_updates_progress_and_records_error(self):
        with TemporaryDirectory() as temp_dir:
            store = JobStore(Path(temp_dir))
            store.initialize()
            job = store.create_job(["12345678000190"])

            store.mark_processing(job.job_id)
            store.record_success(job.job_id)
            store.record_error(job.job_id, "123", "CNPJ deve conter 14 digitos")
            store.mark_completed(job.job_id)
            loaded = store.get_job(job.job_id)

            self.assertEqual(loaded.status, "completed")
            self.assertEqual(loaded.processed, 2)
            self.assertEqual(loaded.success, 1)
            self.assertEqual(
                loaded.errors,
                [{"cnpj": "123", "message": "CNPJ deve conter 14 digitos"}],
            )


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_jobs -v
```

Expected: FAIL because `api_planilhas.jobs` does not exist.

- [ ] **Step 3: Implement JobStore**

Create `src/api_planilhas/jobs.py` with:

```python
from __future__ import annotations

import json
import shutil
import sqlite3
import uuid
from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any


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
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        return connection

    def _job_dir(self, job_id: str) -> Path:
        return self.storage_dir / job_id

    def input_path(self, job_id: str) -> Path:
        return self._job_dir(job_id) / "input.json"

    def output_path(self, job_id: str) -> Path:
        return self._job_dir(job_id) / "resultado.xlsx"

    def create_job(self, cnpjs: list[str]) -> JobSnapshot:
        job_id = uuid.uuid4().hex
        self._job_dir(job_id).mkdir(parents=True, exist_ok=False)
        self.input_path(job_id).write_text(json.dumps(cnpjs), encoding="utf-8")
        created_at = datetime.now(UTC).isoformat()
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO jobs (job_id, status, total, processed, success, errors_json, created_at, finished_at)
                VALUES (?, 'queued', ?, 0, 0, '[]', ?, NULL)
                """,
                (job_id, len(cnpjs), created_at),
            )
        return self.get_job(job_id)

    def get_input_cnpjs(self, job_id: str) -> list[str]:
        path = self.input_path(job_id)
        if not path.exists():
            raise JobNotFoundError(job_id)
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []

    def get_job(self, job_id: str) -> JobSnapshot:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM jobs WHERE job_id = ?",
                (job_id,),
            ).fetchone()
        if row is None:
            raise JobNotFoundError(job_id)
        errors = json.loads(row["errors_json"])
        return JobSnapshot(
            job_id=row["job_id"],
            status=row["status"],
            total=row["total"],
            processed=row["processed"],
            success=row["success"],
            errors=errors,
            download_ready=self.output_path(job_id).exists() and row["status"] == "completed",
        )

    def mark_processing(self, job_id: str) -> None:
        self._update(job_id, "UPDATE jobs SET status = 'processing' WHERE job_id = ?")

    def mark_completed(self, job_id: str) -> None:
        finished_at = datetime.now(UTC).isoformat()
        self._update(
            job_id,
            "UPDATE jobs SET status = 'completed', finished_at = ? WHERE job_id = ?",
            finished_at,
        )

    def mark_failed(self, job_id: str, message: str) -> None:
        self.record_error(job_id, "", message, increment_processed=False)
        finished_at = datetime.now(UTC).isoformat()
        self._update(
            job_id,
            "UPDATE jobs SET status = 'failed', finished_at = ? WHERE job_id = ?",
            finished_at,
        )

    def record_success(self, job_id: str) -> None:
        self._update(
            job_id,
            "UPDATE jobs SET processed = processed + 1, success = success + 1 WHERE job_id = ?",
        )

    def record_error(self, job_id: str, cnpj: str, message: str, increment_processed: bool = True) -> None:
        job = self.get_job(job_id)
        errors = [*job.errors, {"cnpj": cnpj, "message": message}]
        processed_sql = "processed = processed + 1," if increment_processed else ""
        with self._connect() as connection:
            connection.execute(
                f"UPDATE jobs SET {processed_sql} errors_json = ? WHERE job_id = ?",
                (json.dumps(errors), job_id),
            )

    def _update(self, job_id: str, sql: str, *params: Any) -> None:
        with self._connect() as connection:
            result = connection.execute(sql, (*params, job_id))
            if result.rowcount == 0:
                raise JobNotFoundError(job_id)

    def cleanup_expired(self, retention_hours: float) -> None:
        cutoff = datetime.now(UTC) - timedelta(hours=retention_hours)
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT job_id, created_at, finished_at FROM jobs"
            ).fetchall()
            for row in rows:
                timestamp = row["finished_at"] or row["created_at"]
                if datetime.fromisoformat(timestamp) >= cutoff:
                    continue
                shutil.rmtree(self._job_dir(row["job_id"]), ignore_errors=True)
                connection.execute("DELETE FROM jobs WHERE job_id = ?", (row["job_id"],))
```

- [ ] **Step 4: Ignore job storage**

Add to `.gitignore`:

```text
storage/
```

- [ ] **Step 5: Run job tests**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_jobs -v
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add .gitignore src/api_planilhas/jobs.py tests/test_jobs.py
git commit -m "Adiciona store local de jobs"
```

---

### Task 4: Processador de lote

**Files:**
- Create: `src/api_planilhas/batch_processor.py`
- Create: `tests/test_batch_processor.py`

- [ ] **Step 1: Write failing processor test**

Create `tests/test_batch_processor.py`:

```python
import sys
import unittest
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from api_planilhas.config import AppSettings
from api_planilhas.directd import DirectDError
from api_planilhas.jobs import JobStore
from api_planilhas.batch_processor import process_job


class BatchProcessorTest(unittest.TestCase):
    def test_processes_successes_and_errors_without_stopping(self):
        with TemporaryDirectory() as temp_dir:
            store = JobStore(Path(temp_dir))
            store.initialize()
            job = store.create_job(["12.345.678/0001-90", "123", "11.222.333/0001-44"])
            settings = AppSettings(
                directd_token="token",
                basic_user="admin",
                basic_password="secret",
                directd_base_url="https://example.test",
            )

            def fake_fetch(cnpj, _settings):
                if cnpj == "11222333000144":
                    raise DirectDError("Falha DirectD")
                return {
                    "retorno": {
                        "cnpj": cnpj,
                        "razaoSocial": f"EMPRESA {cnpj}",
                        "telefones": [],
                        "emails": [],
                        "enderecos": [],
                        "socios": [],
                    }
                }

            process_job(job.job_id, store, settings, fetcher=fake_fetch, sleeper=lambda _seconds: None)

            loaded = store.get_job(job.job_id)
            self.assertEqual(loaded.status, "completed")
            self.assertEqual(loaded.processed, 3)
            self.assertEqual(loaded.success, 1)
            self.assertEqual(len(loaded.errors), 2)
            self.assertEqual(loaded.errors[0]["cnpj"], "123")
            self.assertEqual(loaded.errors[1]["cnpj"], "11222333000144")
            self.assertTrue(store.output_path(job.job_id).exists())
            with zipfile.ZipFile(store.output_path(job.job_id)) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
            self.assertIn("EMPRESA 12345678000190", sheet_xml)
            self.assertNotIn("11222333000144", sheet_xml)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_batch_processor -v
```

Expected: FAIL because `api_planilhas.batch_processor` does not exist.

- [ ] **Step 3: Implement processor**

Create `src/api_planilhas/batch_processor.py`:

```python
from __future__ import annotations

import time
from collections.abc import Callable

from api_planilhas.config import AppSettings
from api_planilhas.converter import extract_row, write_xlsx
from api_planilhas.directd import DirectDError, fetch_cnpj
from api_planilhas.jobs import JobStore
from api_planilhas.cnpj import validate_cnpj


FetchCnpj = Callable[[str, AppSettings], dict]
Sleeper = Callable[[float], None]


def process_job(
    job_id: str,
    store: JobStore,
    settings: AppSettings,
    fetcher: FetchCnpj = fetch_cnpj,
    sleeper: Sleeper = time.sleep,
) -> None:
    rows = []
    store.mark_processing(job_id)
    try:
        for raw_cnpj in store.get_input_cnpjs(job_id):
            try:
                cnpj = validate_cnpj(raw_cnpj)
                payload = fetcher(cnpj, settings)
                rows.append(extract_row(payload))
                store.record_success(job_id)
            except (ValueError, DirectDError) as exc:
                normalized_or_raw = "".join(ch for ch in raw_cnpj if ch.isdigit()) or raw_cnpj
                store.record_error(job_id, normalized_or_raw, str(exc))
            if settings.directd_batch_delay_seconds > 0:
                sleeper(settings.directd_batch_delay_seconds)
        write_xlsx(store.output_path(job_id), rows)
        store.mark_completed(job_id)
    except Exception as exc:
        store.mark_failed(job_id, str(exc))
```

- [ ] **Step 4: Run processor tests**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_batch_processor -v
```

Expected: PASS.

- [ ] **Step 5: Run previous tests**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_converter tests.test_web tests.test_xlsx_reader tests.test_jobs tests.test_batch_processor -v
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add src/api_planilhas/batch_processor.py tests/test_batch_processor.py
git commit -m "Adiciona processamento de lote CNPJ"
```

---

### Task 5: Endpoints de lote

**Files:**
- Modify: `src/api_planilhas/web.py`
- Modify: `tests/test_web.py`

- [ ] **Step 1: Write failing endpoint tests**

Add this test class in `tests/test_web.py`:

```python
class BatchRoutesTest(unittest.TestCase):
    def test_batch_upload_requires_basic_auth(self):
        env = {
            "DIRECTD_TOKEN": "token",
            "APP_BASIC_USER": "admin",
            "APP_BASIC_PASSWORD": "secret",
        }
        with patch.dict(os.environ, env, clear=True):
            response = TestClient(create_app()).post("/api/lotes")

        self.assertEqual(response.status_code, 401)

    def test_batch_upload_rejects_missing_cnpj_column(self):
        from tests.test_xlsx_reader import build_minimal_xlsx

        env = {
            "DIRECTD_TOKEN": "token",
            "APP_BASIC_USER": "admin",
            "APP_BASIC_PASSWORD": "secret",
        }
        content = build_minimal_xlsx([["Nome"], ["Empresa"]])

        with patch.dict(os.environ, env, clear=True):
            response = TestClient(create_app()).post(
                "/api/lotes",
                files={"file": ("entrada.xlsx", content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                headers=_basic_header(),
            )

        self.assertEqual(response.status_code, 400)
        self.assertIn(
            "nao tem cnpj na sua planilha adicione uma coluna CNPJ e os cnpj em baixo",
            response.text,
        )
```

- [ ] **Step 2: Run endpoint tests to verify failure**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_web.BatchRoutesTest -v
```

Expected: FAIL because `/api/lotes` does not exist.

- [ ] **Step 3: Implement app-level store and endpoints**

In `src/api_planilhas/web.py`, add imports:

```python
from fastapi import BackgroundTasks, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from api_planilhas.batch_processor import process_job
from api_planilhas.jobs import JobNotFoundError, JobStore
from api_planilhas.xlsx_reader import (
    InvalidXlsxError,
    MissingCnpjColumnError,
    extract_cnpj_values,
)
```

Add this helper above `create_app()`:

```python
def _get_job_store(app: FastAPI, settings: AppSettings) -> JobStore:
    store = getattr(app.state, "job_store", None)
    if store is not None and store.storage_dir == settings.job_storage_dir:
        return store
    store = JobStore(settings.job_storage_dir)
    store.initialize()
    store.cleanup_expired(settings.job_retention_hours)
    app.state.job_store = store
    return store
```

Add route:

```python
    @app.post("/api/lotes", dependencies=[Depends(require_basic_auth)])
    async def criar_lote(
        request: Request,
        background_tasks: BackgroundTasks,
        file: UploadFile,
        settings: AppSettings = Depends(get_settings),
    ) -> JSONResponse:
        content = await file.read()
        max_bytes = settings.upload_max_mb * 1024 * 1024
        if len(content) > max_bytes:
            raise HTTPException(status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE, detail="Arquivo acima do limite permitido")
        try:
            cnpjs = extract_cnpj_values(content)
        except MissingCnpjColumnError as exc:
            raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
        except InvalidXlsxError as exc:
            raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc

        store = _get_job_store(request.app, settings)
        job = store.create_job(cnpjs)
        background_tasks.add_task(process_job, job.job_id, store, settings)
        return JSONResponse(status_code=status.HTTP_202_ACCEPTED, content={"job_id": job.job_id})
```

Add status/download routes:

```python
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
            raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Job nao encontrado") from exc
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
            raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Job nao encontrado") from exc
        if not job.download_ready:
            raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="Planilha ainda nao esta pronta")
        return FileResponse(
            store.output_path(job_id),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=f"importacao_oportunidades_lote_{job_id}.xlsx",
        )
```

- [ ] **Step 4: Add completed job route test with mocked processor**

Add:

```python
    def test_batch_upload_returns_job_id(self):
        from tests.test_xlsx_reader import build_minimal_xlsx

        env = {
            "DIRECTD_TOKEN": "token",
            "APP_BASIC_USER": "admin",
            "APP_BASIC_PASSWORD": "secret",
        }
        content = build_minimal_xlsx([["CNPJ"], ["12.345.678/0001-90"]])

        with patch.dict(os.environ, env, clear=True):
            with patch("api_planilhas.web.process_job"):
                response = TestClient(create_app()).post(
                    "/api/lotes",
                    files={"file": ("entrada.xlsx", content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                    headers=_basic_header(),
                )

        self.assertEqual(response.status_code, 202)
        self.assertIn("job_id", response.json())
```

- [ ] **Step 5: Run route tests**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_web.BatchRoutesTest -v
```

Expected: PASS.

- [ ] **Step 6: Run full backend tests**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_web tests.test_converter tests.test_xlsx_reader tests.test_jobs tests.test_batch_processor -v
```

Expected: PASS.

- [ ] **Step 7: Commit**

```powershell
git add src/api_planilhas/web.py tests/test_web.py
git commit -m "Adiciona endpoints de lote CNPJ"
```

---

### Task 6: Frontend de upload/status/download

**Files:**
- Modify: `src/api_planilhas/templates/index.html`
- Modify: `tests/test_web.py`

- [ ] **Step 1: Write failing frontend assertions**

Update `HtmlFrontendTest.test_homepage_has_frontend_controls`:

```python
        self.assertIn('type="file"', response.text)
        self.assertIn("/api/lotes", response.text)
        self.assertIn("Consultar status", response.text)
        self.assertIn("CNPJs com erro", response.text)
        self.assertIn("downloadBatch", response.text)
```

- [ ] **Step 2: Run frontend test to verify failure**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_web.HtmlFrontendTest -v
```

Expected: FAIL because the batch UI does not exist.

- [ ] **Step 3: Add batch UI markup**

In `src/api_planilhas/templates/index.html`, after the current download button:

```html
      <hr />
      <h2>Gerar planilha em lote</h2>
      <form id="batchForm">
        <label for="batchFile">Planilha XLSX com coluna CNPJ</label>
        <input id="batchFile" name="file" type="file" accept=".xlsx" required />
        <button type="submit" id="batchButton">Enviar lote</button>
      </form>
      <p id="batchStatus" class="status" aria-live="polite"></p>
      <button id="batchDownloadButton" type="button" class="hidden">
        Baixar planilha do lote
      </button>
      <section id="batchErrorsSection" class="hidden">
        <h3>CNPJs com erro</h3>
        <ul id="batchErrors"></ul>
      </section>
      <button id="manualStatusButton" type="button" class="hidden">
        Consultar status
      </button>
```

- [ ] **Step 4: Add batch JavaScript**

In the script, add constants/state:

```javascript
      const batchForm = document.getElementById("batchForm");
      const batchFileInput = document.getElementById("batchFile");
      const batchButton = document.getElementById("batchButton");
      const batchStatus = document.getElementById("batchStatus");
      const batchDownloadButton = document.getElementById("batchDownloadButton");
      const batchErrorsSection = document.getElementById("batchErrorsSection");
      const batchErrors = document.getElementById("batchErrors");
      const manualStatusButton = document.getElementById("manualStatusButton");
      let currentBatchJobId = "";
      let batchPollTimer = 0;
```

Add functions:

```javascript
      function setBatchStatus(message, type = "") {
        batchStatus.textContent = message || "";
        batchStatus.className = `status ${type}`;
      }

      function renderBatchErrors(errors) {
        batchErrors.textContent = "";
        batchErrorsSection.classList.toggle("hidden", !errors || errors.length === 0);
        for (const item of errors || []) {
          const li = document.createElement("li");
          li.textContent = `${item.cnpj || "sem cnpj"} - ${item.message}`;
          batchErrors.appendChild(li);
        }
      }

      async function refreshBatchStatus() {
        if (!currentBatchJobId) {
          return;
        }
        const response = await fetch(`/api/lotes/${currentBatchJobId}`);
        if (!response.ok) {
          setBatchStatus("Nao foi possivel consultar o lote.", "error");
          return;
        }
        const payload = await response.json();
        renderBatchErrors(payload.errors);
        setBatchStatus(
          `Status: ${payload.status}. Processados: ${payload.processed}/${payload.total}. Sucessos: ${payload.success}.`,
          payload.status === "completed" ? "success" : ""
        );
        batchDownloadButton.classList.toggle("hidden", !payload.download_ready);
        manualStatusButton.classList.remove("hidden");
        if (payload.status === "completed" || payload.status === "failed") {
          clearInterval(batchPollTimer);
          batchPollTimer = 0;
        }
      }

      function downloadBatch() {
        if (!currentBatchJobId) {
          return;
        }
        window.location.href = `/api/lotes/${currentBatchJobId}/download`;
      }

      async function submitBatch(event) {
        event.preventDefault();
        const file = batchFileInput.files[0];
        if (!file) {
          setBatchStatus("Selecione uma planilha XLSX.", "error");
          return;
        }
        const formData = new FormData();
        formData.append("file", file);
        batchButton.disabled = true;
        setBatchStatus("Enviando lote...");
        renderBatchErrors([]);
        try {
          const response = await fetch("/api/lotes", { method: "POST", body: formData });
          if (!response.ok) {
            throw new Error(await readErrorMessage(response));
          }
          const payload = await response.json();
          currentBatchJobId = payload.job_id;
          setBatchStatus("Lote recebido. Processando...");
          await refreshBatchStatus();
          clearInterval(batchPollTimer);
          batchPollTimer = setInterval(refreshBatchStatus, 3000);
        } catch (error) {
          setBatchStatus(error.message || "Erro ao enviar lote.", "error");
        } finally {
          batchButton.disabled = false;
        }
      }

      batchForm.addEventListener("submit", submitBatch);
      batchDownloadButton.addEventListener("click", downloadBatch);
      manualStatusButton.addEventListener("click", refreshBatchStatus);
```

- [ ] **Step 5: Run frontend tests**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_web.HtmlFrontendTest -v
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add src/api_planilhas/templates/index.html tests/test_web.py
git commit -m "Adiciona interface de lote CNPJ"
```

---

### Task 7: Retencao automatica e verificacao final

**Files:**
- Modify: `tests/test_jobs.py`
- Modify: `src/api_planilhas/jobs.py` if needed

- [ ] **Step 1: Add cleanup regression test**

Add in `tests/test_jobs.py`:

```python
    def test_cleanup_expired_removes_old_jobs(self):
        from datetime import UTC, datetime, timedelta

        with TemporaryDirectory() as temp_dir:
            store = JobStore(Path(temp_dir))
            store.initialize()
            job = store.create_job(["12345678000190"])
            old_timestamp = (datetime.now(UTC) - timedelta(hours=72)).isoformat()
            with store._connect() as connection:
                connection.execute(
                    "UPDATE jobs SET status = 'completed', finished_at = ? WHERE job_id = ?",
                    (old_timestamp, job.job_id),
                )

            store.cleanup_expired(retention_hours=48)

            with self.assertRaises(KeyError):
                store.get_job(job.job_id)
            self.assertFalse((Path(temp_dir) / job.job_id).exists())
```

- [ ] **Step 2: Run cleanup test**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_jobs.JobStoreTest.test_cleanup_expired_removes_old_jobs -v
```

Expected: PASS or fix `cleanup_expired` until it passes.

- [ ] **Step 3: Run full test suite**

Run:

```powershell
.\.venv\Scripts\python.exe -m unittest tests.test_web tests.test_converter tests.test_xlsx_reader tests.test_jobs tests.test_batch_processor -v
```

Expected: all tests PASS.

- [ ] **Step 4: Start local server smoke test**

Run:

```powershell
$env:DIRECTD_TOKEN = 'dev-token'; $env:APP_BASIC_USER = 'admin'; $env:APP_BASIC_PASSWORD = 'admin'; $env:JOB_STORAGE_DIR = 'storage/jobs'; $p = Start-Process -FilePath '.\.venv\Scripts\python.exe' -ArgumentList @('-m','uvicorn','api_planilhas.web:app','--app-dir','src','--host','127.0.0.1','--port','8000') -WindowStyle Hidden -PassThru; Start-Sleep -Seconds 3; try { $bytes = [System.Text.Encoding]::ASCII.GetBytes('admin:admin'); $auth = 'Basic ' + [Convert]::ToBase64String($bytes); $response = Invoke-WebRequest -Uri 'http://127.0.0.1:8000/' -Headers @{ Authorization = $auth } -UseBasicParsing; "StatusCode=$($response.StatusCode)"; "HasBatch=$($response.Content.Contains('/api/lotes'))" } finally { if ($p -and -not $p.HasExited) { Stop-Process -Id $p.Id -Force } }
```

Expected:

```text
StatusCode=200
HasBatch=True
```

- [ ] **Step 5: Commit final cleanup/test fixes**

If Task 7 changed files:

```powershell
git add src/api_planilhas/jobs.py tests/test_jobs.py
git commit -m "Finaliza limpeza automatica de jobs"
```

If no files changed, skip commit.

- [ ] **Step 6: Final Git status**

Run:

```powershell
git status --short --branch
```

Expected: clean worktree except ignored local `.env`, `.venv`, and runtime `storage/`.

---

## Self-Review

- Spec coverage: upload XLSX, coluna `CNPJ`, erro obrigatorio, fila local, SQLite, worker em background, processamento sequencial, status, download, erros na tela, retencao automatica e configuracoes estao cobertos nas Tasks 1-7.
- Placeholder scan: nenhuma pendencia textual ou etapa de codigo omitida foi encontrada.
- Type consistency: `AppSettings`, `JobStore`, `JobSnapshot`, `process_job`, `extract_cnpj_values`, and route names are defined before use.
