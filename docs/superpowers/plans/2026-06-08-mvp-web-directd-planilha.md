# MVP Web DirectD Planilha Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a FastAPI web MVP that protects the UI/API with Basic Auth, accepts one CNPJ, calls DirectD, generates an XLSX in memory, and returns it for automatic browser download.

**Architecture:** Keep the existing converter as the spreadsheet core and add a thin web layer around it. Split responsibilities into focused modules: settings, security, DirectD client, web routes, and template.

**Tech Stack:** Python 3, FastAPI, Uvicorn, standard library `urllib`, existing XLSX writer, `unittest` with FastAPI `TestClient`.

---

## File Structure

- Create `requirements.txt`: runtime/test dependencies for FastAPI app.
- Modify `src/api_planilhas/converter.py`: add `build_xlsx_bytes(rows)` so web routes can return XLSX without permanent files.
- Create `src/api_planilhas/config.py`: read and validate environment configuration.
- Create `src/api_planilhas/security.py`: Basic Auth dependency using constant-time comparison.
- Create `src/api_planilhas/directd.py`: DirectD client using `urllib.request` with timeout and explicit errors.
- Create `src/api_planilhas/web.py`: FastAPI app, CNPJ normalization/validation, HTML route and XLSX download route.
- Create `src/api_planilhas/templates/index.html`: single-page HTML/CSS/JS form with automatic download and retry button.
- Create `tests/test_web.py`: tests for settings, CNPJ validation, auth, route behavior and HTML content.

No commits are included in the steps because the project-level instructions prohibit committing without explicit user approval.

---

### Task 1: Dependencies And In-Memory XLSX

**Files:**
- Create: `requirements.txt`
- Modify: `src/api_planilhas/converter.py`
- Modify: `tests/test_converter.py`

- [ ] **Step 1: Write failing test for in-memory XLSX bytes**

Append this test method to `WriteXlsxTest` in `tests/test_converter.py`:

```python
    def test_build_xlsx_bytes_returns_valid_package(self):
        import zipfile
        from io import BytesIO

        from api_planilhas.converter import build_xlsx_bytes

        content = build_xlsx_bytes([["EMPRESA BYTES LTDA"] + [""] * (len(HEADERS) - 1)])

        self.assertIsInstance(content, bytes)
        with zipfile.ZipFile(BytesIO(content)) as archive:
            self.assertIn("xl/worksheets/sheet1.xml", archive.namelist())
            sheet = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertIn("Pessoa", sheet)
        self.assertIn("EMPRESA BYTES LTDA", sheet)
```

- [ ] **Step 2: Run the failing test**

Run:

```powershell
python -m unittest tests.test_converter.WriteXlsxTest.test_build_xlsx_bytes_returns_valid_package -v
```

Expected: FAIL with `ImportError` or `AttributeError` for missing `build_xlsx_bytes`.

- [ ] **Step 3: Add dependencies file**

Create `requirements.txt`:

```text
fastapi==0.115.6
uvicorn[standard]==0.34.0
httpx==0.28.1
```

- [ ] **Step 4: Refactor XLSX writer to support bytes**

Modify `src/api_planilhas/converter.py` so `write_xlsx` delegates to a new bytes builder:

```python
from io import BytesIO
```

Add this helper near `write_xlsx`:

```python
def _write_xlsx_archive(target: str | Path | BytesIO, rows: list[list[Any]]) -> None:
    worksheet_rows = [HEADERS, *rows]
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types_xml())
        archive.writestr("_rels/.rels", _package_rels_xml())
        archive.writestr("xl/workbook.xml", _workbook_xml())
        archive.writestr("xl/_rels/workbook.xml.rels", _workbook_rels_xml())
        archive.writestr("xl/worksheets/sheet1.xml", _sheet_xml(worksheet_rows))
```

Replace the body of `write_xlsx` with:

```python
def write_xlsx(output_path: str | Path, rows: list[list[Any]]) -> None:
    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    _write_xlsx_archive(output_file, rows)
```

Add:

```python
def build_xlsx_bytes(rows: list[list[Any]]) -> bytes:
    buffer = BytesIO()
    _write_xlsx_archive(buffer, rows)
    return buffer.getvalue()
```

- [ ] **Step 5: Run converter tests**

Run:

```powershell
python -m unittest tests.test_converter -v
```

Expected: OK.

---

### Task 2: Configuration, CNPJ Normalization And Basic Auth

**Files:**
- Create: `src/api_planilhas/config.py`
- Create: `src/api_planilhas/security.py`
- Create: `src/api_planilhas/web.py`
- Create: `tests/test_web.py`

- [ ] **Step 1: Write failing tests for settings, CNPJ and auth**

Create `tests/test_web.py`:

```python
import base64
import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))


class ConfigAndSecurityTest(unittest.TestCase):
    def test_normalize_cnpj_keeps_only_digits(self):
        from api_planilhas.web import normalize_cnpj

        self.assertEqual(normalize_cnpj("12.345.678/0001-90"), "12345678000190")

    def test_validate_cnpj_rejects_non_14_digits(self):
        from api_planilhas.web import validate_cnpj

        with self.assertRaises(ValueError):
            validate_cnpj("123")

    def test_settings_reads_environment(self):
        from api_planilhas.config import get_settings

        env = {
            "DIRECTD_TOKEN": "token",
            "APP_BASIC_USER": "admin",
            "APP_BASIC_PASSWORD": "secret",
            "DIRECTD_BASE_URL": "https://example.test",
        }
        with patch.dict(os.environ, env, clear=True):
            settings = get_settings()

        self.assertEqual(settings.directd_token, "token")
        self.assertEqual(settings.basic_user, "admin")
        self.assertEqual(settings.basic_password, "secret")
        self.assertEqual(settings.directd_base_url, "https://example.test")

    def test_settings_requires_directd_token(self):
        from api_planilhas.config import get_settings

        env = {"APP_BASIC_USER": "admin", "APP_BASIC_PASSWORD": "secret"}
        with patch.dict(os.environ, env, clear=True):
            with self.assertRaises(RuntimeError):
                get_settings()

    def test_basic_auth_accepts_valid_credentials(self):
        from fastapi.testclient import TestClient
        from api_planilhas.web import create_app

        env = {
            "DIRECTD_TOKEN": "token",
            "APP_BASIC_USER": "admin",
            "APP_BASIC_PASSWORD": "secret",
        }
        credentials = base64.b64encode(b"admin:secret").decode("ascii")
        with patch.dict(os.environ, env, clear=True):
            client = TestClient(create_app())
            response = client.get("/", headers={"Authorization": f"Basic {credentials}"})

        self.assertEqual(response.status_code, 200)

    def test_basic_auth_rejects_invalid_credentials(self):
        from fastapi.testclient import TestClient
        from api_planilhas.web import create_app

        env = {
            "DIRECTD_TOKEN": "token",
            "APP_BASIC_USER": "admin",
            "APP_BASIC_PASSWORD": "secret",
        }
        credentials = base64.b64encode(b"admin:wrong").decode("ascii")
        with patch.dict(os.environ, env, clear=True):
            client = TestClient(create_app())
            response = client.get("/", headers={"Authorization": f"Basic {credentials}"})

        self.assertEqual(response.status_code, 401)
```

- [ ] **Step 2: Run failing tests**

Run:

```powershell
python -m unittest tests.test_web.ConfigAndSecurityTest -v
```

Expected: FAIL because modules are missing.

- [ ] **Step 3: Implement settings**

Create `src/api_planilhas/config.py`:

```python
from __future__ import annotations

import os
from dataclasses import dataclass


DEFAULT_DIRECTD_BASE_URL = "https://apiv3.directd.com.br"


@dataclass(frozen=True)
class AppSettings:
    directd_token: str
    basic_user: str
    basic_password: str
    directd_base_url: str = DEFAULT_DIRECTD_BASE_URL
    directd_timeout_seconds: float = 20.0


def _required(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Variavel obrigatoria ausente: {name}")
    return value


def get_settings() -> AppSettings:
    return AppSettings(
        directd_token=_required("DIRECTD_TOKEN"),
        basic_user=_required("APP_BASIC_USER"),
        basic_password=_required("APP_BASIC_PASSWORD"),
        directd_base_url=os.environ.get("DIRECTD_BASE_URL", DEFAULT_DIRECTD_BASE_URL).rstrip("/"),
    )
```

- [ ] **Step 4: Implement Basic Auth dependency**

Create `src/api_planilhas/security.py`:

```python
from __future__ import annotations

import secrets
from typing import Annotated

from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPBasic, HTTPBasicCredentials

from api_planilhas.config import AppSettings, get_settings


security = HTTPBasic()


def require_basic_auth(
    credentials: Annotated[HTTPBasicCredentials, Depends(security)],
    settings: Annotated[AppSettings, Depends(get_settings)],
) -> str:
    user_ok = secrets.compare_digest(credentials.username, settings.basic_user)
    password_ok = secrets.compare_digest(credentials.password, settings.basic_password)
    if not (user_ok and password_ok):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Credenciais invalidas",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username
```

- [ ] **Step 5: Implement initial FastAPI app and CNPJ helpers**

Create `src/api_planilhas/web.py`:

```python
from __future__ import annotations

import re
from typing import Annotated

from fastapi import Depends, FastAPI
from fastapi.responses import HTMLResponse

from api_planilhas.security import require_basic_auth


def normalize_cnpj(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def validate_cnpj(value: str) -> str:
    cnpj = normalize_cnpj(value)
    if len(cnpj) != 14:
        raise ValueError("CNPJ deve conter 14 digitos")
    return cnpj


def create_app() -> FastAPI:
    app = FastAPI(title="API Planilhas")

    @app.get("/", response_class=HTMLResponse)
    def index(_: Annotated[str, Depends(require_basic_auth)]) -> str:
        return "<html><body><form><input name=\"cnpj\"><button>Gerar planilha</button></form></body></html>"

    return app


app = create_app()
```

- [ ] **Step 6: Run config/security tests**

Run:

```powershell
python -m unittest tests.test_web.ConfigAndSecurityTest -v
```

Expected: OK.

---

### Task 3: DirectD Client

**Files:**
- Create: `src/api_planilhas/directd.py`
- Modify: `tests/test_web.py`

- [ ] **Step 1: Add failing DirectD client tests**

Append to `tests/test_web.py`:

```python
class DirectDClientTest(unittest.TestCase):
    def test_builds_cnpj_url_without_logging_token(self):
        from api_planilhas.config import AppSettings
        from api_planilhas.directd import build_cnpj_url

        settings = AppSettings(
            directd_token="secret-token",
            basic_user="admin",
            basic_password="secret",
            directd_base_url="https://example.test",
        )

        url = build_cnpj_url("12345678000190", settings)

        self.assertEqual(
            url,
            "https://example.test/api/CadastroPessoaJuridica?CNPJ=12345678000190&Token=secret-token",
        )

    def test_fetch_cnpj_returns_json_payload(self):
        import json
        from unittest.mock import patch

        from api_planilhas.config import AppSettings
        from api_planilhas.directd import fetch_cnpj

        class FakeResponse:
            status = 200

            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb):
                return False

            def read(self):
                return json.dumps({"retorno": {"cnpj": "12345678000190"}}).encode("utf-8")

        settings = AppSettings("token", "admin", "secret", "https://example.test")
        with patch("api_planilhas.directd.urlopen", return_value=FakeResponse()):
            payload = fetch_cnpj("12345678000190", settings)

        self.assertEqual(payload["retorno"]["cnpj"], "12345678000190")
```

- [ ] **Step 2: Run failing tests**

Run:

```powershell
python -m unittest tests.test_web.DirectDClientTest -v
```

Expected: FAIL because `api_planilhas.directd` is missing.

- [ ] **Step 3: Implement DirectD client**

Create `src/api_planilhas/directd.py`:

```python
from __future__ import annotations

import json
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from api_planilhas.config import AppSettings


class DirectDError(RuntimeError):
    pass


def build_cnpj_url(cnpj: str, settings: AppSettings) -> str:
    query = urlencode({"CNPJ": cnpj, "Token": settings.directd_token})
    return f"{settings.directd_base_url}/api/CadastroPessoaJuridica?{query}"


def fetch_cnpj(cnpj: str, settings: AppSettings) -> dict[str, Any]:
    request = Request(build_cnpj_url(cnpj, settings), headers={"Accept": "application/json"})
    try:
        with urlopen(request, timeout=settings.directd_timeout_seconds) as response:
            raw = response.read()
    except HTTPError as exc:
        raise DirectDError(f"DirectD retornou HTTP {exc.code}") from exc
    except URLError as exc:
        raise DirectDError("Falha ao consultar DirectD") from exc
    except TimeoutError as exc:
        raise DirectDError("Tempo limite ao consultar DirectD") from exc

    try:
        payload = json.loads(raw.decode("utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise DirectDError("DirectD retornou JSON invalido") from exc

    if not isinstance(payload, dict):
        raise DirectDError("DirectD retornou resposta invalida")
    if not isinstance(payload.get("retorno"), dict):
        raise DirectDError("DirectD nao retornou dados uteis para o CNPJ")
    return payload
```

- [ ] **Step 4: Run DirectD tests**

Run:

```powershell
python -m unittest tests.test_web.DirectDClientTest -v
```

Expected: OK.

---

### Task 4: XLSX API Route

**Files:**
- Modify: `src/api_planilhas/web.py`
- Modify: `tests/test_web.py`

- [ ] **Step 1: Add failing route tests**

Append to `tests/test_web.py`:

```python
def _basic_header(user: str = "admin", password: str = "secret") -> dict[str, str]:
    token = base64.b64encode(f"{user}:{password}".encode("utf-8")).decode("ascii")
    return {"Authorization": f"Basic {token}"}


class PlanilhaRouteTest(unittest.TestCase):
    def test_planilha_route_returns_xlsx(self):
        from fastapi.testclient import TestClient
        from api_planilhas.web import create_app

        env = {
            "DIRECTD_TOKEN": "token",
            "APP_BASIC_USER": "admin",
            "APP_BASIC_PASSWORD": "secret",
        }
        payload = {
            "retorno": {
                "cnpj": "12345678000190",
                "razaoSocial": "EMPRESA API LTDA",
                "telefones": [],
                "emails": [],
                "enderecos": [],
                "socios": [],
            }
        }
        with patch.dict(os.environ, env, clear=True):
            with patch("api_planilhas.web.fetch_cnpj", return_value=payload):
                client = TestClient(create_app())
                response = client.post(
                    "/api/planilha",
                    data={"cnpj": "12.345.678/0001-90"},
                    headers=_basic_header(),
                )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.headers["content-type"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.assertIn("attachment;", response.headers["content-disposition"])
        self.assertGreater(len(response.content), 100)

    def test_planilha_route_rejects_invalid_cnpj(self):
        from fastapi.testclient import TestClient
        from api_planilhas.web import create_app

        env = {
            "DIRECTD_TOKEN": "token",
            "APP_BASIC_USER": "admin",
            "APP_BASIC_PASSWORD": "secret",
        }
        with patch.dict(os.environ, env, clear=True):
            client = TestClient(create_app())
            response = client.post(
                "/api/planilha",
                data={"cnpj": "123"},
                headers=_basic_header(),
            )

        self.assertEqual(response.status_code, 400)

    def test_planilha_route_maps_directd_error(self):
        from fastapi.testclient import TestClient
        from api_planilhas.directd import DirectDError
        from api_planilhas.web import create_app

        env = {
            "DIRECTD_TOKEN": "token",
            "APP_BASIC_USER": "admin",
            "APP_BASIC_PASSWORD": "secret",
        }
        with patch.dict(os.environ, env, clear=True):
            with patch("api_planilhas.web.fetch_cnpj", side_effect=DirectDError("Falha DirectD")):
                client = TestClient(create_app())
                response = client.post(
                    "/api/planilha",
                    data={"cnpj": "12345678000190"},
                    headers=_basic_header(),
                )

        self.assertEqual(response.status_code, 502)
        self.assertIn("Falha DirectD", response.text)
```

- [ ] **Step 2: Run failing route tests**

Run:

```powershell
python -m unittest tests.test_web.PlanilhaRouteTest -v
```

Expected: FAIL because `/api/planilha` is missing.

- [ ] **Step 3: Implement route**

Modify `src/api_planilhas/web.py`:

```python
from fastapi import Depends, FastAPI, Form, HTTPException, status
from fastapi.responses import HTMLResponse, Response

from api_planilhas.config import AppSettings, get_settings
from api_planilhas.converter import build_xlsx_bytes, extract_row
from api_planilhas.directd import DirectDError, fetch_cnpj
from api_planilhas.security import require_basic_auth
```

Inside `create_app`, after `index`:

```python
    @app.post("/api/planilha")
    def gerar_planilha(
        _: Annotated[str, Depends(require_basic_auth)],
        settings: Annotated[AppSettings, Depends(get_settings)],
        cnpj: str = Form(...),
    ) -> Response:
        try:
            normalized = validate_cnpj(cnpj)
        except ValueError as exc:
            raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc

        try:
            payload = fetch_cnpj(normalized, settings)
        except DirectDError as exc:
            raise HTTPException(status_code=status.HTTP_502_BAD_GATEWAY, detail=str(exc)) from exc

        content = build_xlsx_bytes([extract_row(payload)])
        filename = f"importacao_oportunidade_{normalized}.xlsx"
        return Response(
            content=content,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
```

- [ ] **Step 4: Run route tests**

Run:

```powershell
python -m unittest tests.test_web.PlanilhaRouteTest -v
```

Expected: OK.

---

### Task 5: HTML Frontend

**Files:**
- Create: `src/api_planilhas/templates/index.html`
- Modify: `src/api_planilhas/web.py`
- Modify: `tests/test_web.py`

- [ ] **Step 1: Add failing HTML tests**

Append to `tests/test_web.py`:

```python
class HtmlFrontendTest(unittest.TestCase):
    def test_index_contains_form_and_download_retry_button(self):
        from fastapi.testclient import TestClient
        from api_planilhas.web import create_app

        env = {
            "DIRECTD_TOKEN": "token",
            "APP_BASIC_USER": "admin",
            "APP_BASIC_PASSWORD": "secret",
        }
        with patch.dict(os.environ, env, clear=True):
            client = TestClient(create_app())
            response = client.get("/", headers=_basic_header())

        self.assertEqual(response.status_code, 200)
        self.assertIn("Gerar planilha", response.text)
        self.assertIn("Baixar novamente", response.text)
        self.assertIn("/api/planilha", response.text)
        self.assertIn("downloadUrl", response.text)
```

- [ ] **Step 2: Run failing HTML test**

Run:

```powershell
python -m unittest tests.test_web.HtmlFrontendTest -v
```

Expected: FAIL because current HTML is minimal and lacks retry behavior.

- [ ] **Step 3: Create HTML template**

Create `src/api_planilhas/templates/index.html`:

```html
<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Gerar planilha CNPJ</title>
  <style>
    :root {
      font-family: Arial, sans-serif;
      color: #1f2933;
      background: #f4f7fb;
    }
    body {
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
    }
    main {
      width: min(520px, calc(100vw - 32px));
      background: #ffffff;
      border: 1px solid #d8e0ea;
      border-radius: 8px;
      padding: 24px;
      box-shadow: 0 8px 24px rgba(31, 41, 51, 0.08);
    }
    h1 {
      margin: 0 0 16px;
      font-size: 22px;
    }
    label {
      display: block;
      margin-bottom: 8px;
      font-weight: 700;
    }
    input {
      box-sizing: border-box;
      width: 100%;
      height: 44px;
      padding: 8px 10px;
      border: 1px solid #b8c4d3;
      border-radius: 6px;
      font-size: 16px;
    }
    .actions {
      display: flex;
      gap: 8px;
      margin-top: 16px;
      flex-wrap: wrap;
    }
    button {
      height: 40px;
      padding: 0 14px;
      border: 0;
      border-radius: 6px;
      background: #1769aa;
      color: #ffffff;
      font-weight: 700;
      cursor: pointer;
    }
    button[disabled] {
      opacity: 0.55;
      cursor: not-allowed;
    }
    #retry {
      background: #52606d;
    }
    #status {
      min-height: 24px;
      margin-top: 14px;
      font-size: 14px;
    }
    .error {
      color: #b42318;
    }
    .success {
      color: #1a7f37;
    }
  </style>
</head>
<body>
  <main>
    <h1>Gerar planilha CNPJ</h1>
    <form id="cnpj-form">
      <label for="cnpj">CNPJ</label>
      <input id="cnpj" name="cnpj" inputmode="numeric" autocomplete="off" required>
      <div class="actions">
        <button id="submit" type="submit">Gerar planilha</button>
        <button id="retry" type="button" hidden>Baixar novamente</button>
      </div>
      <div id="status" role="status" aria-live="polite"></div>
    </form>
  </main>
  <script>
    const form = document.querySelector("#cnpj-form");
    const input = document.querySelector("#cnpj");
    const submit = document.querySelector("#submit");
    const retry = document.querySelector("#retry");
    const statusBox = document.querySelector("#status");
    let downloadUrl = null;
    let downloadName = "importacao_oportunidade.xlsx";

    function setStatus(message, kind) {
      statusBox.textContent = message;
      statusBox.className = kind || "";
    }

    function triggerDownload() {
      if (!downloadUrl) return;
      const link = document.createElement("a");
      link.href = downloadUrl;
      link.download = downloadName;
      document.body.appendChild(link);
      link.click();
      link.remove();
    }

    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      if (downloadUrl) URL.revokeObjectURL(downloadUrl);
      downloadUrl = null;
      retry.hidden = true;
      submit.disabled = true;
      setStatus("Consultando dados...", "");

      const body = new URLSearchParams();
      body.set("cnpj", input.value);

      try {
        const response = await fetch("/api/planilha", {
          method: "POST",
          body,
        });
        if (!response.ok) {
          let message = "Nao foi possivel gerar a planilha.";
          try {
            const payload = await response.json();
            if (payload.detail) message = payload.detail;
          } catch (_) {
            message = await response.text();
          }
          throw new Error(message);
        }
        const blob = await response.blob();
        downloadUrl = URL.createObjectURL(blob);
        const disposition = response.headers.get("content-disposition") || "";
        const match = disposition.match(/filename="([^"]+)"/);
        downloadName = match ? match[1] : "importacao_oportunidade.xlsx";
        triggerDownload();
        retry.hidden = false;
        setStatus("Planilha gerada com sucesso.", "success");
      } catch (error) {
        setStatus(error.message || "Erro ao gerar planilha.", "error");
      } finally {
        submit.disabled = false;
      }
    });

    retry.addEventListener("click", triggerDownload);
  </script>
</body>
</html>
```

- [ ] **Step 4: Serve HTML file**

Modify `src/api_planilhas/web.py`:

```python
from pathlib import Path


TEMPLATE_PATH = Path(__file__).resolve().parent / "templates" / "index.html"
```

Replace `index` body with:

```python
        return TEMPLATE_PATH.read_text(encoding="utf-8")
```

- [ ] **Step 5: Run HTML tests and all web tests**

Run:

```powershell
python -m unittest tests.test_web -v
```

Expected: OK.

---

### Task 6: Local Run Verification

**Files:**
- No source changes expected unless verification exposes a bug.

- [ ] **Step 1: Install dependencies if missing**

Check:

```powershell
python -c "import fastapi; print(fastapi.__version__)"
```

If missing, ask the user for approval before installing dependencies. Suggested command:

```powershell
python -m pip install -r requirements.txt
```

- [ ] **Step 2: Run all automated tests**

Run:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Expected: OK.

- [ ] **Step 3: Start local dev server**

Run with temporary local values:

```powershell
$env:DIRECTD_TOKEN='dev-token'; $env:APP_BASIC_USER='admin'; $env:APP_BASIC_PASSWORD='admin'; python -m uvicorn api_planilhas.web:app --app-dir src --host 127.0.0.1 --port 8000
```

Expected: Uvicorn starts at `http://127.0.0.1:8000`.

- [ ] **Step 4: Verify server responds**

In another command while server is running:

```powershell
python -c "import base64, urllib.request; token=base64.b64encode(b'admin:admin').decode(); req=urllib.request.Request('http://127.0.0.1:8000/', headers={'Authorization':'Basic '+token}); print(urllib.request.urlopen(req, timeout=5).status)"
```

Expected:

```text
200
```

- [ ] **Step 5: Stop the server and check git status**

Stop Uvicorn cleanly, then run:

```powershell
git status --short
```

Expected: source/test/docs/new dependency files are visible. Do not commit unless the user explicitly asks.

---

## Self-Review

- Spec coverage: the plan covers FastAPI, Basic Auth, environment configuration, DirectD integration, one CNPJ per request, in-memory XLSX generation, automatic download with retry button, no persistent JSON/XLSX storage, and automated tests.
- Red-flag scan: no incomplete implementation markers remain; fields outside business rules remain explicitly out of scope.
- Type consistency: `AppSettings`, `DirectDError`, `fetch_cnpj`, `build_xlsx_bytes`, `normalize_cnpj`, `validate_cnpj`, and `create_app` are defined before use in later tasks.
