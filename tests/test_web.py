import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

from fastapi.testclient import TestClient

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from api_planilhas.config import (
    DEFAULT_DIRECTD_BASE_URL,
    DEFAULT_DIRECTD_TIMEOUT_SECONDS,
    DEFAULT_DIRECTD_BATCH_DELAY_SECONDS,
    DEFAULT_JOB_RETENTION_HOURS,
    DEFAULT_JOB_STORAGE_DIR,
    DEFAULT_UPLOAD_MAX_MB,
    get_settings,
)
from api_planilhas.web import create_app, validate_cnpj, normalize_cnpj


def _basic_header(user: str = "admin", password: str = "secret") -> dict[str, str]:
    import base64

    token = base64.b64encode(f"{user}:{password}".encode("utf-8")).decode("ascii")
    return {"Authorization": f"Basic {token}"}


class ConfigAndSecurityTest(unittest.TestCase):
    def setUp(self):
        self._original_env = os.environ.copy()

    def tearDown(self):
        os.environ.clear()
        os.environ.update(self._original_env)

    def test_normalize_cnpj_removes_all_non_digits(self):
        self.assertEqual(normalize_cnpj(" 12.345.678/0001-90 "), "12345678000190")
        self.assertEqual(normalize_cnpj("12.345.678/0001-10"), "12345678000110")
        self.assertEqual(normalize_cnpj(""), "")

    def test_validate_cnpj_accepts_14_digits(self):
        self.assertEqual(validate_cnpj("04.252.011/0001-10"), "04252011000110")

    def test_validate_cnpj_rejects_non_14_digits(self):
        with self.assertRaises(ValueError):
            validate_cnpj("123")

    def test_get_settings_reads_variables_and_defaults(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"

        settings = get_settings()

        self.assertEqual(settings.directd_token, "token-for-test")
        self.assertEqual(settings.basic_user, "admin")
        self.assertEqual(settings.basic_password, "secret")
        self.assertEqual(settings.directd_base_url, DEFAULT_DIRECTD_BASE_URL)
        self.assertEqual(settings.directd_timeout_seconds, DEFAULT_DIRECTD_TIMEOUT_SECONDS)

    def test_get_settings_reads_batch_defaults(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"

        settings = get_settings()

        self.assertEqual(settings.job_storage_dir, DEFAULT_JOB_STORAGE_DIR)
        self.assertEqual(settings.job_retention_hours, DEFAULT_JOB_RETENTION_HOURS)
        self.assertEqual(settings.upload_max_mb, DEFAULT_UPLOAD_MAX_MB)
        self.assertEqual(
            settings.directd_batch_delay_seconds,
            DEFAULT_DIRECTD_BATCH_DELAY_SECONDS,
        )

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

    def test_get_settings_reads_directd_base_url_from_env(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"
        os.environ["DIRECTD_BASE_URL"] = "https://api.directd.test.local"

        settings = get_settings()

        self.assertEqual(
            settings.directd_base_url,
            "https://api.directd.test.local",
        )

    def test_get_settings_rejects_insecure_directd_base_url(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"
        os.environ["DIRECTD_BASE_URL"] = "http://api.directd.test.local"

        with self.assertRaises(RuntimeError) as ctx:
            get_settings()

        self.assertIn("DIRECTD_BASE_URL", str(ctx.exception))

    def test_get_settings_reads_directd_timeout_from_env(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"
        os.environ["DIRECTD_TIMEOUT_SECONDS"] = "7.5"

        settings = get_settings()

        self.assertEqual(settings.directd_timeout_seconds, 7.5)

    def test_get_settings_rejects_invalid_directd_timeout(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"
        os.environ["DIRECTD_TIMEOUT_SECONDS"] = "0"

        with self.assertRaises(RuntimeError) as ctx:
            get_settings()

        self.assertIn("DIRECTD_TIMEOUT_SECONDS", str(ctx.exception))

    def test_get_settings_fails_when_directd_token_is_absent(self):
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"
        os.environ.pop("DIRECTD_TOKEN", None)

        with self.assertRaises(RuntimeError) as ctx:
            get_settings()

        self.assertIn("DIRECTD_TOKEN", str(ctx.exception))

    def test_get_settings_fails_when_basic_user_is_absent(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_PASSWORD"] = "secret"
        os.environ.pop("APP_BASIC_USER", None)

        with self.assertRaises(RuntimeError) as ctx:
            get_settings()

        self.assertIn("APP_BASIC_USER", str(ctx.exception))

    def test_get_settings_fails_when_basic_password_is_absent(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ.pop("APP_BASIC_PASSWORD", None)

        with self.assertRaises(RuntimeError) as ctx:
            get_settings()

        self.assertIn("APP_BASIC_PASSWORD", str(ctx.exception))

    def test_basic_auth_accepts_correct_credentials(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"

        app = create_app()
        with TestClient(app) as client:
            response = client.get("/", auth=("admin", "secret"))

        self.assertEqual(response.status_code, 200)
        self.assertIn("<input", response.text)
        self.assertIn("name=\"cnpj\"", response.text)
        self.assertIn("Gerar planilha", response.text)

    def test_basic_auth_rejects_missing_authorization_header(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"

        app = create_app()
        with TestClient(app) as client:
            response = client.get("/")

        self.assertEqual(response.status_code, 401)
        self.assertIn("WWW-Authenticate", response.headers)
        self.assertEqual(response.headers["WWW-Authenticate"], "Basic")

    def test_basic_auth_rejects_incorrect_credentials(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"

        app = create_app()
        with TestClient(app) as client:
            response = client.get("/", auth=("admin", "wrong"))

        self.assertEqual(response.status_code, 401)
        self.assertIn("WWW-Authenticate", response.headers)
        self.assertEqual(response.headers["WWW-Authenticate"], "Basic")


class HtmlFrontendTest(unittest.TestCase):
    def test_homepage_has_frontend_controls(self):
        os.environ["DIRECTD_TOKEN"] = "token-for-test"
        os.environ["APP_BASIC_USER"] = "admin"
        os.environ["APP_BASIC_PASSWORD"] = "secret"

        app = create_app()
        with TestClient(app) as client:
            response = client.get("/", headers=_basic_header())

        self.assertEqual(response.status_code, 200)
        self.assertIn('lang="pt-BR"', response.text)
        self.assertIn("Gerar planilha", response.text)
        self.assertIn("Baixar novamente", response.text)
        self.assertIn("/api/planilha", response.text)
        self.assertIn("downloadUrl", response.text)
        self.assertIn("triggerDownload", response.text)
        self.assertIn("filename\\*=", response.text)
        self.assertIn("readErrorMessage", response.text)
        self.assertNotIn("response.text()", response.text)


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
            (
                "https://example.test/api/CadastroPessoaJuridica"
                "?CNPJ=12345678000190&Token=secret-token"
            ),
        )

    def test_builds_cnpj_url_without_double_slash(self):
        from api_planilhas.config import AppSettings
        from api_planilhas.directd import build_cnpj_url

        settings = AppSettings(
            directd_token="secret-token",
            basic_user="admin",
            basic_password="secret",
            directd_base_url="https://example.test/",
        )

        url = build_cnpj_url("12345678000190", settings)

        self.assertEqual(
            url,
            "https://example.test/api/CadastroPessoaJuridica?CNPJ=12345678000190&Token=secret-token",
        )

    def test_fetch_cnpj_returns_json_payload(self):
        import json

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

        settings = AppSettings(
            directd_token="secret-token",
            basic_user="admin",
            basic_password="secret",
            directd_base_url="https://example.test",
        )

        with patch("api_planilhas.directd.urlopen", return_value=FakeResponse()):
            payload = fetch_cnpj("12345678000190", settings)

        self.assertEqual(payload["retorno"]["cnpj"], "12345678000190")

    def test_fetch_cnpj_uses_configured_timeout(self):
        import json

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

        settings = AppSettings(
            directd_token="secret-token",
            basic_user="admin",
            basic_password="secret",
            directd_base_url="https://example.test",
            directd_timeout_seconds=12.5,
        )

        call = {}

        def fake_urlopen(request, timeout):
            call["timeout"] = timeout
            return FakeResponse()

        with patch("api_planilhas.directd.urlopen", side_effect=fake_urlopen):
            fetch_cnpj("12345678000190", settings)

        self.assertEqual(call["timeout"], 12.5)

    def test_fetch_cnpj_raises_directd_error_on_http_error(self):
        from api_planilhas.config import AppSettings
        from api_planilhas.directd import DirectDError, fetch_cnpj
        from urllib.error import HTTPError

        settings = AppSettings(
            directd_token="secret-token",
            basic_user="admin",
            basic_password="secret",
            directd_base_url="https://example.test",
        )

        error = HTTPError("https://example.test/api/CadastroPessoaJuridica", 503, "Service Unavailable", None, None)

        with patch("api_planilhas.directd.urlopen", side_effect=error):
            with self.assertRaises(DirectDError) as ctx:
                fetch_cnpj("12345678000190", settings)

        self.assertIn("HTTP", str(ctx.exception))
        self.assertIn("503", str(ctx.exception))

    def test_fetch_cnpj_raises_directd_error_on_url_error(self):
        from api_planilhas.config import AppSettings
        from api_planilhas.directd import DirectDError, fetch_cnpj
        from urllib.error import URLError

        settings = AppSettings(
            directd_token="secret-token",
            basic_user="admin",
            basic_password="secret",
            directd_base_url="https://example.test",
        )

        with patch("api_planilhas.directd.urlopen", side_effect=URLError("connection failed")):
            with self.assertRaises(DirectDError):
                fetch_cnpj("12345678000190", settings)

    def test_fetch_cnpj_raises_directd_error_on_socket_timeout(self):
        import socket

        from api_planilhas.config import AppSettings
        from api_planilhas.directd import DirectDError, fetch_cnpj
        from urllib.error import URLError

        settings = AppSettings(
            directd_token="secret-token",
            basic_user="admin",
            basic_password="secret",
            directd_base_url="https://example.test",
        )

        with patch("api_planilhas.directd.urlopen", side_effect=URLError(socket.timeout("timed out"))):
            with self.assertRaises(DirectDError) as ctx:
                fetch_cnpj("12345678000190", settings)

        self.assertIn("Tempo limite", str(ctx.exception))

    def test_fetch_cnpj_raises_directd_error_on_invalid_json(self):
        import json

        from api_planilhas.config import AppSettings
        from api_planilhas.directd import DirectDError, fetch_cnpj

        class FakeResponse:
            status = 200

            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb):
                return False

            def read(self):
                return b"{invalid-json"

        settings = AppSettings(
            directd_token="secret-token",
            basic_user="admin",
            basic_password="secret",
            directd_base_url="https://example.test",
        )

        with patch("api_planilhas.directd.urlopen", return_value=FakeResponse()):
            with self.assertRaises(DirectDError) as ctx:
                fetch_cnpj("12345678000190", settings)

        self.assertIn("JSON", str(ctx.exception))

    def test_fetch_cnpj_raises_directd_error_when_retorno_is_not_dict(self):
        import json

        from api_planilhas.config import AppSettings
        from api_planilhas.directd import DirectDError, fetch_cnpj

        class FakeResponse:
            status = 200

            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb):
                return False

            def read(self):
                return json.dumps({"retorno": "nao-e-dict"}).encode("utf-8")

        settings = AppSettings(
            directd_token="secret-token",
            basic_user="admin",
            basic_password="secret",
            directd_base_url="https://example.test",
        )

        with patch("api_planilhas.directd.urlopen", return_value=FakeResponse()):
            with self.assertRaises(DirectDError):
                fetch_cnpj("12345678000190", settings)


class PlanilhaRouteTest(unittest.TestCase):
    def test_planilha_route_returns_xlsx(self):
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
                response = TestClient(create_app()).post(
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
        self.assertIn(
            'filename="importacao_oportunidade_12345678000190.xlsx"',
            response.headers["content-disposition"],
        )
        self.assertGreater(len(response.content), 100)

    def test_planilha_route_rejects_invalid_cnpj(self):
        env = {
            "DIRECTD_TOKEN": "token",
            "APP_BASIC_USER": "admin",
            "APP_BASIC_PASSWORD": "secret",
        }

        with patch.dict(os.environ, env, clear=True):
            response = TestClient(create_app()).post(
                "/api/planilha",
                data={"cnpj": "123"},
                headers=_basic_header(),
            )

        self.assertEqual(response.status_code, 400)

    def test_planilha_route_requires_basic_auth(self):
        from api_planilhas.web import create_app

        env = {
            "DIRECTD_TOKEN": "token",
            "APP_BASIC_USER": "admin",
            "APP_BASIC_PASSWORD": "secret",
        }

        with patch.dict(os.environ, env, clear=True):
            response = TestClient(create_app()).post(
                "/api/planilha",
                data={"cnpj": "12.345.678/0001-90"},
            )

        self.assertEqual(response.status_code, 401)
        self.assertEqual(response.headers["WWW-Authenticate"], "Basic")

    def test_planilha_route_maps_directd_error(self):
        from api_planilhas.directd import DirectDError
        from api_planilhas.web import create_app

        env = {
            "DIRECTD_TOKEN": "token",
            "APP_BASIC_USER": "admin",
            "APP_BASIC_PASSWORD": "secret",
        }

        with patch.dict(os.environ, env, clear=True):
            with patch(
                "api_planilhas.web.fetch_cnpj",
                side_effect=DirectDError("Falha DirectD"),
            ):
                response = TestClient(create_app()).post(
                    "/api/planilha",
                    data={"cnpj": "12345678000190"},
                    headers=_basic_header(),
                )

        self.assertEqual(response.status_code, 502)
        self.assertIn("Falha DirectD", response.text)


