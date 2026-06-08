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


def _settings() -> AppSettings:
    return AppSettings(
        directd_token="token",
        basic_user="admin",
        basic_password="secret",
        directd_batch_delay_seconds=0.25,
    )


def _payload(cnpj: str) -> dict:
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


class BatchProcessorTest(unittest.TestCase):
    def test_processes_successes_and_errors_without_stopping(self):
        with TemporaryDirectory() as temp_dir:
            store = JobStore(Path(temp_dir))
            store.initialize()
            job = store.create_job(
                ["11.222.333/0001-44", "123", "12.345.678/0001-90"]
            )
            fetched: list[str] = []
            sleeps: list[float] = []

            def fake_fetch(cnpj, settings):
                fetched.append(cnpj)
                if cnpj == "11222333000144":
                    raise DirectDError("DirectD indisponivel")
                return _payload(cnpj)

            process_job(
                job.job_id,
                store,
                _settings(),
                fetcher=fake_fetch,
                sleeper=sleeps.append,
            )

            loaded = store.get_job(job.job_id)
            self.assertEqual(loaded.status, "completed")
            self.assertEqual(loaded.processed, 3)
            self.assertEqual(loaded.success, 1)
            self.assertEqual(
                loaded.errors,
                [
                    {"cnpj": "11222333000144", "message": "DirectD indisponivel"},
                    {"cnpj": "123", "message": "CNPJ deve conter 14 digitos"},
                ],
            )
            self.assertEqual(fetched, ["11222333000144", "12345678000190"])
            self.assertEqual(sleeps, [0.25])
            self.assertTrue(store.output_path(job.job_id).exists())

            with zipfile.ZipFile(store.output_path(job.job_id)) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")

            self.assertIn("EMPRESA 12345678000190", sheet_xml)
            self.assertNotIn("EMPRESA 11222333000144", sheet_xml)

    def test_marks_job_failed_on_unexpected_structural_failure(self):
        with TemporaryDirectory() as temp_dir:
            store = JobStore(Path(temp_dir))
            store.initialize()
            job = store.create_job(["12.345.678/0001-90"])

            def broken_fetch(_cnpj, _settings):
                raise RuntimeError("falha estrutural")

            process_job(job.job_id, store, _settings(), fetcher=broken_fetch)

            loaded = store.get_job(job.job_id)
            self.assertEqual(loaded.status, "failed")
            self.assertEqual(loaded.processed, 0)
            self.assertEqual(loaded.success, 0)
            self.assertEqual(
                loaded.errors,
                [{"cnpj": "", "message": "falha estrutural"}],
            )
            self.assertFalse(store.output_path(job.job_id).exists())


if __name__ == "__main__":
    unittest.main()
