import json
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

    def test_resumes_from_saved_payloads_without_refetching_completed_cnpjs(self):
        with TemporaryDirectory() as temp_dir:
            store = JobStore(Path(temp_dir))
            store.initialize()
            job = store.create_job(["11.222.333/0001-44", "12.345.678/0001-90"])
            store.record_success_payload(
                job.job_id,
                "11222333000144",
                _payload("11222333000144"),
            )
            fetched: list[str] = []

            def fake_fetch(cnpj, settings):
                fetched.append(cnpj)
                return _payload(cnpj)

            process_job(job.job_id, store, _settings(), fetcher=fake_fetch)

            loaded = store.get_job(job.job_id)
            self.assertEqual(loaded.status, "completed")
            self.assertEqual(loaded.processed, 2)
            self.assertEqual(loaded.success, 2)
            self.assertEqual(fetched, ["12345678000190"])

            with zipfile.ZipFile(store.output_path(job.job_id)) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")

            self.assertIn("EMPRESA 11222333000144", sheet_xml)
            self.assertIn("EMPRESA 12345678000190", sheet_xml)

    def test_enriches_socios_with_cpf_phones_before_saving_payload(self):
        with TemporaryDirectory() as temp_dir:
            store = JobStore(Path(temp_dir))
            store.initialize()
            job = store.create_job(["12.345.678/0001-90"])
            fetched_cnpjs: list[str] = []
            fetched_cpfs: list[str] = []
            sleeps: list[float] = []

            def fake_fetch_cnpj(cnpj, settings):
                fetched_cnpjs.append(cnpj)
                return {
                    "retorno": {
                        "cnpj": cnpj,
                        "razaoSocial": "EMPRESA COM SOCIOS LTDA",
                        "telefones": [],
                        "emails": [],
                        "enderecos": [],
                        "socios": [
                            {
                                "nome": "JOAO SOCIO",
                                "documento": "111.111.111-11",
                            },
                            {
                                "nome": "MARIA SOCIA",
                                "documento": "222.222.222-22",
                            },
                        ],
                    }
                }

            def fake_fetch_cpf(cpf, settings):
                fetched_cpfs.append(cpf)
                if cpf == "111.111.111-11":
                    return {
                        "retorno": {
                            "telefones": [
                                {"telefoneComDDD": "44999990000"},
                                {"telefoneComDDD": "4430281122"},
                            ]
                        }
                    }
                return {"retorno": {"telefones": []}}

            process_job(
                job.job_id,
                store,
                _settings(),
                fetcher=fake_fetch_cnpj,
                cpf_fetcher=fake_fetch_cpf,
                sleeper=sleeps.append,
            )

            self.assertEqual(fetched_cnpjs, ["12345678000190"])
            self.assertEqual(fetched_cpfs, ["111.111.111-11", "222.222.222-22"])
            self.assertEqual(sleeps, [0.25, 0.25])

            saved_payload = json.loads(
                store.result_path(job.job_id, "12345678000190").read_text(
                    encoding="utf-8"
                )
            )
            socios = saved_payload["retorno"]["socios"]
            self.assertEqual(
                socios[0]["telefonesSocio"],
                ["44999990000", "4430281122"],
            )
            self.assertEqual(socios[1]["telefonesSocio"], [])

            with zipfile.ZipFile(store.output_path(job.job_id)) as archive:
                sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")

            self.assertIn("JOAO SOCIO", sheet_xml)
            self.assertIn("MARIA SOCIA", sheet_xml)
            self.assertIn("44999990000;4430281122", sheet_xml)

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
