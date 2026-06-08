import sys
import unittest
from datetime import UTC, datetime, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from api_planilhas.jobs import JobNotFoundError, JobStore


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
            self.assertFalse(loaded.download_ready)
            self.assertEqual(
                store.get_input_cnpjs(job.job_id),
                ["12345678000190", "11222333000144"],
            )
            self.assertTrue(store.input_path(job.job_id).exists())

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

    def test_download_ready_requires_completed_status_and_output_file(self):
        with TemporaryDirectory() as temp_dir:
            store = JobStore(Path(temp_dir))
            store.initialize()
            job = store.create_job(["12345678000190"])

            store.output_path(job.job_id).write_bytes(b"xlsx")
            self.assertFalse(store.get_job(job.job_id).download_ready)

            store.mark_completed(job.job_id)
            self.assertTrue(store.get_job(job.job_id).download_ready)

    def test_missing_job_raises_job_not_found(self):
        with TemporaryDirectory() as temp_dir:
            store = JobStore(Path(temp_dir))
            store.initialize()

            with self.assertRaises(JobNotFoundError):
                store.get_job("missing")
            with self.assertRaises(JobNotFoundError):
                store.mark_processing("missing")
            with self.assertRaises(JobNotFoundError):
                store.get_input_cnpjs("missing")

    def test_cleanup_expired_removes_old_jobs(self):
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

            with self.assertRaises(JobNotFoundError):
                store.get_job(job.job_id)
            self.assertFalse((Path(temp_dir) / job.job_id).exists())


if __name__ == "__main__":
    unittest.main()
