import io
import subprocess
import sys
import unittest
import zipfile
from pathlib import Path

import app


ROOT = Path(__file__).resolve().parents[1]


class AppTests(unittest.TestCase):
    def test_serverless_starts_without_writable_project_directory(self):
        code = (
            "import os, pathlib\n"
            "os.environ['VERCEL'] = '1'\n"
            "def read_only(*args, **kwargs): raise OSError('read-only filesystem')\n"
            "pathlib.Path.mkdir = read_only\n"
            "import app\n"
            "assert app.app.test_client().get('/').status_code == 200\n"
        )
        result = subprocess.run(
            [sys.executable, "-c", code], cwd=ROOT, capture_output=True, text=True
        )
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_hosted_sample_download(self):
        client = app.app.test_client()
        page = client.get("/", base_url="https://davomilik.vercel.app")
        self.assertEqual(page.status_code, 200)
        self.assertIn("Vercel serverida qayta ishlanadi", page.get_data(as_text=True))

        sample = (ROOT / "NAMUNA.csv").read_bytes()
        result = client.post(
            "/generate",
            data={"files": (io.BytesIO(sample), "NAMUNA.csv")},
            base_url="https://davomilik.vercel.app",
        )
        self.assertEqual(result.status_code, 200, result.get_data(as_text=True) if result.is_json else "")
        self.assertEqual(result.headers["X-Total-Rows"], "2")
        with zipfile.ZipFile(io.BytesIO(result.data)) as workbook:
            self.assertIn("xl/workbook.xml", workbook.namelist())
            self.assertIn("Свод".encode(), workbook.read("xl/workbook.xml"))

    def test_hosted_upload_limit_returns_413(self):
        result = app.app.test_client().post(
            "/generate",
            data={"files": (io.BytesIO(b"x" * (5 * 1024 * 1024)), "large.csv")},
            base_url="https://davomilik.vercel.app",
        )
        self.assertEqual(result.status_code, 413)


if __name__ == "__main__":
    unittest.main()
