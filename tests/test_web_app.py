import os
import sys
import tempfile
import time
import types
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from web_app import app


class WebAppSmokeTest(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()
        self.client.post("/api/device/disconnect")

    def poll_until_idle(self, timeout=3.0):
        deadline = time.time() + timeout
        while time.time() < deadline:
            payload = self.client.get("/api/status").get_json()
            self.assertTrue(payload["ok"], payload)
            if not payload["data"]["measurement_in_progress"]:
                return payload["data"]
            time.sleep(0.05)
        self.fail("measurement did not finish before timeout")

    def test_static_page_and_base_api(self):
        page = self.client.get("/")
        try:
            self.assertEqual(page.status_code, 200)
            self.assertIn("N9918A SA/EMC 控制台".encode("utf-8"), page.data)
            self.assertIn("连接与配置说明".encode("utf-8"), page.data)
        finally:
            page.close()

        presets = self.client.get("/api/presets").get_json()
        self.assertTrue(presets["ok"])
        self.assertIn("EMC_30MHz_1GHz", presets["data"])

        status = self.client.get("/api/status").get_json()
        self.assertTrue(status["ok"])
        self.assertIn("connected", status["data"])

    def test_launch_scripts_point_to_web_by_default(self):
        run_bat = (ROOT / "run.bat").read_text(encoding="utf-8")
        self.assertIn("web_app.py", run_bat)

    def test_diagnostics_api(self):
        diagnostics = self.client.get("/api/diagnostics").get_json()
        self.assertTrue(diagnostics["ok"], diagnostics)
        data = diagnostics["data"]
        self.assertIn("packages", data)
        self.assertIn("files", data)
        self.assertIn("environment", data)
        self.assertTrue(any(item["name"] == "报告 logo" for item in data["files"]))

    def test_demo_result_and_measurement_flow(self):
        demo = self.client.post("/api/demo/load", json={"duration_seconds": 15}).get_json()
        self.assertTrue(demo["ok"], demo)
        self.assertTrue(demo["data"]["status"]["demo_mode"])
        self.assertGreater(len(demo["data"]["peaks"]), 0)
        self.assertTrue(demo["data"]["series"])

        single = self.client.post("/api/measure/single").get_json()
        self.assertTrue(single["ok"], single)
        self.poll_until_idle()
        result = self.client.get("/api/result").get_json()
        self.assertTrue(result["ok"], result)
        self.assertTrue(result["data"]["series"])
        self.assertFalse(result["data"]["status"]["has_emi_data"])

        timed = self.client.post("/api/measure/timed", json={"duration_seconds": 15}).get_json()
        self.assertTrue(timed["ok"], timed)
        self.poll_until_idle()
        result = self.client.get("/api/result").get_json()
        self.assertTrue(result["ok"], result)
        self.assertTrue(result["data"]["status"]["has_emi_data"])
        self.assertTrue(result["data"]["measurement_summary"]["demo"])
        self.assertIn("QUASI_PEAK", result["data"]["modes"])

    def test_demo_save_data(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_cwd = os.getcwd()
            os.chdir(tmp)
            try:
                demo = self.client.post("/api/demo/load", json={"duration_seconds": 15}).get_json()
                self.assertTrue(demo["ok"], demo)

                saved = self.client.post("/api/data/save").get_json()
                self.assertTrue(saved["ok"], saved)
                files = saved["data"]["files"]
                self.assertGreaterEqual(len(files), 4)
                for file_path in files:
                    self.assertTrue(Path(file_path).exists(), file_path)
            finally:
                os.chdir(old_cwd)

    def test_demo_report_export_without_reportlab(self):
        demo = self.client.post("/api/demo/load", json={"duration_seconds": 15}).get_json()
        self.assertTrue(demo["ok"], demo)

        fake_module = types.ModuleType("utils.create_pdf")

        def fake_generate_test_report(filename, **_kwargs):
            Path(filename).parent.mkdir(parents=True, exist_ok=True)
            Path(filename).write_bytes(b"%PDF-1.4\n% fake smoke report\n")

        fake_module.generate_test_report = fake_generate_test_report
        previous = sys.modules.get("utils.create_pdf")
        sys.modules["utils.create_pdf"] = fake_module
        output = None
        try:
            report = self.client.post(
                "/api/report/export",
                json={
                    "auto_analyze": False,
                    "user_info": {"eut": "DemoDevice"},
                },
            ).get_json()
        finally:
            if previous is None:
                sys.modules.pop("utils.create_pdf", None)
            else:
                sys.modules["utils.create_pdf"] = previous

        self.assertTrue(report["ok"], report)
        output = Path(report["data"]["file"])
        self.assertTrue(output.exists())
        self.assertIn("/api/report/download/", report["data"]["download_url"])

        download = self.client.get(report["data"]["download_url"])
        try:
            self.assertEqual(download.status_code, 200)
            self.assertTrue(download.data.startswith(b"%PDF"))
        finally:
            download.close()

        try:
            if output and output.exists():
                output.unlink()
            reports_dir = ROOT / "reports"
            if reports_dir.exists() and not any(reports_dir.iterdir()):
                reports_dir.rmdir()
        except OSError:
            pass


if __name__ == "__main__":
    unittest.main()
