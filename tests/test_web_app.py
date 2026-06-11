import os
import sys
import tempfile
import time
import types
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import n9918a_backend
from n9918a_backend import N9918AController
from n9918a_na_backend import (
    NA_PRESET_CONFIGS,
    N9918ANAController,
    build_na_result,
    frequency_axis,
)
from web_app import app


class FakeVisaDevice:
    def __init__(self):
        self.timeout = None
        self.commands = []

    def write(self, command):
        self.commands.append(("write", command))

    def query(self, command):
        self.commands.append(("query", command))
        if command == "*IDN?":
            return "Fake,N9918A,0,1"
        if "SWE:TIME?" in command:
            return "0.01"
        if "FDATa?" in command:
            return "-1,-8,-20,-8,-1"
        if "SDATA?" in command:
            return "0.9,0,0.4,0.1,0.1,0.0,0.4,-0.1,0.9,0"
        return "1"

    def read(self):
        return "1,2,3"

    def close(self):
        pass


class FakeResourceManager:
    last_device = None

    def open_resource(self, _resource):
        self.__class__.last_device = FakeVisaDevice()
        return self.__class__.last_device

    def close(self):
        pass


class FakePyVisa:
    class errors:
        VisaIOError = Exception

    @staticmethod
    def ResourceManager():
        return FakeResourceManager()


class FakeSwitch:
    def __init__(self):
        self.connected = True
        self.calls = []

    def set_switch(self, switch_name, position):
        self.calls.append((switch_name, position))


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
            self.assertIn("N9918A SA/NA 控制台".encode("utf-8"), page.data)
            self.assertIn("公共连接与 Switch".encode("utf-8"), page.data)
            self.assertIn("设备 IP".encode("utf-8"), page.data)
            self.assertIn("NA 天线测量".encode("utf-8"), page.data)
            self.assertIn("3dB / 10dB 双口径带宽".encode("utf-8"), page.data)
        finally:
            page.close()

        presets = self.client.get("/api/presets").get_json()
        self.assertTrue(presets["ok"])
        self.assertIn("EMC_30MHz_1GHz", presets["data"])

        status = self.client.get("/api/status").get_json()
        self.assertTrue(status["ok"])
        self.assertIn("connected", status["data"])

        mode = self.client.get("/api/mode").get_json()
        self.assertTrue(mode["ok"])
        self.assertEqual(mode["data"]["current_mode"], "SA")

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

    def test_mode_and_na_demo_flow(self):
        demo = self.client.post("/api/demo/load", json={"duration_seconds": 15}).get_json()
        self.assertTrue(demo["ok"], demo)

        mode = self.client.post("/api/mode", json={"mode": "NA"}).get_json()
        self.assertTrue(mode["ok"], mode)
        self.assertEqual(mode["data"]["current_mode"], "NA")

        presets = self.client.get("/api/na/presets").get_json()
        self.assertTrue(presets["ok"], presets)
        self.assertIn("ANT_433", presets["data"])
        self.assertIn("ANT_FULL", presets["data"])

        configured = self.client.post("/api/na/configure", json={"preset_key": "ANT_433"}).get_json()
        self.assertTrue(configured["ok"], configured)
        self.assertTrue(configured["data"]["status"]["configured"])

        calibrated = self.client.post("/api/na/calibrate").get_json()
        self.assertTrue(calibrated["ok"], calibrated)
        self.assertTrue(calibrated["data"]["status"]["calibration"]["complete"])
        steps = [event["switch_position"] for event in calibrated["data"]["status"]["calibration"]["events"]]
        self.assertEqual(steps[:4], ["B2D1", "B1D1", "B1D1", "B2D2"])

        measure = self.client.post("/api/na/measure").get_json()
        self.assertTrue(measure["ok"], measure)
        self.poll_until_idle()
        result = self.client.get("/api/na/result").get_json()
        self.assertTrue(result["ok"], result)
        data = result["data"]
        self.assertTrue(data["series"]["frequency_mhz"])
        self.assertTrue(data["smith"]["real"])
        self.assertTrue(data["primary_valley"])
        self.assertIn("absolute_3db", data["bandwidths"])
        self.assertGreater(len(data["valleys"]), 0)

    def test_na_save_and_report_export_in_demo(self):
        with tempfile.TemporaryDirectory() as tmp:
            old_cwd = os.getcwd()
            os.chdir(tmp)
            try:
                self.client.post("/api/demo/load", json={"duration_seconds": 15})
                self.client.post("/api/mode", json={"mode": "NA"})
                self.client.post("/api/na/configure", json={"preset_key": "ANT_2450"})
                self.client.post("/api/na/calibrate")
                self.client.post("/api/na/measure")
                self.poll_until_idle()

                saved = self.client.post("/api/na/data/save").get_json()
                self.assertTrue(saved["ok"], saved)
                self.assertGreaterEqual(len(saved["data"]["files"]), 3)
                for file_path in saved["data"]["files"]:
                    self.assertTrue(Path(file_path).exists(), file_path)

                report = self.client.post("/api/na/report/export", json={"user_info": {"eut": "Antenna"}}).get_json()
                self.assertTrue(report["ok"], report)
                output = Path(report["data"]["file"])
                self.assertTrue(output.exists())
                self.assertTrue(output.read_text(encoding="utf-8").lstrip().startswith("<!doctype html>"))
            finally:
                os.chdir(old_cwd)


class BackendRegressionTest(unittest.TestCase):
    def test_sa_scpi_query_strings_are_valid(self):
        previous_pyvisa = n9918a_backend.pyvisa
        n9918a_backend.pyvisa = FakePyVisa
        try:
            controller = N9918AController(ip_address="192.0.2.1")
            self.assertTrue(controller.connect())
            controller.start_freq = 1
            controller.stop_freq = 3
            controller.n_points = 3
            frequencies, amplitudes = controller.read_trace_data()
        finally:
            n9918a_backend.pyvisa = previous_pyvisa

        self.assertEqual(frequencies, [1.0, 2.0, 3.0])
        self.assertEqual(amplitudes, [1.0, 2.0, 3.0])
        commands = [command for _kind, command in FakeResourceManager.last_device.commands]
        self.assertIn("*IDN?", commands)
        self.assertIn("INST:SEL 'SA';*OPC?", commands)
        self.assertIn(":INIT:IMM;*OPC?", commands)
        self.assertIn(":TRAC:DATA?", commands)
        self.assertFalse(any("[WARN]" in command for command in commands))

    def test_na_calibration_switch_and_scpi_sequence(self):
        device = FakeVisaDevice()
        controller = N9918ANAController(ip_address="192.0.2.1")
        controller.device = device
        controller.connected = True
        controller.configure_preset("ANT_433")
        switch = FakeSwitch()
        result = controller.perform_calibration(switch)

        self.assertTrue(result["complete"])
        self.assertEqual(
            switch.calls,
            [("B", 2), ("D", 1), ("B", 1), ("D", 1), ("B", 2), ("D", 2)],
        )
        commands = [command for _kind, command in device.commands]
        self.assertIn("CORR:COLL:METH:QCAL:CAL 1", commands)
        self.assertLess(commands.index("CORR:COLL:INT 1;*OPC?"), commands.index("CORR:COLL:LOAD 1;*OPC?"))
        self.assertIn("CORR:COLL:SAVE 0", commands)


class NAAlgorithmTest(unittest.TestCase):
    def test_valley_and_bandwidth_calculation(self):
        frequencies = [0, 1e6, 2e6, 3e6, 4e6]
        s11_db = [0, -5, -20, -5, 0]
        real = [0.9, 0.5, 0.1, 0.5, 0.9]
        imag = [0, 0.1, 0, -0.1, 0]
        result = build_na_result(
            frequencies,
            s11_db,
            real,
            imag,
            {"start_freq": 0, "stop_freq": 4e6, "points": 5, "full_sweep": False},
            "TEST",
        )

        self.assertAlmostEqual(result["primary_valley"]["frequency_mhz"], 2.0)
        self.assertAlmostEqual(result["primary_valley"]["s11_db"], -20.0)
        self.assertAlmostEqual(result["bandwidths"]["absolute_10db"]["left_hz"], 1.333333e6, delta=2)
        self.assertAlmostEqual(result["bandwidths"]["absolute_10db"]["right_hz"], 2.666667e6, delta=2)
        self.assertAlmostEqual(result["bandwidths"]["relative_3db"]["width_hz"], 0.4e6, delta=2)
        self.assertTrue(result["smith"]["markers"])

    def test_full_sweep_lists_multiple_valleys_without_smith(self):
        frequencies = frequency_axis(1e6, 10e6, 10)
        s11_db = [0, -10, 0, 0, -20, 0, 0, -15, 0, 0]
        result = build_na_result(
            frequencies,
            s11_db,
            [0.1] * 10,
            [0.0] * 10,
            {"start_freq": 1e6, "stop_freq": 10e6, "points": 10, "full_sweep": True},
            "ANT_FULL",
        )

        self.assertIsNone(result["smith"])
        self.assertEqual([round(v["frequency_mhz"]) for v in result["valleys"]], [2, 5, 8])
        self.assertAlmostEqual(result["primary_valley"]["frequency_mhz"], 5.0)


if __name__ == "__main__":
    unittest.main()
