import json
import os
import sys
import tempfile
import time
import types
import unittest
from unittest.mock import patch
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from chat import ChatBot, sys_prompt
import n9918a_backend
from n9918a_backend import (
    N9918AController,
    collapse_contiguous_indices,
    dbm_to_dbuv,
    dbuv_to_dbm,
    get_emission_limit_info,
    get_fcc_ce_limits,
    linear_average_dbuv,
    post_process_peak_search,
)
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
        self.start_freq = 1.0
        self.stop_freq = 3.0
        self.n_points = 3
        self.rbw = 100000.0
        self.vbw = 100000.0
        self.amplitude_unit = "DBM"

    def write(self, command):
        self.commands.append(("write", command))
        upper = command.upper()
        parts = command.split()
        if upper.startswith(":SENS:FREQ:STAR") and len(parts) >= 2:
            self.start_freq = float(parts[-1])
        elif upper.startswith(":SENS:FREQ:STOP") and len(parts) >= 2:
            self.stop_freq = float(parts[-1])
        elif upper.startswith(":SENS:SWE:POIN") and len(parts) >= 2:
            self.n_points = int(float(parts[-1]))
        elif upper.startswith(":SENS:BAND:RES ") and len(parts) >= 2:
            self.rbw = float(parts[-1])
        elif upper.startswith(":SENS:BAND:VID ") and len(parts) >= 2:
            self.vbw = float(parts[-1])
        elif upper.startswith(":SENS:AMPL:UNIT") and len(parts) >= 2:
            self.amplitude_unit = parts[-1].upper()

    def query(self, command):
        self.commands.append(("query", command))
        if command == "*IDN?":
            return "Fake,N9918A,0,1"
        if command == "SYST:ERR?":
            return '+0,"No error"'
        if "FREQ:STAR?" in command:
            return str(self.start_freq)
        if "FREQ:STOP?" in command:
            return str(self.stop_freq)
        if "SWE:POIN?" in command:
            return str(self.n_points)
        if "SWE:TIME?" in command:
            return "0.01"
        if "BAND:RES?" in command:
            return str(self.rbw)
        if "BAND:VID?" in command:
            return str(self.vbw)
        if "TRAC1:XVAL?" in command:
            if self.n_points <= 1:
                return str(self.start_freq)
            step = (self.stop_freq - self.start_freq) / (self.n_points - 1)
            return ",".join(str(self.start_freq + i * step) for i in range(self.n_points))
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
            self.assertIn("回波损耗".encode("utf-8"), page.data)
            self.assertIn("VSWR".encode("utf-8"), page.data)
            self.assertIn("VSWR 驻波比曲线".encode("utf-8"), page.data)
            self.assertIn("中心与带宽端点详情".encode("utf-8"), page.data)
            self.assertIn("理想频点附近损耗".encode("utf-8"), page.data)
            self.assertIn("Smith Chart 与阻抗".encode("utf-8"), page.data)
            self.assertIn("手动清理 Trace".encode("utf-8"), page.data)
        finally:
            page.close()

        app_js = (ROOT / "web_frontend" / "app.js").read_text(encoding="utf-8")
        self.assertIn('const switchOrder = ["A", "B", "D", "C"];', app_js)
        self.assertIn("pos-one", app_js)
        self.assertIn("pos-two", app_js)
        self.assertIn("renderVswr", app_js)
        self.assertIn("drawSmithGrid", app_js)

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
        self.assertIn("N9918A_WEB_URL=http://127.0.0.1:5000", run_bat)
        web_app_py = (ROOT / "web_app.py").read_text(encoding="utf-8")
        self.assertIn("webbrowser.open", web_app_py)
        gitignore = (ROOT / ".gitignore").read_text(encoding="utf-8")
        self.assertIn("ai_config.local.json", gitignore)

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
        self.assertEqual(result["data"]["detector_mode"], "PEAK")

        timed = self.client.post("/api/measure/timed", json={"duration_seconds": 15}).get_json()
        self.assertTrue(timed["ok"], timed)
        self.poll_until_idle()
        result = self.client.get("/api/result").get_json()
        self.assertTrue(result["ok"], result)
        self.assertTrue(result["data"]["status"]["has_emi_data"])
        self.assertTrue(result["data"]["measurement_summary"]["demo"])
        self.assertIn("QUASI_PEAK", result["data"]["modes"])
        self.assertEqual(result["data"]["detector_mode"], "QUASI_PEAK")

        cleared = self.client.post("/api/sa/clear").get_json()
        self.assertTrue(cleared["ok"], cleared)
        result = self.client.get("/api/result").get_json()
        self.assertTrue(result["ok"], result)
        self.assertFalse(result["data"]["series"])
        self.assertFalse(result["data"]["peaks"])
        self.assertFalse(result["data"]["status"]["has_single_data"])
        self.assertFalse(result["data"]["status"]["has_emi_data"])

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
        self.assertIn("ANT_868", presets["data"])
        self.assertNotIn("ANT_898", presets["data"])
        self.assertIn("ANT_FULL", presets["data"])
        self.assertEqual(presets["data"]["ANT_868"]["start_freq"], 768e6)
        self.assertEqual(presets["data"]["ANT_868"]["stop_freq"], 968e6)

        configured = self.client.post("/api/na/configure", json={"preset_key": "ANT_433"}).get_json()
        self.assertTrue(configured["ok"], configured)
        self.assertTrue(configured["data"]["status"]["configured"])

        calibrated = self.client.post("/api/na/calibrate").get_json()
        self.assertTrue(calibrated["ok"], calibrated)
        self.assertTrue(calibrated["data"]["status"]["calibration"]["complete"])
        steps = [event["switch_position"] for event in calibrated["data"]["status"]["calibration"]["events"]]
        self.assertEqual(steps[:4], ["B2C1", "B1C1", "B1C1", "B2C2"])

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
        self.assertIn("return_loss_db", data["primary_valley"])
        self.assertIn("vswr", data["primary_valley"])
        self.assertTrue(data["points_of_interest"])
        self.assertTrue(data["target_summary"])
        self.assertTrue(data["target_window"])
        self.assertAlmostEqual(data["target_summary"]["target_frequency_mhz"], 433.0)
        self.assertGreater(len(data["valleys"]), 0)

    def test_na_save_and_report_export_in_demo(self):
        output = None
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
                self.assertTrue(output.read_bytes().startswith(b"%PDF"))
                download = self.client.get(report["data"]["download_url"])
                try:
                    self.assertEqual(download.status_code, 200)
                    self.assertTrue(download.data.startswith(b"%PDF"))
                finally:
                    download.close()
            finally:
                os.chdir(old_cwd)
                if output and output.exists():
                    output.unlink()
                    reports_dir = output.parent
                    if reports_dir.exists() and not any(reports_dir.iterdir()):
                        reports_dir.rmdir()


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
        self.assertIn("INIT:CONT OFF", commands)
        self.assertIn(":SENS:AMPL:UNIT DBUV", commands)
        self.assertIn(":TRAC1:TYPE CLRW", commands)
        self.assertIn(":INIT:IMM;*OPC?", commands)
        self.assertIn(":TRAC1:DATA?", commands)
        self.assertFalse(any("[WARN]" in command for command in commands))

    def test_sa_config_applies_screening_trace_setup(self):
        previous_pyvisa = n9918a_backend.pyvisa
        n9918a_backend.pyvisa = FakePyVisa
        try:
            controller = N9918AController(ip_address="192.0.2.1")
            self.assertTrue(controller.connect())
            self.assertTrue(controller.configure_settings("EMC_30MHz_1GHz"))
        finally:
            n9918a_backend.pyvisa = previous_pyvisa

        commands = [command for _kind, command in FakeResourceManager.last_device.commands]
        self.assertIn(":SENS:AMPL:UNIT DBUV", commands)
        self.assertIn(":SENS:BAND:RES:AUTO OFF", commands)
        self.assertIn(":SENS:BAND:VID:AUTO OFF", commands)
        self.assertIn(":SENS:DET POS", commands)
        self.assertIn(":TRAC1:TYPE CLRW", commands)
        self.assertIn("SYST:ERR?", commands)

    def test_sa_manual_clear_blanks_fieldfox_trace(self):
        previous_pyvisa = n9918a_backend.pyvisa
        n9918a_backend.pyvisa = FakePyVisa
        try:
            controller = N9918AController(ip_address="192.0.2.1")
            self.assertTrue(controller.connect())
            self.assertTrue(controller.clear_sa_display_state())
        finally:
            n9918a_backend.pyvisa = previous_pyvisa

        commands = [command for _kind, command in FakeResourceManager.last_device.commands]
        self.assertIn("*CLS", commands)
        self.assertIn("INST:SEL 'SA';*OPC?", commands)
        self.assertIn("INIT:CONT OFF", commands)
        self.assertIn(":SENS:AMPL:UNIT DBUV", commands)
        self.assertIn(":TRAC1:TYPE CLRW", commands)
        self.assertIn(":TRAC1:TYPE BLAN", commands)
        self.assertIn("SYST:ERR?", commands)

    def test_sa_screening_math_and_limit_grouping(self):
        self.assertAlmostEqual(dbm_to_dbuv(-107), -0.0103, places=3)
        self.assertAlmostEqual(dbuv_to_dbm(dbm_to_dbuv(-25)), -25.0, places=6)
        self.assertAlmostEqual(linear_average_dbuv([0, 20]), 14.807, places=3)

        info = get_emission_limit_info(100e6)
        self.assertEqual(info["unit"], "dBuV/m")
        self.assertEqual(info["measurement_type"], "radiated_3m_screening")
        self.assertAlmostEqual(info["fcc_limit"], 43.5)
        above_peak = get_emission_limit_info(1.5e9, detector_type="PEAK")
        self.assertEqual(above_peak["fcc_detector"], "PEAK")
        self.assertAlmostEqual(above_peak["fcc_limit"], 74.0)
        fcc_peak, ce_peak = get_fcc_ce_limits(1.5e9, detector_type="PEAK")
        self.assertEqual((fcc_peak, ce_peak), (74.0, 70.0))

        grouped = collapse_contiguous_indices([1, 2, 3, 7, 8], [0, 2, 5, 3, 0, 0, 0, 9, 4])
        self.assertEqual(grouped, [2, 7])

        frequencies = [30e6, 31e6, 32e6, 100e6, 101e6, 500e6]
        amplitudes = [10, 45, 43, 50, 49, 10]
        peaks = post_process_peak_search(frequencies, amplitudes, peak_distance=1, min_prominence=0.1)
        exceeded = [peak for peak in peaks if peak["exceed_fcc"] or peak["exceed_ce"]]
        self.assertEqual([round(peak["frequency_mhz"]) for peak in exceeded], [31, 100])

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
            [("B", 2), ("C", 1), ("B", 1), ("C", 1), ("B", 2), ("C", 2)],
        )
        commands = [command for _kind, command in device.commands]
        self.assertIn("CORR:COLL:METH:QCAL:CAL 1", commands)
        self.assertLess(commands.index("CORR:COLL:INT 1;*OPC?"), commands.index("CORR:COLL:LOAD 1;*OPC?"))
        self.assertIn("CORR:COLL:SAVE 0", commands)


class AIClientRegressionTest(unittest.TestCase):
    def test_ai_prompt_keeps_utf8_chinese(self):
        self.assertIn("SA 筛查结果分析助手", sys_prompt)
        self.assertIn("整改建议", sys_prompt)
        self.assertGreater(sum(1 for char in sys_prompt if "\u4e00" <= char <= "\u9fff"), 100)
        self.assertEqual(sys_prompt.count("?"), 0)

    def test_ai_client_uses_responses_api_payload(self):
        captured = {}

        class FakeResponse:
            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb):
                return False

            def read(self):
                return b'{"output":[{"content":[{"type":"output_text","text":"analysis ok"}]}]}'

        def fake_urlopen(request, timeout=None):
            captured["url"] = request.full_url
            captured["headers"] = dict(request.header_items())
            captured["body"] = json.loads(request.data.decode("utf-8"))
            captured["timeout"] = timeout
            return FakeResponse()

        bot = ChatBot(
            api_key="test-key",
            base_url="http://192.0.2.10:3000/",
            model="gpt-5.5",
            system_message="system prompt",
            reasoning_effort="xhigh",
            max_output_tokens=123,
        )
        with patch("urllib.request.urlopen", fake_urlopen):
            response = bot.chat_no_stream("analyze this")

        self.assertEqual(response.output_text, "analysis ok")
        self.assertEqual(captured["url"], "http://192.0.2.10:3000/v1/responses")
        self.assertEqual(captured["body"]["model"], "gpt-5.5")
        self.assertEqual(captured["body"]["instructions"], "system prompt")
        self.assertEqual(captured["body"]["input"], "analyze this")
        self.assertEqual(captured["body"]["reasoning"], {"effort": "xhigh"})
        self.assertEqual(captured["body"]["max_output_tokens"], 123)
        self.assertFalse(captured["body"]["store"])
        self.assertNotIn("messages", captured["body"])
        self.assertNotIn("chat/completions", captured["url"])

    def test_ai_output_text_extraction_supports_proxy_shapes(self):
        self.assertEqual(ChatBot.extract_output_text({"output_text": "direct"}), "direct")
        self.assertEqual(
            ChatBot.extract_output_text(
                {"choices": [{"message": {"content": "chat-like compatibility"}}]}
            ),
            "chat-like compatibility",
        )


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
            {"start_freq": 0, "stop_freq": 4e6, "points": 5, "full_sweep": False, "target_freq": 2e6},
            "TEST",
        )

        self.assertAlmostEqual(result["primary_valley"]["frequency_mhz"], 2.0)
        self.assertAlmostEqual(result["primary_valley"]["s11_db"], -20.0)
        self.assertAlmostEqual(result["primary_valley"]["return_loss_db"], 20.0)
        self.assertAlmostEqual(result["primary_valley"]["vswr"], 1.2222, places=3)
        self.assertAlmostEqual(result["bandwidths"]["absolute_10db"]["left_hz"], 1.333333e6, delta=2)
        self.assertAlmostEqual(result["bandwidths"]["absolute_10db"]["right_hz"], 2.666667e6, delta=2)
        self.assertAlmostEqual(result["bandwidths"]["absolute_10db"]["left_vswr"], 1.925, places=3)
        self.assertAlmostEqual(result["bandwidths"]["relative_3db"]["width_hz"], 0.4e6, delta=2)
        self.assertAlmostEqual(result["target_summary"]["target_frequency_mhz"], 2.0)
        self.assertAlmostEqual(result["target_summary"]["frequency_error_mhz"], 0.0)
        self.assertEqual(result["target_summary"]["status"], "good")
        self.assertEqual(result["points_of_interest"][1]["type"], "target")
        self.assertTrue(result["target_window"])
        self.assertTrue(result["points_of_interest"])
        self.assertTrue(result["smith"]["markers"])
        self.assertEqual(result["smith"]["reference_ohm"], 50.0)
        self.assertTrue(any(marker["type"] == "target" for marker in result["smith"]["markers"]))
        self.assertTrue(any("impedance_label" in marker for marker in result["smith"]["markers"]))

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
