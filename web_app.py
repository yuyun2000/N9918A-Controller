import importlib.util
import json
import os
import threading
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, request, send_file, send_from_directory

from sa_test_service import SATestService, ServiceError, SWITCH_IMPORT_ERROR

ROOT = Path(__file__).resolve().parent
WEB_ROOT = ROOT / "web_frontend"
AI_LOCAL_CONFIG = ROOT / "ai_config.local.json"

app = Flask(__name__, static_folder=str(WEB_ROOT), static_url_path="")
service = SATestService()


def ok(data=None, **extra):
    payload = {"ok": True}
    if data is not None:
        payload["data"] = data
    payload.update(extra)
    return jsonify(payload)


def module_available(import_name):
    return importlib.util.find_spec(import_name) is not None


def build_diagnostics():
    package_checks = [
        ("Flask", "flask"),
        ("Matplotlib", "matplotlib"),
        ("NumPy", "numpy"),
        ("PyVISA", "pyvisa"),
        ("SciPy", "scipy"),
        ("ReportLab", "reportlab"),
        ("pythonnet clr", "clr"),
    ]
    packages = [
        {
            "name": label,
            "ok": module_available(module),
            "detail": module,
        }
        for label, module in package_checks
    ]

    required_files = [
        ("Mini-Circuits DLL", ROOT / "mcl_RF_Switch_Controller64.dll"),
        ("报告 logo", ROOT / "assets" / "m5logo2022.png"),
        ("PDF 仿宋字体", ROOT / "utils" / "simfang.ttf"),
        ("PDF 黑体字体", ROOT / "utils" / "simhei.ttf"),
    ]
    doc_files = sorted((ROOT / "doc").glob("*.pdf")) if (ROOT / "doc").exists() else []
    files = [
        {
            "name": label,
            "ok": path.exists(),
            "detail": str(path),
        }
        for label, path in required_files
    ]
    files.append(
        {
            "name": "N9918A 文档",
            "ok": bool(doc_files),
            "detail": ", ".join(path.name for path in doc_files) if doc_files else "缺少 doc/*.pdf",
        }
    )

    ai_keys = ["N9918A_AI_API_KEY", "OPENAI_API_KEY", "ARK_API_KEY", "VOLCENGINE_API_KEY"]
    has_ai_local_key = False
    if AI_LOCAL_CONFIG.exists():
        try:
            has_ai_local_key = bool(json.loads(AI_LOCAL_CONFIG.read_text(encoding="utf-8-sig")).get("api_key"))
        except Exception:
            has_ai_local_key = False
    has_ai_key = has_ai_local_key or any(os.getenv(key) for key in ai_keys)
    environment = [
        {
            "name": "AI API 密钥",
            "ok": has_ai_key,
            "detail": "已通过本地配置文件配置" if has_ai_local_key else ("已通过环境变量配置" if has_ai_key else "未配置 ai_config.local.json 或 N9918A_AI_API_KEY/OPENAI_API_KEY"),
        },
        {
            "name": "N9918A_AI_BASE_URL",
            "ok": True,
            "detail": "已配置" if os.getenv("N9918A_AI_BASE_URL") or os.getenv("OPENAI_BASE_URL") else "默认 http://192.168.20.38:3000/",
        },
        {
            "name": "N9918A_AI_MODEL",
            "ok": True,
            "detail": "已配置" if os.getenv("N9918A_AI_MODEL") or os.getenv("OPENAI_MODEL") else "默认 gpt-5.5",
        },
        {
            "name": "N9918A_AI_REASONING_EFFORT",
            "ok": True,
            "detail": "已配置" if os.getenv("N9918A_AI_REASONING_EFFORT") or os.getenv("OPENAI_REASONING_EFFORT") else "默认 xhigh",
        },
    ]

    return {
        "packages": packages,
        "files": files,
        "environment": environment,
        "switch_import_error": str(SWITCH_IMPORT_ERROR) if SWITCH_IMPORT_ERROR else None,
    }


@app.errorhandler(ServiceError)
def handle_service_error(exc):
    return jsonify({"ok": False, "error": str(exc)}), 400


@app.errorhandler(Exception)
def handle_error(exc):
    return jsonify({"ok": False, "error": str(exc)}), 500


@app.get("/")
def index():
    return send_from_directory(WEB_ROOT, "index.html")


@app.get("/api/status")
def api_status():
    return ok(service.status())


@app.get("/api/mode")
def api_mode_get():
    return ok(service.mode_status())


@app.post("/api/mode")
def api_mode_post():
    data = request.get_json(silent=True) or {}
    return ok(service.switch_mode(data.get("mode")))


@app.get("/api/presets")
def api_presets():
    return ok(service.presets())


@app.get("/api/na/presets")
def api_na_presets():
    return ok(service.na_presets())


@app.post("/api/na/configure")
def api_na_configure():
    data = request.get_json(silent=True) or {}
    return ok(
        service.na_configure(
            data.get("preset_key", ""),
            points=data.get("points"),
            ifbw=data.get("ifbw"),
        )
    )


@app.post("/api/na/calibrate")
def api_na_calibrate():
    return ok(service.na_calibrate())


@app.post("/api/na/measure")
def api_na_measure():
    return ok(service.start_na_measurement())


@app.post("/api/na/stop")
def api_na_stop():
    return ok(service.stop_na_measurement())


@app.get("/api/na/result")
def api_na_result():
    return ok(service.na_result_payload())


@app.post("/api/na/data/save")
def api_na_data_save():
    return ok({"files": service.save_na_data()})


@app.post("/api/na/report/export")
def api_na_report_export():
    data = request.get_json(silent=True) or {}
    report_path = service.export_na_report(data.get("user_info"))
    return ok(
        {
            "file": str(report_path),
            "download_url": f"/api/report/download/{report_path.name}",
        }
    )


@app.get("/api/diagnostics")
def api_diagnostics():
    return ok(build_diagnostics())


@app.get("/api/result")
def api_result():
    return ok(service.result_payload())


@app.post("/api/demo/load")
def api_demo_load():
    data = request.get_json(silent=True) or {}
    return ok(
        service.load_demo_data(
            data.get("preset_key", "EMC_30MHz_1GHz"),
            data.get("duration_seconds", 15),
        )
    )


@app.post("/api/user-info")
def api_user_info():
    return ok(service.update_user_info(request.get_json(silent=True) or {}))


@app.post("/api/device/connect")
def api_device_connect():
    data = request.get_json(silent=True) or {}
    return ok(service.connect_device(data.get("ip_address", "")))


@app.post("/api/device/disconnect")
def api_device_disconnect():
    return ok(service.disconnect_device())


@app.post("/api/configure")
def api_configure():
    data = request.get_json(silent=True) or {}
    return ok(service.configure(data.get("preset_key", "")))


@app.post("/api/sa/clear")
def api_sa_clear():
    return ok(service.clear_sa_state())


@app.post("/api/switch/connect")
def api_switch_connect():
    return ok(service.connect_switch())


@app.post("/api/switch/disconnect")
def api_switch_disconnect():
    return ok(service.disconnect_switch())


@app.get("/api/switch/status")
def api_switch_status():
    return ok(service.switch_status())


@app.post("/api/switch/set")
def api_switch_set():
    data = request.get_json(silent=True) or {}
    return ok(service.set_switch_position(data.get("switch"), data.get("position")))


@app.post("/api/measure/single")
def api_measure_single():
    return ok(service.start_single_measurement())


@app.post("/api/measure/timed")
def api_measure_timed():
    data = request.get_json(silent=True) or {}
    return ok(service.start_emi_measurement(data.get("duration_seconds", 15)))


@app.post("/api/measure/stop")
def api_measure_stop():
    return ok(service.stop_measurement())


@app.post("/api/data/save")
def api_data_save():
    return ok({"files": service.save_data()})


@app.post("/api/ai/analyze")
def api_ai_analyze():
    return ok({"result": service.analyze()})


@app.post("/api/report/export")
def api_report_export():
    data = request.get_json(silent=True) or {}
    report_path = service.export_pdf(data.get("user_info"), auto_analyze=data.get("auto_analyze", True))
    return ok(
        {
            "file": str(report_path),
            "download_url": f"/api/report/download/{report_path.name}",
        }
    )


@app.get("/api/report/download/<path:filename>")
def api_report_download(filename):
    reports_dir = (ROOT / "reports").resolve()
    report_path = (reports_dir / filename).resolve()
    if reports_dir not in report_path.parents or not report_path.exists():
        raise ServiceError("报告文件不存在。")
    return send_file(report_path, as_attachment=True)


def _auto_open_browser(url):
    if os.getenv("N9918A_AUTO_OPEN_BROWSER", "1").lower() in {"0", "false", "no"}:
        return
    timer = threading.Timer(1.0, lambda: webbrowser.open(url))
    timer.daemon = True
    timer.start()


def main():
    host = os.getenv("N9918A_WEB_HOST", "127.0.0.1")
    port = int(os.getenv("N9918A_WEB_PORT", "5000"))
    url = os.getenv("N9918A_WEB_URL", f"http://{host}:{port}")
    print(f"N9918A Web Control Deck: {url}")
    _auto_open_browser(url)
    app.run(host=host, port=port, threaded=True)


if __name__ == "__main__":
    main()
