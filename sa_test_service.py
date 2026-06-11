import os
import re
import tempfile
import threading
import time
import math
import random
from pathlib import Path

import matplotlib

matplotlib.use("Agg", force=True)
import matplotlib.pyplot as plt

from n9918a_backend import (
    N9918AController,
    get_fcc_ce_limits,
    post_process_peak_search,
    save_emi_measurement_data,
    save_peak_analysis,
    save_spectrum_data,
)
from n9918a_na_backend import (
    NA_PRESET_CONFIGS,
    N9918ANAController,
    N9918ANAError,
    build_na_result,
    export_na_report,
    frequency_axis,
    save_na_measurement_data,
)

ROOT = Path(__file__).resolve().parent
try:
    from Switch import MiniCircuitsSwitchController

    SWITCH_IMPORT_ERROR = None
except Exception as exc:  # pragma: no cover - depends on Windows DLL/runtime.
    MiniCircuitsSwitchController = None
    SWITCH_IMPORT_ERROR = exc


class ServiceError(RuntimeError):
    """Raised for user-facing workflow errors."""


class SATestService:
    """Hardware workflow service shared by the web API."""

    def __init__(self, default_ip="192.168.20.233"):
        self.controller = N9918AController(ip_address=default_ip)
        self.na_controller = N9918ANAController(ip_address=default_ip)
        self.switch_controller = MiniCircuitsSwitchController() if MiniCircuitsSwitchController else None
        self.lock = threading.RLock()
        self.stop_event = threading.Event()
        self.measurement_thread = None
        self.measurement_in_progress = False
        self.measurement_kind = None
        self.current_mode = "SA"
        self.switching_mode = False
        self.progress_message = "就绪"
        self.last_error = None

        self.current_frequencies = None
        self.current_amplitudes = None
        self.current_peaks = None
        self.emi_results = {}
        self.last_ai_result = ""
        self.last_report_path = None
        self.demo_mode = False
        self.na_config_key = None
        self.na_config = None
        self.na_result = None
        self.na_calibration = self._empty_na_calibration()
        self.na_last_report_path = None

        self.user_info = {
            "customer": "M5Stack",
            "eut": "产品A",
            "model": "Model-X",
            "engineer": "张工程师",
            "remark": "首次测试",
        }

    def status(self):
        with self.lock:
            status = self.controller.get_current_status()
            status.update(
                {
                    "connected": self.controller.connected or self.na_controller.connected,
                    "current_mode": self.current_mode,
                    "switching_mode": self.switching_mode,
                    "na_connected": self.na_controller.connected,
                    "na_configured": bool(self.na_config),
                    "na_calibrated": bool(self.na_calibration.get("complete")),
                    "measurement_in_progress": self.measurement_in_progress,
                    "measurement_kind": self.measurement_kind,
                    "progress_message": self.progress_message,
                    "last_error": self.last_error,
                    "has_single_data": bool(self.current_frequencies and self.current_amplitudes),
                    "has_emi_data": bool(self.emi_results),
                    "has_ai_result": bool(self.last_ai_result),
                    "last_report": str(self.last_report_path) if self.last_report_path else None,
                    "demo_mode": self.demo_mode,
                    "user_info": self.user_info.copy(),
                }
            )
            return status

    def presets(self):
        return self.controller.get_preset_configs()

    @staticmethod
    def _empty_na_calibration():
        return {
            "complete": False,
            "in_progress": False,
            "events": [],
            "error": None,
            "last_scpi": None,
            "switch_position": None,
        }

    def mode_status(self):
        with self.lock:
            return {
                "current_mode": self.current_mode,
                "connected": self.controller.connected or self.na_controller.connected,
                "switching": self.switching_mode,
                "available_modes": [
                    {"mode": "SA", "label": "SA 频谱测试"},
                    {"mode": "NA", "label": "NA 天线测量"},
                ],
            }

    def switch_mode(self, mode):
        mode = str(mode or "").upper()
        if mode not in {"SA", "NA"}:
            raise ServiceError("模式只能是 SA 或 NA。")

        with self.lock:
            if self.measurement_in_progress:
                raise ServiceError("已有测量或校准正在进行，不能切换模式。")
            if self.current_mode == mode and not self.switching_mode:
                return self.mode_status()
            self.switching_mode = True
            self.progress_message = f"正在切换到 {mode} 模式..."
            self.last_error = None

        try:
            if not self.demo_mode and (self.controller.connected or self.na_controller.connected):
                if mode == "NA":
                    self._select_na_hardware_mode()
                else:
                    self._select_sa_hardware_mode()

            with self.lock:
                self.current_mode = mode
                if mode == "NA":
                    self._clear_sa_results_locked()
                else:
                    self._clear_na_results_locked()
                self.progress_message = f"已切换到 {mode} 模式"
        except Exception as exc:
            with self.lock:
                self.last_error = str(exc)
                self.progress_message = "模式切换失败"
            raise
        finally:
            with self.lock:
                self.switching_mode = False
        return self.mode_status()

    def _select_na_hardware_mode(self):
        self.na_controller.ip_address = self.controller.ip_address
        if not self.na_controller.connected:
            if not self.na_controller.connect():
                raise ServiceError("无法连接 N9918A NA 控制器，模式切换失败。")
        else:
            self.na_controller.select_mode()

    def _select_sa_hardware_mode(self):
        if not self.controller.connected:
            self.controller.ip_address = self.na_controller.ip_address or self.controller.ip_address
            if not self.controller.connect():
                raise ServiceError("无法连接 N9918A SA 控制器，模式切换失败。")
            return
        if self.controller.device:
            self.controller.device.query("INST:SEL 'SA';*OPC?")

    def _clear_sa_results_locked(self):
        self.current_frequencies = None
        self.current_amplitudes = None
        self.current_peaks = None
        self.emi_results = {}
        self.last_ai_result = ""

    def _clear_na_results_locked(self):
        self.na_result = None
        self.na_last_report_path = None

    def update_user_info(self, data):
        with self.lock:
            for key in self.user_info:
                if key in data:
                    self.user_info[key] = str(data[key]).strip()
            return self.user_info.copy()

    def connect_device(self, ip_address):
        with self.lock:
            self.controller.ip_address = ip_address.strip() or self.controller.ip_address
            self.na_controller.ip_address = self.controller.ip_address
            self.progress_message = "正在连接仪器..."
            self.last_error = None

        ok = self.controller.connect()
        with self.lock:
            self.demo_mode = False
            self.current_mode = "SA"
            self.progress_message = "仪器已连接" if ok else "仪器连接失败"
            if not ok:
                self.last_error = "无法连接 N9918A"
        if not ok:
            raise ServiceError("无法连接 N9918A，请检查 IP、网络、VISA backend 和仪器状态。")
        return self.status()

    def disconnect_device(self):
        with self.lock:
            self.controller.disconnect()
            self.na_controller.disconnect()
            self.current_frequencies = None
            self.current_amplitudes = None
            self.current_peaks = None
            self.emi_results = {}
            self.last_ai_result = ""
            self.last_report_path = None
            self.na_config_key = None
            self.na_config = None
            self.na_result = None
            self.na_calibration = self._empty_na_calibration()
            self.na_last_report_path = None
            self.current_mode = "SA"
            self.switching_mode = False
            self.progress_message = "仪器已断开"
            self.demo_mode = False
        return self.status()

    def connect_switch(self):
        if not self.switch_controller:
            raise ServiceError(f"切换器模块不可用: {SWITCH_IMPORT_ERROR}")
        ok = self.switch_controller.connect()
        if not ok:
            raise ServiceError("无法连接 Mini-Circuits RF Switch，请检查 USB、驱动和 DLL。")
        return self.switch_status()

    def disconnect_switch(self):
        if self.switch_controller:
            self.switch_controller.disconnect()
        return {"connected": False}

    def switch_status(self):
        if not self.switch_controller:
            raise ServiceError(f"切换器模块不可用: {SWITCH_IMPORT_ERROR}")
        if not self.switch_controller.connected:
            return {"connected": False}
        return {
            "connected": True,
            "positions": self.switch_controller.get_switch_status(),
            "model": self.switch_controller.get_model_name(),
            "serial": self.switch_controller.get_serial_number(),
            "firmware": self.switch_controller.get_firmware(),
            "temperature": self.switch_controller.get_temperature(),
            "usb_status": self.switch_controller.get_usb_status(),
        }

    def set_switch_position(self, switch_name, position):
        if self.measurement_in_progress and self.measurement_kind == "NA 自动校准":
            raise ServiceError("NA 自动校准期间不能手动切换 switchbox。")
        if not self.switch_controller or not self.switch_controller.connected:
            raise ServiceError("切换器未连接。")
        self.switch_controller.set_switch(str(switch_name).upper(), int(position))
        return self.switch_status()

    def configure(self, preset_key):
        if not self.controller.connected:
            raise ServiceError("请先连接 N9918A。")
        if self.current_mode != "SA":
            self.switch_mode("SA")
        with self.lock:
            self.progress_message = "正在配置仪器..."
            self.last_error = None

        if self.demo_mode:
            self._apply_preset_fields(preset_key)
            with self.lock:
                self.current_frequencies = None
                self.current_amplitudes = None
                self.current_peaks = None
                self.emi_results = {}
                self.last_ai_result = ""
                self.progress_message = "演示配置完成"
            return self.status()

        ok = self.controller.configure_settings(preset_key)
        if not ok:
            with self.lock:
                self.progress_message = "配置失败"
                self.last_error = f"配置失败: {preset_key}"
            raise ServiceError("配置仪器失败，请检查 SCPI 支持、仪器模式和连接状态。")

        warnings = self._auto_set_switch_positions()
        with self.lock:
            self.current_frequencies = None
            self.current_amplitudes = None
            self.current_peaks = None
            self.emi_results = {}
            self.last_ai_result = ""
            self.progress_message = "配置完成"
        result = self.status()
        result["warnings"] = warnings
        return result

    def _apply_preset_fields(self, preset_key):
        config = self.controller.get_preset_configs().get(preset_key)
        if not config:
            raise ServiceError(f"未知测试配置: {preset_key}")
        self.controller.current_config = preset_key
        self.controller.start_freq = config["start_freq"]
        self.controller.stop_freq = config["stop_freq"]
        self.controller.n_points = config["n_points"]
        self.controller.rbw = config["rbw"]
        self.controller.vbw = config["vbw"]

    def _auto_set_switch_positions(self):
        if not self.switch_controller or not self.switch_controller.connected:
            return ["切换器未连接，已跳过自动切换。"]

        config = self.controller.get_preset_configs().get(self.controller.current_config, {})
        start_freq = config.get("start_freq", 0)
        freq_mhz = start_freq / 1e6

        try:
            if freq_mhz < 30:
                positions = {"A": 2, "D": 2, "B": 1, "C": 1}
            elif 30 <= freq_mhz <= 3000:
                positions = {"A": 2, "D": 1, "B": 1, "C": 1}
            else:
                positions = {"A": 1, "B": 1, "C": 1, "D": 1}
            for switch, position in positions.items():
                self.switch_controller.set_switch(switch, position)
        except Exception as exc:
            return [f"自动切换失败: {exc}"]
        return []

    def start_single_measurement(self):
        self._start_measurement_thread("单次扫描", self._run_single_measurement)
        return self.status()

    def start_emi_measurement(self, duration_seconds):
        duration_seconds = int(duration_seconds)
        if duration_seconds <= 0:
            raise ServiceError("测量时长必须大于 0。")
        self._start_measurement_thread(
            f"EMI {duration_seconds} 秒采样",
            lambda: self._run_emi_measurement(duration_seconds),
        )
        return self.status()

    def _start_measurement_thread(self, kind, target):
        with self.lock:
            if self.measurement_in_progress:
                raise ServiceError("已有测量正在进行，请先停止或等待完成。")
            if self.current_mode != "SA":
                raise ServiceError("当前不是 SA 模式，请先切换到 SA 频谱测试。")
            if not self.controller.connected:
                raise ServiceError("请先连接 N9918A。")
            if not self.controller.current_config:
                raise ServiceError("请先选择并应用 Test Config。")

            self.stop_event.clear()
            self.measurement_in_progress = True
            self.measurement_kind = kind
            self.progress_message = f"{kind} 运行中"
            self.last_error = None
            self.measurement_thread = threading.Thread(target=target, daemon=True)
            self.measurement_thread.start()

    def load_demo_data(self, preset_key="EMC_30MHz_1GHz", duration_seconds=15):
        """Populate realistic sample data without touching hardware."""
        with self.lock:
            self.demo_mode = True
            self.controller.connected = True
            self.current_mode = "SA"
            self._apply_preset_fields(preset_key)
            self.progress_message = "演示数据已加载"
            self.last_error = None
            self.measurement_in_progress = False
            self.measurement_kind = None
            self.na_result = None

        results = self._generate_demo_results(duration_seconds)
        frequencies, amplitudes = results["QUASI_PEAK"]
        peaks = post_process_peak_search(frequencies, amplitudes)
        with self.lock:
            self.emi_results = results
            self.current_frequencies = frequencies
            self.current_amplitudes = amplitudes
            self.current_peaks = peaks
            self.last_ai_result = (
                "Demo AI 分析：175 MHz 与 275 MHz 附近出现高风险峰值，"
                "疑似 25 MHz 基准时钟谐波或线缆耦合路径引入。请在真实硬件上复测后再用于整改决策。"
            )
        return self.result_payload()

    def _run_single_measurement(self):
        if self.demo_mode:
            try:
                results = self._generate_demo_results(1)
                frequencies, amplitudes = results["PEAK"]
                peaks = post_process_peak_search(frequencies, amplitudes)
                with self.lock:
                    self.current_frequencies = frequencies
                    self.current_amplitudes = amplitudes
                    self.current_peaks = peaks
                    self.emi_results = {}
                    self.last_ai_result = ""
                    self.progress_message = "演示单次扫描完成"
            finally:
                with self.lock:
                    self.measurement_in_progress = False
                    self.measurement_kind = None
            return

        try:
            frequencies, amplitudes = self.controller.read_trace_data()
            if not frequencies or not amplitudes:
                raise ServiceError("单次测量未返回有效数据。")
            peaks = post_process_peak_search(frequencies, amplitudes)
            with self.lock:
                self.current_frequencies = frequencies
                self.current_amplitudes = amplitudes
                self.current_peaks = peaks
                self.emi_results = {}
                self.last_ai_result = ""
                self.progress_message = "单次扫描完成"
        except Exception as exc:
            with self.lock:
                self.last_error = str(exc)
                self.progress_message = "单次扫描失败"
        finally:
            with self.lock:
                self.measurement_in_progress = False
                self.measurement_kind = None

    def _run_emi_measurement(self, duration_seconds):
        if self.demo_mode:
            try:
                time.sleep(0.4)
                results = self._generate_demo_results(duration_seconds)
                frequencies, amplitudes = results["QUASI_PEAK"]
                peaks = post_process_peak_search(frequencies, amplitudes)
                with self.lock:
                    self.emi_results = results
                    self.current_frequencies = frequencies
                    self.current_amplitudes = amplitudes
                    self.current_peaks = peaks
                    self.last_ai_result = ""
                    self.progress_message = f"演示 EMI {duration_seconds} 秒采样完成"
            finally:
                with self.lock:
                    self.measurement_in_progress = False
                    self.measurement_kind = None
            return

        try:
            results = self.controller.get_emc_measurement_fast(
                duration_seconds,
                should_stop=self.stop_event.is_set,
            )
            if not results:
                raise ServiceError("EMI 测量未返回有效数据。")

            display_mode = "QUASI_PEAK" if "QUASI_PEAK" in results else "PEAK"
            frequencies, amplitudes = results[display_mode]
            peaks = post_process_peak_search(frequencies, amplitudes)
            with self.lock:
                self.emi_results = results
                self.current_frequencies = frequencies
                self.current_amplitudes = amplitudes
                self.current_peaks = peaks
                self.last_ai_result = ""
                self.progress_message = f"EMI {duration_seconds} 秒采样完成"
        except Exception as exc:
            with self.lock:
                self.last_error = str(exc)
                self.progress_message = "EMI 采样失败"
        finally:
            with self.lock:
                self.measurement_in_progress = False
                self.measurement_kind = None

    def _generate_demo_results(self, duration_seconds):
        n_points = self.controller.n_points or 1001
        start_freq = self.controller.start_freq or 30e6
        stop_freq = self.controller.stop_freq or 1e9
        step = (stop_freq - start_freq) / (n_points - 1)
        frequencies = [start_freq + i * step for i in range(n_points)]
        rng = random.Random(9918 + int(duration_seconds) + int(start_freq))

        samples = []
        sample_count = max(4, min(24, int(duration_seconds / 0.8) if duration_seconds > 1 else 4))
        for sample_index in range(sample_count):
            amplitudes = []
            for freq in frequencies:
                mhz = freq / 1e6
                log_pos = math.log10(max(mhz, 0.01))
                baseline = 18 + 7 * math.sin(log_pos * 3.7) + 2.5 * math.sin(mhz / 41.0)
                noise = rng.uniform(-1.2, 1.2)
                value = baseline + noise
                for center, height, width in [
                    (47.0, 19, 1.8),
                    (175.0, 27, 2.8),
                    (225.0, 17, 3.2),
                    (275.0, 30, 3.6),
                    (500.0, 14, 5.5),
                ]:
                    value += height * math.exp(-((mhz - center) ** 2) / (2 * width ** 2))
                amplitudes.append(round(max(value, 0), 3))
            samples.append(
                {
                    "timestamp": sample_index * max(duration_seconds / sample_count, 0.2),
                    "frequencies": frequencies,
                    "amplitudes": amplitudes,
                }
            )

        peak = [max(sample["amplitudes"][i] for sample in samples) for i in range(n_points)]
        average = [
            sum(sample["amplitudes"][i] for sample in samples) / len(samples)
            for i in range(n_points)
        ]
        quasi_peak = [max(avg * 0.74 + pk * 0.26, avg + 1.5) for avg, pk in zip(average, peak)]

        return {
            "PEAK": (frequencies, peak),
            "QUASI_PEAK": (frequencies, quasi_peak),
            "AVERAGE": (frequencies, average),
            "sampling_data": samples,
            "sampling_info": {
                "total_samples": len(samples),
                "sample_duration": duration_seconds,
                "collection_time": duration_seconds,
                "calculation_time": 0.05,
                "data_points": n_points,
                "rbw": self.controller.rbw,
                "start_time": samples[0]["timestamp"],
                "end_time": samples[-1]["timestamp"],
            },
            "measurement_summary": {
                "total_duration": duration_seconds,
                "actual_measurement_time": duration_seconds,
                "data_points": n_points,
                "total_samples": len(samples),
                "modes_computed": ["PEAK", "QUASI_PEAK", "AVERAGE"],
                "measurement_time": time.strftime("%Y-%m-%d %H:%M:%S"),
                "demo": True,
            },
        }

    def na_presets(self):
        return self.na_controller.get_preset_configs()

    def na_configure(self, preset_key, points=None, ifbw=None):
        if preset_key not in NA_PRESET_CONFIGS:
            raise ServiceError(f"未知 NA 天线预设: {preset_key}")
        with self.lock:
            if self.measurement_in_progress:
                raise ServiceError("已有测量或校准正在进行，请先等待完成。")
            self.progress_message = "正在配置 NA 天线测量..."
            self.last_error = None

        if self.demo_mode:
            config = self._apply_na_preset_fields(preset_key, points=points, ifbw=ifbw)
        else:
            if not (self.controller.connected or self.na_controller.connected):
                raise ServiceError("请先连接 N9918A。")
            if self.current_mode != "NA":
                self.switch_mode("NA")
            try:
                config = self.na_controller.configure_preset(preset_key, points=points, ifbw=ifbw)
            except N9918ANAError as exc:
                raise ServiceError(str(exc)) from exc

        with self.lock:
            self.current_mode = "NA"
            self.na_config_key = preset_key
            self.na_config = config
            self.na_result = None
            self.na_calibration = self._empty_na_calibration()
            self.progress_message = "NA 配置完成"
        return self.na_result_payload()

    def _apply_na_preset_fields(self, preset_key, points=None, ifbw=None):
        config = dict(NA_PRESET_CONFIGS[preset_key])
        if points:
            config["points"] = int(points)
        if ifbw:
            config["ifbw"] = float(ifbw)
        self.na_controller.current_preset_key = preset_key
        self.na_controller.current_config = config
        self.na_controller.start_freq = config["start_freq"]
        self.na_controller.stop_freq = config["stop_freq"]
        self.na_controller.points = config["points"]
        self.na_controller.ifbw = config["ifbw"]
        return config

    def na_calibrate(self):
        with self.lock:
            if self.measurement_in_progress:
                raise ServiceError("已有测量或校准正在进行，请先等待完成。")
            if self.current_mode != "NA":
                raise ServiceError("请先切换到 NA 天线测量模式。")
            if not self.na_config:
                raise ServiceError("请先选择并应用 NA 天线预设。")
            self.stop_event.clear()
            self.measurement_in_progress = True
            self.measurement_kind = "NA 自动校准"
            self.na_calibration = self._empty_na_calibration()
            self.na_calibration["in_progress"] = True
            self.progress_message = "NA 自动校准运行中"
            self.last_error = None

        if self.demo_mode:
            try:
                time.sleep(0.1)
                events = [
                    {"step": "OPEN", "label": "OPEN 校准", "switch_position": "B2C1", "scpi": "CORR:COLL:INT 1;*OPC?", "ok": True, "message": "演示"},
                    {"step": "LOAD", "label": "LOAD 校准", "switch_position": "B1C1", "scpi": "CORR:COLL:LOAD 1;*OPC?", "ok": True, "message": "演示"},
                    {"step": "SAVE", "label": "保存校准", "switch_position": "B1C1", "scpi": "CORR:COLL:SAVE 0", "ok": True, "message": "演示"},
                    {"step": "ANTENNA", "label": "切到天线测量", "switch_position": "B2C2", "scpi": None, "ok": True, "message": "演示"},
                ]
                with self.lock:
                    self.na_calibration = {
                        "complete": True,
                        "in_progress": False,
                        "events": events,
                        "error": None,
                        "last_scpi": "CORR:COLL:SAVE 0",
                        "switch_position": "B2C2",
                    }
                    self.progress_message = "NA 演示校准完成"
            finally:
                with self.lock:
                    self.measurement_in_progress = False
                    self.measurement_kind = None
            return self.na_result_payload()

        if not (self.na_controller.connected and self.na_controller.device):
            with self.lock:
                self.measurement_in_progress = False
                self.measurement_kind = None
                self.na_calibration["in_progress"] = False
            raise ServiceError("请先连接 N9918A 并切换到 NA 模式。")
        if not self.switch_controller or not self.switch_controller.connected:
            with self.lock:
                self.measurement_in_progress = False
                self.measurement_kind = None
                self.na_calibration["in_progress"] = False
            raise ServiceError("NA 校准需要 N9918A 和 switchbox 都已连接。")

        def record_progress(event):
            with self.lock:
                self.na_calibration["events"].append(event)
                self.na_calibration["last_scpi"] = event.get("scpi")
                self.na_calibration["switch_position"] = event.get("switch_position")
                self.progress_message = event.get("label") or "NA 自动校准运行中"

        try:
            result = self.na_controller.perform_calibration(
                self.switch_controller,
                progress_callback=record_progress,
                should_stop=self.stop_event.is_set,
            )
            with self.lock:
                self.na_calibration.update(result)
                self.na_calibration["complete"] = True
                self.na_calibration["in_progress"] = False
                self.na_calibration["error"] = None
                self.progress_message = "NA 自动校准完成"
        except N9918ANAError as exc:
            with self.lock:
                self.na_calibration["complete"] = False
                self.na_calibration["in_progress"] = False
                self.na_calibration["error"] = exc.as_dict()
                self.last_error = str(exc)
                self.progress_message = "NA 自动校准失败"
            raise ServiceError(f"NA 自动校准失败: {exc}") from exc
        finally:
            with self.lock:
                self.measurement_in_progress = False
                self.measurement_kind = None
        return self.na_result_payload()

    def start_na_measurement(self):
        with self.lock:
            if self.measurement_in_progress:
                raise ServiceError("已有测量或校准正在进行，请先停止或等待完成。")
            if self.current_mode != "NA":
                raise ServiceError("请先切换到 NA 天线测量模式。")
            if not self.na_config:
                raise ServiceError("请先选择并应用 NA 天线预设。")
            if not self.na_calibration.get("complete"):
                raise ServiceError("请先完成 NA 自动校准。")
            if not self.demo_mode and not self.na_controller.connected:
                raise ServiceError("请先连接 N9918A。")

            self.stop_event.clear()
            self.measurement_in_progress = True
            self.measurement_kind = "NA 天线测量"
            self.progress_message = "NA 天线测量运行中"
            self.last_error = None
            self.measurement_thread = threading.Thread(target=self._run_na_measurement, daemon=True)
            self.measurement_thread.start()
        return self.na_result_payload()

    def _run_na_measurement(self):
        try:
            if self.demo_mode:
                time.sleep(0.25)
                result = self._generate_na_demo_result()
            else:
                result = self.na_controller.measure_s11(should_stop=self.stop_event.is_set)
            with self.lock:
                self.na_result = result
                self.progress_message = "NA 天线测量完成"
        except Exception as exc:
            with self.lock:
                self.last_error = str(exc)
                self.progress_message = "NA 天线测量失败"
        finally:
            with self.lock:
                self.measurement_in_progress = False
                self.measurement_kind = None

    def _generate_na_demo_result(self):
        config = dict(self.na_config or NA_PRESET_CONFIGS["ANT_433"])
        points = int(config.get("points") or 2001)
        if config.get("full_sweep"):
            points = min(points, 5001)
        frequencies = frequency_axis(config["start_freq"], config["stop_freq"], points)
        s11_db = []
        centers = self._demo_na_centers(config)
        for freq in frequencies:
            mhz = freq / 1e6
            baseline = -2.2 + 0.8 * math.sin(math.log10(max(mhz, 0.001)) * 3.1)
            value = baseline
            for center_mhz, depth_db, width_mhz in centers:
                value -= depth_db * math.exp(-((mhz - center_mhz) ** 2) / (2 * width_mhz ** 2))
            s11_db.append(round(value, 4))

        real = []
        imag = []
        span = max(frequencies[-1] - frequencies[0], 1)
        for freq, db in zip(frequencies, s11_db):
            mag = min(0.98, max(0.02, 10 ** (db / 20)))
            phase = -math.pi + 2 * math.pi * ((freq - frequencies[0]) / span)
            real.append(mag * math.cos(phase))
            imag.append(mag * math.sin(phase))

        return build_na_result(frequencies, s11_db, real, imag, config, self.na_config_key)

    @staticmethod
    def _demo_na_centers(config):
        label = config.get("label", "")
        if config.get("full_sweep"):
            return [
                (433.0, 18, 8.5),
                (898.0, 16, 10.0),
                (915.0, 19, 8.0),
                (2450.0, 22, 30.0),
                (5200.0, 15, 55.0),
            ]
        if "433" in label:
            return [(433.0, 23, 10.0), (470.0, 7, 5.5)]
        if "898" in label:
            return [(898.0, 21, 9.5), (930.0, 8, 6.0)]
        if "915" in label:
            return [(915.0, 24, 8.0), (880.0, 7, 6.0)]
        if "2450" in label:
            return [(2450.0, 25, 28.0), (2412.0, 9, 12.0)]
        if "5G" in label or "5GHz" in label:
            return [(5200.0, 20, 60.0), (5750.0, 16, 45.0)]
        center = (config["start_freq"] + config["stop_freq"]) / 2 / 1e6
        return [(center, 20, max((config["stop_freq"] - config["start_freq"]) / 1e6 / 30, 1.0))]

    def stop_na_measurement(self):
        self.stop_event.set()
        if self.na_controller.connected and self.na_controller.device:
            try:
                self.na_controller.device.write("INIT:CONT 0")
            except Exception as exc:
                with self.lock:
                    self.last_error = str(exc)
        with self.lock:
            self.progress_message = "已请求停止 NA 测量/校准"
        return self.na_result_payload()

    def na_result_payload(self):
        with self.lock:
            result = self.na_result or {}
            status = {
                "current_mode": self.current_mode,
                "connected": self.controller.connected or self.na_controller.connected,
                "configured": bool(self.na_config),
                "config": self.na_config,
                "preset_key": self.na_config_key,
                "calibration": self.na_calibration.copy(),
                "switch_position": self.na_calibration.get("switch_position"),
                "measurement_in_progress": self.measurement_in_progress,
                "measurement_kind": self.measurement_kind,
                "progress_message": self.progress_message,
                "error": self.last_error,
                "last_report": str(self.na_last_report_path) if self.na_last_report_path else None,
                "demo_mode": self.demo_mode,
            }
            return {
                "status": status,
                "series": result.get("series"),
                "smith": result.get("smith"),
                "primary_valley": result.get("primary_valley"),
                "bandwidths": result.get("bandwidths") or {},
                "valleys": result.get("valleys") or [],
                "config": result.get("config") or self.na_config,
                "is_full_sweep": bool(result.get("is_full_sweep") or (self.na_config or {}).get("full_sweep")),
            }

    def save_na_data(self):
        with self.lock:
            result = self.na_result
        if not result:
            raise ServiceError("没有可保存的 NA 测量数据。")
        return save_na_measurement_data(result)

    def export_na_report(self, user_info=None):
        with self.lock:
            if user_info:
                self.update_user_info(user_info)
            result = self.na_result
            project_info = self.user_info.copy()
        if not result:
            raise ServiceError("没有可导出的 NA 测量结果。")
        report_path = export_na_report(result, user_info=project_info)
        with self.lock:
            self.na_last_report_path = report_path
        return report_path

    def stop_measurement(self):
        self.stop_event.set()
        if self.controller.connected and self.controller.device:
            try:
                self.controller.device.write("INIT:CONT OFF")
            except Exception as exc:
                with self.lock:
                    self.last_error = str(exc)
        with self.lock:
            self.progress_message = "已请求停止"
        return self.status()

    def result_payload(self):
        with self.lock:
            return {
                "status": self.status(),
                "series": self._series_payload_locked(),
                "modes": self._modes_payload_locked(),
                "peaks": self.current_peaks or [],
                "peak_table": self.format_peak_table(),
                "measurement_summary": self.emi_results.get("measurement_summary", {}),
                "sampling_info": self.emi_results.get("sampling_info", {}),
                "ai_result": self.last_ai_result,
            }

    def _series_payload_locked(self):
        if not self.current_frequencies or not self.current_amplitudes:
            return None
        return {
            "frequency_mhz": [round(freq / 1e6, 6) for freq in self.current_frequencies],
            "amplitude_dbuv": [round(value, 3) for value in self.current_amplitudes],
        }

    def _modes_payload_locked(self):
        modes = {}
        for mode in ("PEAK", "QUASI_PEAK", "AVERAGE"):
            data = self.emi_results.get(mode)
            if isinstance(data, tuple) and len(data) >= 2:
                frequencies, amplitudes = data[:2]
                modes[mode] = {
                    "frequency_mhz": [round(freq / 1e6, 6) for freq in frequencies],
                    "amplitude_dbuv": [round(value, 3) for value in amplitudes],
                }
        return modes

    def format_peak_table(self):
        peaks = self.current_peaks or []
        if not peaks:
            return "暂无峰值数据。"

        lines = [
            "No   频率 [MHz]   幅度 [dBμV]   FCC限值 [dBμV]   FCC裕量 [dB]    CE限值 [dBμV]    CE裕量 [dB]     状态",
            "-" * 128,
        ]
        for index, peak in enumerate(peaks, start=1):
            status = []
            if peak["exceed_fcc"]:
                status.append("FCC 超限")
            if peak["exceed_ce"]:
                status.append("CE 超限")
            if not status:
                status.append("通过")
            lines.append(
                f"{index:<4} {peak['frequency_mhz']:<12.3f} "
                f"{peak['amplitude_dbuv']:<18.2f} {peak['fcc_limit']:<18.1f} "
                f"{peak['fcc_margin']:<18.2f} {peak['ce_limit']:<18.1f} "
                f"{peak['ce_margin']:<18.2f} {', '.join(status):<15}"
            )
        return "\n".join(lines)

    def build_ai_analysis_input(self):
        if not self.current_peaks:
            raise ServiceError("没有可分析的峰值数据。")

        start_freq = self.controller.start_freq or 0
        stop_freq = self.controller.stop_freq or 0
        duration = 0
        if "measurement_summary" in self.emi_results:
            duration = self.emi_results["measurement_summary"].get("actual_measurement_time", 0)

        return (
            f"频段:{start_freq/1e6:.3f}MHz-{stop_freq/1e6:.3f}MHz 测量时长：{duration}s 测量数据：\n"
            "QUASI_PEAK Mode Results:\n"
            f"{'=' * 100}\n"
            f"{self.format_peak_table()}\n"
        )

    def analyze(self):
        input_text = self.build_ai_analysis_input()
        from chat import ChatBot, sys_prompt

        bot = ChatBot(system_message=sys_prompt)
        response = bot.chat_no_stream(input_text)
        msg_obj = response.choices[0].message
        result = msg_obj.content if hasattr(msg_obj, "content") else msg_obj.get("content", "")
        with self.lock:
            self.last_ai_result = result
        return result

    def save_data(self):
        saved_files = []
        with self.lock:
            emi_results = self.emi_results
            frequencies = self.current_frequencies
            amplitudes = self.current_amplitudes
            peaks = self.current_peaks

        if emi_results:
            saved_files.extend(save_emi_measurement_data(emi_results))
            if peaks:
                saved_files.append(save_peak_analysis(peaks))
        elif frequencies and amplitudes:
            saved_files.append(save_spectrum_data(frequencies, amplitudes))
            if peaks:
                saved_files.append(save_peak_analysis(peaks))
        else:
            raise ServiceError("没有可保存的数据。")

        return saved_files

    def export_pdf(self, user_info=None, auto_analyze=True):
        with self.lock:
            if user_info:
                self.update_user_info(user_info)
            if not self.emi_results:
                raise ServiceError("PDF 报告仅支持 15s/5min 等 EMI 测量结果。")

        if auto_analyze and not self.last_ai_result:
            self.analyze()

        graph_path = self._render_graph_png()
        reports_dir = ROOT / "reports"
        reports_dir.mkdir(exist_ok=True)

        with self.lock:
            summary = self.emi_results.get("measurement_summary", {})
            duration = summary.get("actual_measurement_time", 0)
            start_freq = self.controller.start_freq or 0
            stop_freq = self.controller.stop_freq or 0
            mode = f"{start_freq/1e6:.3f}MHz-{stop_freq/1e6:.3f}MHz_{duration}s"
            project_info = self.user_info.copy()
            project_info["mode"] = mode
            filename = f"{project_info.get('eut', 'N9918A')}-{mode}.pdf"
            filename = self._safe_filename(filename)
            output_path = reports_dir / filename
            summary_text = self.last_ai_result
            peak_table = self.format_peak_table()

        try:
            try:
                from utils.create_pdf import generate_test_report
            except ImportError as exc:
                raise ServiceError("缺少 PDF 依赖，请先运行 `pip install -r requirements.txt`。") from exc

            generate_test_report(
                filename=str(output_path),
                logo_path="./assets/m5logo2022.png",
                project_info=project_info,
                test_graph_path=str(graph_path),
                spectrum_data=peak_table,
                summary_text=summary_text,
            )
        finally:
            try:
                os.remove(graph_path)
            except OSError:
                pass

        with self.lock:
            self.last_report_path = output_path
        return output_path

    def _render_graph_png(self):
        with self.lock:
            frequencies = list(self.current_frequencies or [])
            amplitudes = list(self.current_amplitudes or [])
            peaks = list(self.current_peaks or [])

        if not frequencies or not amplitudes:
            raise ServiceError("没有可绘制的频谱数据。")

        temp = tempfile.NamedTemporaryFile(prefix="n9918a_report_", suffix=".png", delete=False)
        temp.close()

        freq_mhz = [freq / 1e6 for freq in frequencies]
        fig, ax = plt.subplots(figsize=(12, 7))
        ax.semilogx(freq_mhz, amplitudes, color="#175c7f", linewidth=1.2, label="测量曲线")

        fcc_limits = []
        ce_limits = []
        for freq in frequencies:
            fcc_limit, ce_limit = get_fcc_ce_limits(freq)
            fcc_limits.append(fcc_limit)
            ce_limits.append(ce_limit)
        ax.semilogx(freq_mhz, fcc_limits, color="#d95032", linewidth=1, linestyle="--", label="FCC Class B")
        ax.semilogx(freq_mhz, ce_limits, color="#287a3e", linewidth=1, linestyle="--", label="CE Class B")

        for peak in peaks[:15]:
            color = "#d95032" if peak["exceed_fcc"] or peak["exceed_ce"] else "#222222"
            ax.plot(peak["frequency_mhz"], peak["amplitude_dbuv"], "o", color=color, markersize=4)

        ax.set_title("N9918A EMC Spectrum")
        ax.set_xlabel("Frequency (MHz)")
        ax.set_ylabel("Amplitude (dBμV)")
        ax.grid(True, which="both", alpha=0.25)
        ax.legend(loc="upper right")
        fig.tight_layout()
        fig.savefig(temp.name, dpi=160)
        plt.close(fig)
        return temp.name

    @staticmethod
    def _safe_filename(filename):
        filename = re.sub(r'[<>:"/\\\\|?*]+', "_", filename)
        return filename.strip().strip(".") or f"report_{int(time.time())}.pdf"
