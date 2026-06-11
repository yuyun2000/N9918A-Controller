
# n9918a_na_backend.py
"""N9918A NA/S11 antenna measurement controller and analysis helpers."""

from __future__ import annotations

import csv
import html
import json
import math
import os
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable, List, Optional, Sequence, Tuple

try:
    import pyvisa
except ImportError:  # pragma: no cover - depends on local hardware environment.
    class _PyVisaStub:
        class errors:
            VisaIOError = Exception

        @staticmethod
        def ResourceManager(*_args, **_kwargs):
            raise RuntimeError("pyvisa is not installed. Run `pip install -r requirements.txt`.")

    pyvisa = _PyVisaStub()


NA_PRESET_CONFIGS = {
    "ANT_433": {
        "name": "433MHz 天线",
        "label": "433MHz",
        "start_freq": 300e6,
        "stop_freq": 500e6,
        "points": 2001,
        "ifbw": 10e3,
        "full_sweep": False,
        "description": "300-500MHz，搜索 433MHz 附近 S11 谷值与带宽。",
    },
    "ANT_898": {
        "name": "898MHz 天线",
        "label": "898MHz",
        "start_freq": 798e6,
        "stop_freq": 998e6,
        "points": 2001,
        "ifbw": 10e3,
        "full_sweep": False,
        "description": "798-998MHz，适合 898MHz 天线调试。",
    },
    "ANT_915": {
        "name": "915MHz 天线",
        "label": "915MHz",
        "start_freq": 815e6,
        "stop_freq": 1015e6,
        "points": 2001,
        "ifbw": 10e3,
        "full_sweep": False,
        "description": "815-1015MHz，适合 915MHz 天线调试。",
    },
    "ANT_2450": {
        "name": "2450MHz 天线",
        "label": "2450MHz",
        "start_freq": 2.2e9,
        "stop_freq": 2.7e9,
        "points": 2001,
        "ifbw": 10e3,
        "full_sweep": False,
        "description": "2.2-2.7GHz，覆盖 2.4GHz ISM 频段。",
    },
    "ANT_5G": {
        "name": "5GHz 天线",
        "label": "5GHz",
        "start_freq": 4.8e9,
        "stop_freq": 6.0e9,
        "points": 2001,
        "ifbw": 10e3,
        "full_sweep": False,
        "description": "4.8-6.0GHz，覆盖常见 5GHz Wi-Fi 天线调试范围。",
    },
    "ANT_FULL": {
        "name": "全扫宽",
        "label": "全扫宽",
        "start_freq": 30e3,
        "stop_freq": 26.5e9,
        "points": 5001,
        "ifbw": 10e3,
        "full_sweep": True,
        "description": "30kHz-26.5GHz，列出全范围所有 S11 局部谷值；不生成 Smith Chart。",
    },
}

SWITCH_POSITIONS = {
    "LOAD": {"B": 1, "D": 1},
    "OPEN": {"B": 2, "D": 1},
    "ANTENNA": {"B": 2, "D": 2},
}


@dataclass
class CalibrationEvent:
    step: str
    label: str
    switch_position: str
    scpi: Optional[str]
    ok: bool = True
    message: str = ""

    def as_dict(self):
        return {
            "step": self.step,
            "label": self.label,
            "switch_position": self.switch_position,
            "scpi": self.scpi,
            "ok": self.ok,
            "message": self.message,
        }


class N9918ANAError(RuntimeError):
    """NA workflow error with hardware context for the UI."""

    def __init__(self, message, step=None, last_scpi=None, switch_position=None):
        super().__init__(message)
        self.step = step
        self.last_scpi = last_scpi
        self.switch_position = switch_position

    def as_dict(self):
        return {
            "message": str(self),
            "step": self.step,
            "last_scpi": self.last_scpi,
            "switch_position": self.switch_position,
        }


class N9918ANAController:
    """PyVISA controller for FieldFox NA/S11 antenna measurements."""

    def __init__(self, ip_address="192.168.20.233", timeout=30000):
        self.ip_address = ip_address
        self.timeout = timeout
        self.rm = None
        self.device = None
        self.connected = False
        self.last_scpi = None
        self.current_preset_key = None
        self.current_config = None
        self.start_freq = None
        self.stop_freq = None
        self.points = None
        self.ifbw = None
        self.last_switch_position = None

    def connect(self):
        try:
            self.rm = pyvisa.ResourceManager()
            self.device = self.rm.open_resource(f"TCPIP0::{self.ip_address}::inst0::INSTR")
            self.device.timeout = self.timeout
            self._write("*CLS")
            device_id = self._query("*IDN?")
            self.connected = True
            self.select_mode()
            print(f"Connected to N9918A NA: {device_id}")
            return True
        except Exception as exc:
            print(f"ERROR: Unable to connect to N9918A NA - {exc}")
            self.connected = False
            return False

    def disconnect(self):
        if self.device:
            try:
                self.device.close()
            except Exception:
                pass
        if self.rm:
            try:
                self.rm.close()
            except Exception:
                pass
        self.device = None
        self.rm = None
        self.connected = False

    def select_mode(self):
        self._require_connected()
        return self._query('INST:SEL "NA";*OPC?')

    def get_preset_configs(self):
        return NA_PRESET_CONFIGS

    def configure_preset(self, preset_key, points=None, ifbw=None):
        self._require_connected()
        if preset_key not in NA_PRESET_CONFIGS:
            raise N9918ANAError(f"Unknown NA preset: {preset_key}")
        config = dict(NA_PRESET_CONFIGS[preset_key])
        if points:
            config["points"] = int(points)
        if ifbw:
            config["ifbw"] = float(ifbw)

        self.select_mode()
        self._write("CALC:PAR:DEF S11")
        self._write("CALC:FORM MLOG")
        self._write(f"SENS:FREQ:START {config['start_freq']}")
        self._write(f"SENS:FREQ:STOP {config['stop_freq']}")
        self._write(f"SENS:SWE:POIN {config['points']}")
        self._write(f"BWID {config['ifbw']}")
        self._query("INIT:CONT 0;*OPC?")

        self.current_preset_key = preset_key
        self.current_config = config
        self.start_freq = config["start_freq"]
        self.stop_freq = config["stop_freq"]
        self.points = config["points"]
        self.ifbw = config["ifbw"]
        return config

    def perform_calibration(self, switch_controller, progress_callback=None, should_stop=None):
        self._require_connected()
        if not self.current_config:
            raise N9918ANAError("Configure an NA preset before calibration.")
        if not switch_controller or not getattr(switch_controller, "connected", False):
            raise N9918ANAError("Switchbox is not connected.")

        events = []

        def emit(event):
            events.append(event.as_dict())
            if progress_callback:
                progress_callback(event.as_dict())

        try:
            self._check_stop(should_stop)
            self._set_switch_position(switch_controller, "LOAD")
            self._write("CORR:COLL:METH:QCAL:CAL 1")
            emit(CalibrationEvent("LOAD", "LOAD 校准", "B1D1", self.last_scpi, True, "已切到负载通道"))
            self._query("CORR:COLL:LOAD 1;*OPC?")
            emit(CalibrationEvent("LOAD_DONE", "LOAD 校准完成", "B1D1", self.last_scpi, True, "负载采集完成"))

            self._check_stop(should_stop)
            self._set_switch_position(switch_controller, "OPEN")
            emit(CalibrationEvent("OPEN", "OPEN 校准", "B2D1", None, True, "已切到开路通道"))
            self._query("CORR:COLL:INT 1;*OPC?")
            emit(CalibrationEvent("OPEN_DONE", "OPEN 校准完成", "B2D1", self.last_scpi, True, "QuickCal open-port 采集完成"))

            self._check_stop(should_stop)
            self._write("CORR:COLL:SAVE 0")
            emit(CalibrationEvent("SAVE", "保存校准", "B2D1", self.last_scpi, True, "校准参数已保存"))
            self._set_switch_position(switch_controller, "ANTENNA")
            emit(CalibrationEvent("ANTENNA", "切到天线测量", "B2D2", None, True, "已切到天线通道"))
        except N9918ANAError:
            raise
        except Exception as exc:
            raise N9918ANAError(str(exc), step="CALIBRATION", last_scpi=self.last_scpi, switch_position=self.last_switch_position) from exc

        return {
            "complete": True,
            "events": events,
            "last_scpi": self.last_scpi,
            "switch_position": self.last_switch_position,
        }

    def measure_s11(self, should_stop=None):
        self._require_connected()
        if not self.current_config:
            raise N9918ANAError("Configure an NA preset before measurement.")

        self._check_stop(should_stop)
        self._write("CALC:FORM MLOG")
        self._query("INIT:IMM;*OPC?")
        self._check_stop(should_stop)
        fdata = self._query("CALC:DATA:FDATa?")
        self._check_stop(should_stop)
        sdata = self._query("CALC:DATA:SDATA?")

        s11_db = parse_float_csv(fdata)
        real, imag = parse_complex_csv(sdata)
        points = int(self.current_config["points"])
        if len(s11_db) != points:
            raise N9918ANAError(
                f"FDATa point count mismatch: expected {points}, got {len(s11_db)}",
                step="MEASURE",
                last_scpi=self.last_scpi,
                switch_position=self.last_switch_position,
            )
        if len(real) != points or len(imag) != points:
            raise N9918ANAError(
                f"SDATA point count mismatch: expected {points}, got {len(real)} complex points",
                step="MEASURE",
                last_scpi=self.last_scpi,
                switch_position=self.last_switch_position,
            )

        frequencies = frequency_axis(self.current_config["start_freq"], self.current_config["stop_freq"], points)
        return build_na_result(
            frequencies,
            s11_db,
            real,
            imag,
            self.current_config,
            self.current_preset_key,
        )

    def _set_switch_position(self, switch_controller, position_key):
        positions = SWITCH_POSITIONS[position_key]
        for switch_name in ("B", "D"):
            switch_controller.set_switch(switch_name, positions[switch_name])
        self.last_switch_position = f"B{positions['B']}D{positions['D']}"

    def _require_connected(self):
        if not self.connected or not self.device:
            raise N9918ANAError("N9918A NA controller is not connected.", last_scpi=self.last_scpi, switch_position=self.last_switch_position)

    def _write(self, command):
        self.last_scpi = command
        self.device.write(command)

    def _query(self, command):
        self.last_scpi = command
        return self.device.query(command)

    @staticmethod
    def _check_stop(should_stop):
        if should_stop and should_stop():
            raise N9918ANAError("NA operation was stopped by user.", step="STOPPED")


def parse_float_csv(raw):
    if raw is None:
        return []
    values = []
    for part in str(raw).replace("\n", ",").split(","):
        part = part.strip()
        if part:
            values.append(float(part))
    return values


def parse_complex_csv(raw):
    values = parse_float_csv(raw)
    if len(values) % 2 != 0:
        raise ValueError("SDATA must contain real/imag pairs.")
    return values[0::2], values[1::2]


def frequency_axis(start_hz, stop_hz, points):
    points = int(points)
    if points <= 1:
        return [float(start_hz)]
    step = (float(stop_hz) - float(start_hz)) / (points - 1)
    return [float(start_hz) + step * i for i in range(points)]


def find_s11_valleys(frequencies, s11_db, min_separation_points=4, max_valleys=None):
    if not frequencies or not s11_db or len(frequencies) != len(s11_db):
        return []
    if len(s11_db) < 3:
        idx = min(range(len(s11_db)), key=lambda i: s11_db[i])
        return [_valley_dict(frequencies, s11_db, idx, 0.0)]

    candidates = []
    window = max(3, min(40, len(s11_db) // 80 or 3))
    for idx in range(1, len(s11_db) - 1):
        if s11_db[idx] <= s11_db[idx - 1] and s11_db[idx] <= s11_db[idx + 1]:
            left_max = max(s11_db[max(0, idx - window):idx] or [s11_db[idx]])
            right_max = max(s11_db[idx + 1:min(len(s11_db), idx + window + 1)] or [s11_db[idx]])
            prominence = min(left_max, right_max) - s11_db[idx]
            candidates.append((idx, prominence, s11_db[idx]))

    if not candidates:
        idx = min(range(len(s11_db)), key=lambda i: s11_db[i])
        candidates = [(idx, 0.0, s11_db[idx])]

    candidates.sort(key=lambda item: (item[2], -item[1]))
    selected = []
    for idx, prominence, _value in candidates:
        if all(abs(idx - prev_idx) >= min_separation_points for prev_idx, _ in selected):
            selected.append((idx, prominence))
        if max_valleys and len(selected) >= max_valleys:
            break

    selected.sort(key=lambda item: frequencies[item[0]])
    return [_valley_dict(frequencies, s11_db, idx, prominence) for idx, prominence in selected]


def _valley_dict(frequencies, s11_db, idx, prominence):
    return {
        "index": idx,
        "frequency_hz": frequencies[idx],
        "frequency_mhz": frequencies[idx] / 1e6,
        "s11_db": s11_db[idx],
        "prominence_db": max(0.0, prominence),
    }


def choose_primary_valley(valleys):
    if not valleys:
        return None
    valley = min(valleys, key=lambda item: item["s11_db"])
    return {k: v for k, v in valley.items() if k != "index"}


def calculate_bandwidth(frequencies, s11_db, valley_index, threshold_db):
    if not frequencies or not s11_db or valley_index is None:
        return _empty_bandwidth(threshold_db)
    if s11_db[valley_index] > threshold_db:
        return _empty_bandwidth(threshold_db)

    left_hz = frequencies[0]
    left_s11 = s11_db[0]
    left_complete = False
    for idx in range(valley_index, 0, -1):
        inside = s11_db[idx] <= threshold_db
        outside = s11_db[idx - 1] > threshold_db
        if inside and outside:
            left_hz = interpolate_x(frequencies[idx - 1], s11_db[idx - 1], frequencies[idx], s11_db[idx], threshold_db)
            left_s11 = threshold_db
            left_complete = True
            break
    else:
        left_complete = frequencies[0] != frequencies[valley_index] and s11_db[0] <= threshold_db

    right_hz = frequencies[-1]
    right_s11 = s11_db[-1]
    right_complete = False
    for idx in range(valley_index, len(s11_db) - 1):
        inside = s11_db[idx] <= threshold_db
        outside = s11_db[idx + 1] > threshold_db
        if inside and outside:
            right_hz = interpolate_x(frequencies[idx], s11_db[idx], frequencies[idx + 1], s11_db[idx + 1], threshold_db)
            right_s11 = threshold_db
            right_complete = True
            break
    else:
        right_complete = frequencies[-1] != frequencies[valley_index] and s11_db[-1] <= threshold_db

    complete = left_complete and right_complete
    return {
        "left_hz": left_hz,
        "right_hz": right_hz,
        "width_hz": max(0.0, right_hz - left_hz),
        "left_s11_db": left_s11,
        "right_s11_db": right_s11,
        "complete": complete,
        "threshold_db": threshold_db,
    }


def _empty_bandwidth(threshold_db):
    return {
        "left_hz": None,
        "right_hz": None,
        "width_hz": None,
        "left_s11_db": None,
        "right_s11_db": None,
        "complete": False,
        "threshold_db": threshold_db,
    }


def interpolate_x(x0, y0, x1, y1, target_y):
    if y1 == y0:
        return (x0 + x1) / 2
    ratio = (target_y - y0) / (y1 - y0)
    return x0 + ratio * (x1 - x0)


def interpolate_series(frequencies, values, target_hz):
    if target_hz is None or not frequencies or not values:
        return None
    if target_hz <= frequencies[0]:
        return values[0]
    if target_hz >= frequencies[-1]:
        return values[-1]
    for idx in range(1, len(frequencies)):
        if frequencies[idx] >= target_hz:
            return _interpolate_y(frequencies[idx - 1], values[idx - 1], frequencies[idx], values[idx], target_hz)
    return values[-1]


def _interpolate_y(x0, y0, x1, y1, target_x):
    if x1 == x0:
        return y0
    ratio = (target_x - x0) / (x1 - x0)
    return y0 + ratio * (y1 - y0)


def calculate_all_bandwidths(frequencies, s11_db, valley):
    if not valley:
        return {
            "absolute_3db": _empty_bandwidth(-3.0),
            "absolute_10db": _empty_bandwidth(-10.0),
            "relative_3db": _empty_bandwidth(None),
            "relative_10db": _empty_bandwidth(None),
        }
    valley_index = valley["index"]
    valley_s11 = valley["s11_db"]
    return {
        "absolute_3db": calculate_bandwidth(frequencies, s11_db, valley_index, -3.0),
        "absolute_10db": calculate_bandwidth(frequencies, s11_db, valley_index, -10.0),
        "relative_3db": calculate_bandwidth(frequencies, s11_db, valley_index, valley_s11 + 3.0),
        "relative_10db": calculate_bandwidth(frequencies, s11_db, valley_index, valley_s11 + 10.0),
    }


def bandwidth_for_valley(frequencies, s11_db, valley):
    if not valley:
        return calculate_all_bandwidths(frequencies, s11_db, None)
    return calculate_all_bandwidths(frequencies, s11_db, valley)


def build_smith_payload(frequencies, real, imag, s11_db, primary_valley_with_index, bandwidths, full_sweep=False):
    if full_sweep:
        return None
    markers = []
    if primary_valley_with_index:
        idx = primary_valley_with_index["index"]
        markers.append(
            {
                "type": "center",
                "label": "中心谷值",
                "frequency_hz": frequencies[idx],
                "frequency_mhz": frequencies[idx] / 1e6,
                "s11_db": s11_db[idx],
                "real": real[idx],
                "imag": imag[idx],
            }
        )

    marker_labels = {
        "absolute_3db": "绝对 -3dB",
        "absolute_10db": "绝对 -10dB",
        "relative_3db": "相对 +3dB",
        "relative_10db": "相对 +10dB",
    }
    for key, bw in (bandwidths or {}).items():
        for side in ("left", "right"):
            target = bw.get(f"{side}_hz")
            if target is None:
                continue
            markers.append(
                {
                    "type": f"{key}_{side}",
                    "label": f"{marker_labels.get(key, key)} {'左端点' if side == 'left' else '右端点'}",
                    "frequency_hz": target,
                    "frequency_mhz": target / 1e6,
                    "s11_db": bw.get(f"{side}_s11_db"),
                    "real": interpolate_series(frequencies, real, target),
                    "imag": interpolate_series(frequencies, imag, target),
                }
            )

    return {
        "real": [round(v, 8) for v in real],
        "imag": [round(v, 8) for v in imag],
        "markers": markers,
    }


def build_na_result(frequencies, s11_db, real=None, imag=None, config=None, preset_key=None):
    frequencies = [float(v) for v in frequencies]
    s11_db = [float(v) for v in s11_db]
    real = [float(v) for v in (real if real is not None else [0.0] * len(frequencies))]
    imag = [float(v) for v in (imag if imag is not None else [0.0] * len(frequencies))]
    full_sweep = bool(config.get("full_sweep")) if config else False

    max_valleys = None if full_sweep else 50
    valleys = find_s11_valleys(
        frequencies,
        s11_db,
        min_separation_points=max(3, len(frequencies) // 400),
        max_valleys=max_valleys,
    )
    primary_with_index = min(valleys, key=lambda item: item["s11_db"]) if valleys else None
    primary = choose_primary_valley(valleys)
    bandwidths = calculate_all_bandwidths(frequencies, s11_db, primary_with_index)

    valley_lookup = {item["index"]: item for item in valleys}
    valleys_payload = []
    for valley in valleys:
        bw = bandwidth_for_valley(frequencies, s11_db, valley)
        item = {k: v for k, v in valley.items() if k != "index"}
        item["bandwidths"] = bw
        valleys_payload.append(item)

    smith = build_smith_payload(frequencies, real, imag, s11_db, primary_with_index, bandwidths, full_sweep=full_sweep)
    return {
        "config": config or {},
        "preset_key": preset_key,
        "series": {
            "frequency_mhz": [round(freq / 1e6, 6) for freq in frequencies],
            "s11_db": [round(value, 4) for value in s11_db],
        },
        "smith": smith,
        "primary_valley": primary,
        "bandwidths": bandwidths,
        "valleys": valleys_payload,
        "is_full_sweep": full_sweep,
        "measurement_time": time.strftime("%Y-%m-%d %H:%M:%S"),
    }


def save_na_measurement_data(result, filename_prefix=None, output_dir="measurement_data"):
    if not result or not result.get("series"):
        raise ValueError("No NA result to save.")
    output = Path(output_dir)
    output.mkdir(exist_ok=True)
    if filename_prefix is None:
        filename_prefix = f"na_antenna_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    saved = []
    csv_path = output / f"{filename_prefix}_s11.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["Frequency (MHz)", "S11 (dB)"])
        for freq_mhz, s11 in zip(result["series"].get("frequency_mhz", []), result["series"].get("s11_db", [])):
            writer.writerow([freq_mhz, s11])
    saved.append(str(csv_path))

    valleys_path = output / f"{filename_prefix}_valleys.csv"
    with valleys_path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["Frequency (MHz)", "S11 (dB)", "Prominence (dB)"])
        for valley in result.get("valleys", []):
            writer.writerow([valley.get("frequency_mhz"), valley.get("s11_db"), valley.get("prominence_db")])
    saved.append(str(valleys_path))

    json_path = output / f"{filename_prefix}_result.json"
    json_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    saved.append(str(json_path))
    return saved


def export_na_report(result, user_info=None, output_dir="reports"):
    if not result or not result.get("series"):
        raise ValueError("No NA result to export.")
    output = Path(output_dir)
    output.mkdir(exist_ok=True)
    config = result.get("config") or {}
    label = config.get("label") or result.get("preset_key") or "NA"
    filename = safe_filename(f"NA天线测量-{label}-{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
    path = output / filename

    primary = result.get("primary_valley") or {}
    bandwidths = result.get("bandwidths") or {}
    valleys = result.get("valleys") or []
    user_info = user_info or {}

    rows = []
    for key, title in [
        ("absolute_3db", "绝对 -3dB"),
        ("absolute_10db", "绝对 -10dB"),
        ("relative_3db", "相对 +3dB"),
        ("relative_10db", "相对 +10dB"),
    ]:
        bw = bandwidths.get(key) or {}
        rows.append(
            "<tr>"
            f"<td>{html.escape(title)}</td>"
            f"<td>{format_hz(bw.get('left_hz'))}</td>"
            f"<td>{format_hz(bw.get('right_hz'))}</td>"
            f"<td>{format_hz(bw.get('width_hz'))}</td>"
            f"<td>{'完整' if bw.get('complete') else '不完整/未跨阈值'}</td>"
            "</tr>"
        )
    valley_rows = []
    for index, valley in enumerate(valleys[:200], start=1):
        valley_rows.append(
            "<tr>"
            f"<td>{index}</td>"
            f"<td>{valley.get('frequency_mhz', 0):.6f}</td>"
            f"<td>{valley.get('s11_db', 0):.3f}</td>"
            f"<td>{valley.get('prominence_db', 0):.3f}</td>"
            "</tr>"
        )

    body = f"""
<!doctype html>
<html lang=\"zh-CN\">
<head>
  <meta charset=\"utf-8\" />
  <title>NA 天线测量报告</title>
  <style>
    body {{ font-family: 'Microsoft YaHei', sans-serif; color: #172026; margin: 32px; }}
    h1 {{ font-size: 30px; }}
    table {{ border-collapse: collapse; width: 100%; margin: 14px 0 28px; }}
    th, td {{ border: 1px solid #ccd6d8; padding: 8px 10px; text-align: left; }}
    th {{ background: #eef6f5; }}
    .muted {{ color: #60717a; }}
  </style>
</head>
<body>
  <h1>N9918A NA 天线测量报告</h1>
  <p class=\"muted\">生成时间：{html.escape(result.get('measurement_time') or '')}</p>
  <h2>项目信息</h2>
  <p>客户：{html.escape(str(user_info.get('customer', '')))}；EUT：{html.escape(str(user_info.get('eut', '')))}；型号：{html.escape(str(user_info.get('model', '')))}；工程师：{html.escape(str(user_info.get('engineer', '')))}</p>
  <h2>配置</h2>
  <p>{html.escape(config.get('name', 'NA'))}，{format_hz(config.get('start_freq'))} - {format_hz(config.get('stop_freq'))}，点数 {config.get('points', '--')}，IFBW {format_hz(config.get('ifbw'))}</p>
  <h2>中心谷值</h2>
  <p>频率：{primary.get('frequency_mhz', 0):.6f} MHz；S11：{primary.get('s11_db', 0):.3f} dB</p>
  <h2>带宽结果</h2>
  <table><thead><tr><th>口径</th><th>左端点</th><th>右端点</th><th>带宽</th><th>状态</th></tr></thead><tbody>{''.join(rows)}</tbody></table>
  <h2>谷值列表</h2>
  <table><thead><tr><th>#</th><th>频率 MHz</th><th>S11 dB</th><th>显著性 dB</th></tr></thead><tbody>{''.join(valley_rows)}</tbody></table>
</body>
</html>
"""
    path.write_text(body, encoding="utf-8")
    return path


def safe_filename(filename):
    import re

    filename = re.sub(r'[<>:"/\\|?*]+', "_", filename)
    return filename.strip().strip(".") or f"na_report_{int(time.time())}.html"


def format_hz(value):
    if value is None:
        return "--"
    value = float(value)
    abs_value = abs(value)
    if abs_value >= 1e9:
        return f"{value / 1e9:.6g} GHz"
    if abs_value >= 1e6:
        return f"{value / 1e6:.6g} MHz"
    if abs_value >= 1e3:
        return f"{value / 1e3:.6g} kHz"
    return f"{value:.6g} Hz"
