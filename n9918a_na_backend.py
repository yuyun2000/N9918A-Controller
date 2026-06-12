
# n9918a_na_backend.py
"""N9918A NA/S11 antenna measurement controller and analysis helpers."""

from __future__ import annotations

import csv
import json
import math
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
    "ANT_868": {
        "name": "868MHz 天线",
        "label": "868MHz",
        "start_freq": 768e6,
        "stop_freq": 968e6,
        "points": 2001,
        "ifbw": 10e3,
        "full_sweep": False,
        "description": "768-968MHz，适合 868MHz 天线调试。",
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
    "LOAD": {"B": 1, "C": 1},
    "OPEN": {"B": 2, "C": 1},
    "ANTENNA": {"B": 2, "C": 2},
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
            self._set_switch_position(switch_controller, "OPEN")
            self._write("CORR:COLL:METH:QCAL:CAL 1")
            emit(CalibrationEvent("OPEN", "OPEN 校准", "B2C1", self.last_scpi, True, "已切到开路通道"))
            self._query("CORR:COLL:INT 1;*OPC?")
            emit(CalibrationEvent("OPEN_DONE", "OPEN 校准完成", "B2C1", self.last_scpi, True, "QuickCal 内部开路/短路采集完成"))

            self._check_stop(should_stop)
            self._set_switch_position(switch_controller, "LOAD")
            emit(CalibrationEvent("LOAD", "LOAD 校准", "B1C1", None, True, "已切到负载通道"))
            self._query("CORR:COLL:LOAD 1;*OPC?")
            emit(CalibrationEvent("LOAD_DONE", "LOAD 校准完成", "B1C1", self.last_scpi, True, "负载采集完成"))

            self._check_stop(should_stop)
            self._write("CORR:COLL:SAVE 0")
            emit(CalibrationEvent("SAVE", "保存校准", "B1C1", self.last_scpi, True, "校准参数已保存"))
            self._set_switch_position(switch_controller, "ANTENNA")
            emit(CalibrationEvent("ANTENNA", "切到天线测量", "B2C2", None, True, "已切到天线通道"))
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
        for switch_name in ("B", "C"):
            switch_controller.set_switch(switch_name, positions[switch_name])
        self.last_switch_position = f"B{positions['B']}C{positions['C']}"

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
    s11_value = s11_db[idx]
    return {
        "index": idx,
        "frequency_hz": frequencies[idx],
        "frequency_mhz": frequencies[idx] / 1e6,
        "s11_db": s11_value,
        "return_loss_db": return_loss_from_s11_db(s11_value),
        "vswr": vswr_from_s11_db(s11_value),
        "prominence_db": max(0.0, prominence),
    }


def choose_primary_valley(valleys):
    if not valleys:
        return None
    valley = min(valleys, key=lambda item: item["s11_db"])
    return {k: v for k, v in valley.items() if k != "index"}


def return_loss_from_s11_db(s11_db):
    if s11_db is None:
        return None
    return round(-float(s11_db), 4)


def gamma_from_s11_db(s11_db):
    if s11_db is None:
        return None
    return 10 ** (float(s11_db) / 20.0)


def vswr_from_s11_db(s11_db):
    gamma = gamma_from_s11_db(s11_db)
    if gamma is None or gamma >= 0.999999:
        return None
    return round((1.0 + gamma) / (1.0 - gamma), 4)


def metric_point(label, frequency_hz, s11_db, point_type=None):
    if frequency_hz is None or s11_db is None:
        return None
    return {
        "type": point_type or label,
        "label": label,
        "frequency_hz": frequency_hz,
        "frequency_mhz": frequency_hz / 1e6,
        "s11_db": s11_db,
        "return_loss_db": return_loss_from_s11_db(s11_db),
        "vswr": vswr_from_s11_db(s11_db),
    }


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
        "left_return_loss_db": return_loss_from_s11_db(left_s11),
        "right_return_loss_db": return_loss_from_s11_db(right_s11),
        "left_vswr": vswr_from_s11_db(left_s11),
        "right_vswr": vswr_from_s11_db(right_s11),
        "complete": complete,
        "threshold_db": threshold_db,
        "threshold_return_loss_db": return_loss_from_s11_db(threshold_db),
        "threshold_vswr": vswr_from_s11_db(threshold_db),
    }


def _empty_bandwidth(threshold_db):
    return {
        "left_hz": None,
        "right_hz": None,
        "width_hz": None,
        "left_s11_db": None,
        "right_s11_db": None,
        "left_return_loss_db": None,
        "right_return_loss_db": None,
        "left_vswr": None,
        "right_vswr": None,
        "complete": False,
        "threshold_db": threshold_db,
        "threshold_return_loss_db": return_loss_from_s11_db(threshold_db),
        "threshold_vswr": vswr_from_s11_db(threshold_db),
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
                "return_loss_db": return_loss_from_s11_db(s11_db[idx]),
                "vswr": vswr_from_s11_db(s11_db[idx]),
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
                    "return_loss_db": bw.get(f"{side}_return_loss_db"),
                    "vswr": bw.get(f"{side}_vswr"),
                    "real": interpolate_series(frequencies, real, target),
                    "imag": interpolate_series(frequencies, imag, target),
                }
            )

    return {
        "real": [round(v, 8) for v in real],
        "imag": [round(v, 8) for v in imag],
        "markers": markers,
    }


def build_points_of_interest(frequencies, s11_db, primary_valley_with_index, bandwidths):
    points = []
    if primary_valley_with_index:
        idx = primary_valley_with_index["index"]
        point = metric_point("中心谷值", frequencies[idx], s11_db[idx], "center")
        if point:
            points.append(point)

    labels = {
        "absolute_3db": "绝对 -3dB",
        "absolute_10db": "绝对 -10dB",
        "relative_3db": "相对 +3dB",
        "relative_10db": "相对 +10dB",
    }
    for key in ("absolute_3db", "absolute_10db", "relative_3db", "relative_10db"):
        bw = (bandwidths or {}).get(key) or {}
        for side, side_label in (("left", "左端点"), ("right", "右端点")):
            point = metric_point(
                f"{labels[key]} {side_label}",
                bw.get(f"{side}_hz"),
                bw.get(f"{side}_s11_db"),
                f"{key}_{side}",
            )
            if point:
                point["bandwidth_key"] = key
                point["side"] = side
                points.append(point)
    return points


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
    points_of_interest = build_points_of_interest(frequencies, s11_db, primary_with_index, bandwidths)

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
            "return_loss_db": [return_loss_from_s11_db(value) for value in s11_db],
            "vswr": [vswr_from_s11_db(value) for value in s11_db],
        },
        "smith": smith,
        "primary_valley": primary,
        "bandwidths": bandwidths,
        "points_of_interest": points_of_interest,
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
        writer.writerow(["Frequency (MHz)", "S11 (dB)", "Return Loss (dB)", "VSWR"])
        for freq_mhz, s11, return_loss, vswr in zip(
            result["series"].get("frequency_mhz", []),
            result["series"].get("s11_db", []),
            result["series"].get("return_loss_db", []),
            result["series"].get("vswr", []),
        ):
            writer.writerow([freq_mhz, s11, return_loss, vswr if vswr is not None else "Infinity"])
    saved.append(str(csv_path))

    valleys_path = output / f"{filename_prefix}_valleys.csv"
    with valleys_path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["Frequency (MHz)", "S11 (dB)", "Return Loss (dB)", "VSWR", "Prominence (dB)"])
        for valley in result.get("valleys", []):
            writer.writerow(
                [
                    valley.get("frequency_mhz"),
                    valley.get("s11_db"),
                    valley.get("return_loss_db"),
                    valley.get("vswr") if valley.get("vswr") is not None else "Infinity",
                    valley.get("prominence_db"),
                ]
            )
    saved.append(str(valleys_path))

    json_path = output / f"{filename_prefix}_result.json"
    json_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    saved.append(str(json_path))
    return saved


def export_na_report(result, user_info=None, output_dir="reports"):
    if not result or not result.get("series"):
        raise ValueError("No NA result to export.")

    try:
        import matplotlib

        matplotlib.use("Agg", force=True)
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_pdf import PdfPages
        from matplotlib.font_manager import FontProperties
    except ImportError as exc:
        raise ValueError("缺少 matplotlib，无法生成 NA PDF 报告。") from exc

    output = Path(output_dir)
    output.mkdir(exist_ok=True)
    config = result.get("config") or {}
    label = config.get("label") or result.get("preset_key") or "NA"
    filename = safe_filename(f"NA天线测量-{label}-{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
    path = output / filename
    font = _load_report_font(FontProperties)
    logo_path = Path(__file__).resolve().parent / "assets" / "m5logo2022.png"

    with PdfPages(path) as pdf:
        fig = _build_na_summary_page(plt, result, user_info or {}, logo_path, font)
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        fig = _build_na_s11_page(plt, result, logo_path, font)
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        fig = _build_na_smith_page(plt, result, logo_path, font)
        pdf.savefig(fig, bbox_inches="tight")
        plt.close(fig)

        for fig in _build_na_valley_pages(plt, result, logo_path, font):
            pdf.savefig(fig, bbox_inches="tight")
            plt.close(fig)

    return path


def _load_report_font(font_properties_cls):
    for candidate in (
        Path(__file__).resolve().parent / "utils" / "simhei.ttf",
        Path(__file__).resolve().parent / "utils" / "simfang.ttf",
        Path("C:/Windows/Fonts/simhei.ttf"),
        Path("C:/Windows/Fonts/simfang.ttf"),
    ):
        if candidate.exists():
            return font_properties_cls(fname=str(candidate))
    return None


def _new_report_figure(plt, title, logo_path, font):
    fig = plt.figure(figsize=(11.69, 8.27))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis("off")
    if logo_path.exists():
        try:
            image = plt.imread(str(logo_path))
            logo_ax = fig.add_axes([0.055, 0.885, 0.095, 0.08])
            logo_ax.imshow(image)
            logo_ax.axis("off")
        except Exception:
            pass
    ax.text(0.16, 0.94, title, fontsize=21, fontproperties=font, weight="bold", color="#172026")
    ax.text(0.16, 0.902, "M5Stack Technology Co., Ltd / N9918A FieldFox", fontsize=10.5, fontproperties=font, color="#60717a")
    ax.plot([0.055, 0.945], [0.875, 0.875], color="#1f393f", linewidth=1.2)
    return fig, ax


def _build_na_summary_page(plt, result, user_info, logo_path, font):
    fig, ax = _new_report_figure(plt, "N9918A NA 天线测量报告", logo_path, font)
    config = result.get("config") or {}
    primary = result.get("primary_valley") or {}
    measurement_time = result.get("measurement_time") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    info_rows = [
        ["客户", user_info.get("customer", ""), "产品名/EUT", user_info.get("eut", "")],
        ["型号", user_info.get("model", ""), "工程师", user_info.get("engineer", "")],
        ["备注", user_info.get("remark", ""), "测试时间", measurement_time],
        ["天线预设", config.get("name", result.get("preset_key", "NA")), "频率范围", f"{format_hz(config.get('start_freq'))} - {format_hz(config.get('stop_freq'))}"],
        ["扫描点数", str(config.get("points", "--")), "IFBW", format_hz(config.get("ifbw"))],
    ]
    _draw_report_table(ax, info_rows, [0.06, 0.62, 0.88, 0.22], font, font_size=9.2)

    metric_rows = [
        ["中心频率", _format_mhz(primary.get("frequency_mhz")), "S11", _format_db(primary.get("s11_db"))],
        ["回波损耗", _format_db(primary.get("return_loss_db")), "驻波比 VSWR", _format_vswr(primary.get("vswr"))],
    ]
    _draw_report_table(ax, metric_rows, [0.06, 0.49, 0.88, 0.09], font, font_size=10)

    bandwidth_rows = [["口径", "左频率", "左S11", "左RL", "左VSWR", "右频率", "右S11", "右RL", "右VSWR", "带宽", "状态"]]
    for key, title in _bandwidth_titles():
        bw = (result.get("bandwidths") or {}).get(key) or {}
        bandwidth_rows.append(
            [
                title,
                format_hz(bw.get("left_hz")),
                _format_db(bw.get("left_s11_db")),
                _format_db(bw.get("left_return_loss_db")),
                _format_vswr(bw.get("left_vswr")),
                format_hz(bw.get("right_hz")),
                _format_db(bw.get("right_s11_db")),
                _format_db(bw.get("right_return_loss_db")),
                _format_vswr(bw.get("right_vswr")),
                format_hz(bw.get("width_hz")),
                "完整" if bw.get("complete") else "不完整/未跨阈值",
            ]
        )
    _draw_report_table(ax, bandwidth_rows, [0.035, 0.18, 0.93, 0.24], font, font_size=6.8, header=True)

    ax.text(0.06, 0.115, "说明：RL=Return Loss= -S11(dB)；VSWR 由 |Γ|=10^(S11/20) 计算。绝对口径使用 -3/-10dB 阈值，相对口径使用谷值+3/+10dB。", fontsize=9, fontproperties=font, color="#60717a")
    return fig


def _build_na_s11_page(plt, result, logo_path, font):
    fig, ax_title = _new_report_figure(plt, "S11 曲线与端点标记", logo_path, font)
    ax = fig.add_axes([0.08, 0.18, 0.84, 0.62])
    series = result.get("series") or {}
    x_vals = series.get("frequency_mhz") or []
    y_vals = series.get("s11_db") or []
    if not x_vals or not y_vals:
        ax.text(0.5, 0.5, "暂无 S11 数据", transform=ax.transAxes, ha="center", va="center", fontproperties=font)
        return fig

    if max(x_vals) / max(min(v for v in x_vals if v > 0), 0.001) > 4:
        ax.semilogx(x_vals, y_vals, color="#0a6a72", linewidth=1.5, label="S11 dB")
    else:
        ax.plot(x_vals, y_vals, color="#0a6a72", linewidth=1.5, label="S11 dB")
    ax.axhline(-3, color="#d9822b", linestyle="--", linewidth=1, label="-3dB")
    ax.axhline(-10, color="#b7442e", linestyle="--", linewidth=1, label="-10dB")

    colors = {
        "center": "#b7442e",
        "absolute_3db_left": "#d9822b",
        "absolute_3db_right": "#d9822b",
        "absolute_10db_left": "#b7442e",
        "absolute_10db_right": "#b7442e",
    }
    for point in result.get("points_of_interest", []):
        point_type = point.get("type")
        if point_type not in colors:
            continue
        ax.scatter(point.get("frequency_mhz"), point.get("s11_db"), s=38, color=colors[point_type], zorder=5)
        ax.annotate(
            f"{point.get('label')}\n{point.get('frequency_mhz'):.3f}MHz\nVSWR {_format_vswr(point.get('vswr'))}",
            (point.get("frequency_mhz"), point.get("s11_db")),
            textcoords="offset points",
            xytext=(8, 8),
            fontsize=7,
            fontproperties=font,
        )

    ax.set_xlabel("Frequency (MHz)", fontproperties=font)
    ax.set_ylabel("S11 (dB)", fontproperties=font)
    ax.grid(True, alpha=0.25)
    ax.legend(prop=font, loc="best")
    ax_title.text(0.08, 0.1, "图中标记中心谷值、绝对 -3dB 端点和绝对 -10dB 端点；报告首页表格给出端点 S11、回波损耗和 VSWR。", fontsize=9, fontproperties=font, color="#60717a")
    return fig


def _build_na_smith_page(plt, result, logo_path, font):
    fig, ax_title = _new_report_figure(plt, "Smith Chart 标记", logo_path, font)
    ax = fig.add_axes([0.1, 0.16, 0.55, 0.66])
    smith = result.get("smith")
    if not smith:
        ax.axis("off")
        ax.text(0.5, 0.5, "全扫宽结果不显示 Smith Chart", transform=ax.transAxes, ha="center", va="center", fontproperties=font)
        return fig

    circle = plt.Circle((0, 0), 1, color="#1f393f", fill=False, linewidth=1)
    ax.add_artist(circle)
    for radius in (0.25, 0.5, 0.75):
        ax.add_artist(plt.Circle((0, 0), radius, color="#1f393f", fill=False, alpha=0.18, linewidth=0.8))
    ax.axhline(0, color="#1f393f", linewidth=0.8, alpha=0.35)
    ax.axvline(0, color="#1f393f", linewidth=0.8, alpha=0.35)
    ax.plot(smith.get("real", []), smith.get("imag", []), color="#0a6a72", linewidth=1.4)

    marker_rows = [["标记", "频率", "S11", "RL", "VSWR"]]
    for marker in smith.get("markers", []):
        marker_type = marker.get("type")
        if marker_type not in {"center", "absolute_3db_left", "absolute_3db_right"}:
            continue
        color = "#b7442e" if marker_type == "center" else "#d9822b"
        ax.scatter(marker.get("real"), marker.get("imag"), s=46, color=color, zorder=5)
        ax.annotate(marker.get("label", ""), (marker.get("real"), marker.get("imag")), textcoords="offset points", xytext=(8, 6), fontsize=7, fontproperties=font)
        marker_rows.append(
            [
                marker.get("label", ""),
                _format_mhz(marker.get("frequency_mhz")),
                _format_db(marker.get("s11_db")),
                _format_db(marker.get("return_loss_db")),
                _format_vswr(marker.get("vswr")),
            ]
        )

    ax.set_aspect("equal", adjustable="box")
    ax.set_xlim(-1.05, 1.05)
    ax.set_ylim(-1.05, 1.05)
    ax.set_xlabel("Real(Γ)", fontproperties=font)
    ax.set_ylabel("Imag(Γ)", fontproperties=font)
    ax.grid(True, alpha=0.18)
    _draw_report_table(ax_title, marker_rows, [0.68, 0.25, 0.27, 0.42], font, font_size=7.2, header=True)
    ax_title.text(0.68, 0.71, "Smith Chart 按需求标记中心频率和绝对 -3dB 左右端点。", fontsize=9, fontproperties=font, color="#60717a")
    return fig


def _build_na_valley_pages(plt, result, logo_path, font):
    valleys = result.get("valleys") or []
    if not valleys:
        fig, ax = _new_report_figure(plt, "S11 谷值列表", logo_path, font)
        ax.text(0.06, 0.72, "暂无谷值数据。", fontsize=11, fontproperties=font)
        return [fig]

    pages = []
    chunk_size = 28
    for page_index, start in enumerate(range(0, len(valleys), chunk_size), start=1):
        fig, ax = _new_report_figure(plt, f"S11 谷值列表 第 {page_index} 页", logo_path, font)
        rows = [["#", "频率 MHz", "S11", "RL", "VSWR", "绝对-10dB带宽", "相对+3dB带宽"]]
        for index, valley in enumerate(valleys[start:start + chunk_size], start=start + 1):
            abs10 = (valley.get("bandwidths") or {}).get("absolute_10db") or {}
            rel3 = (valley.get("bandwidths") or {}).get("relative_3db") or {}
            rows.append(
                [
                    str(index),
                    f"{valley.get('frequency_mhz', 0):.6f}",
                    _format_db(valley.get("s11_db")),
                    _format_db(valley.get("return_loss_db")),
                    _format_vswr(valley.get("vswr")),
                    format_hz(abs10.get("width_hz")),
                    format_hz(rel3.get("width_hz")),
                ]
            )
        _draw_report_table(ax, rows, [0.045, 0.12, 0.91, 0.72], font, font_size=7.4, header=True)
        pages.append(fig)
    return pages


def _draw_report_table(ax, rows, bbox, font, font_size=8.0, header=False):
    table = ax.table(cellText=rows, bbox=bbox, cellLoc="left")
    table.auto_set_font_size(False)
    table.set_fontsize(font_size)
    for (row, _col), cell in table.get_celld().items():
        cell.set_edgecolor("#ccd6d8")
        cell.set_linewidth(0.55)
        if font:
            cell.get_text().set_fontproperties(font)
        if header and row == 0:
            cell.set_facecolor("#1f393f")
            cell.get_text().set_color("#fff7e8")
        elif row % 2 == 0:
            cell.set_facecolor("#f6f0e2")
        else:
            cell.set_facecolor("#ffffff")
    return table


def _bandwidth_titles():
    return [
        ("absolute_3db", "绝对 -3dB"),
        ("absolute_10db", "绝对 -10dB"),
        ("relative_3db", "相对 +3dB"),
        ("relative_10db", "相对 +10dB"),
    ]


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


def _format_mhz(value):
    if value is None:
        return "--"
    return f"{float(value):.6f} MHz"


def _format_db(value):
    if value is None:
        return "--"
    return f"{float(value):.3f} dB"


def _format_vswr(value):
    if value is None:
        return "∞"
    return f"{float(value):.3f}"
