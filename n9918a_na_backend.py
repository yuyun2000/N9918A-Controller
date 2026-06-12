
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
    "ANT_315": {
        "name": "315MHz 天线",
        "label": "315MHz",
        "start_freq": 215e6,
        "stop_freq": 415e6,
        "target_freq": 315e6,
        "points": 2001,
        "ifbw": 10e3,
        "full_sweep": False,
        "description": "215-415MHz，适合 315MHz 天线调试。",
    },
    "ANT_433": {
        "name": "433MHz 天线",
        "label": "433MHz",
        "start_freq": 300e6,
        "stop_freq": 500e6,
        "target_freq": 433e6,
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
        "target_freq": 868e6,
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
        "target_freq": 915e6,
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
        "target_freq": 2450e6,
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
        "target_freq": 5.0e9,
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
        "target_freq": None,
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

REFERENCE_IMPEDANCE_OHM = 50.0


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


def find_s11_valleys(frequencies, s11_db, min_separation_points=4, max_valleys=None, min_prominence_db=0.3):
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
    else:
        meaningful = [item for item in candidates if item[1] >= min_prominence_db]
        if meaningful:
            candidates = meaningful

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


def impedance_from_gamma(real, imag, reference_ohm=REFERENCE_IMPEDANCE_OHM):
    if real is None or imag is None:
        return None
    gamma = complex(float(real), float(imag))
    denominator = 1 - gamma
    if abs(denominator) < 1e-12:
        return None
    impedance = float(reference_ohm) * (1 + gamma) / denominator
    if not math.isfinite(impedance.real) or not math.isfinite(impedance.imag):
        return None
    magnitude = abs(impedance)
    return {
        "reference_ohm": float(reference_ohm),
        "resistance_ohm": round(impedance.real, 4),
        "reactance_ohm": round(impedance.imag, 4),
        "impedance_magnitude_ohm": round(magnitude, 4),
        "impedance_label": format_impedance(impedance.real, impedance.imag),
    }


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


def build_smith_payload(
    frequencies,
    real,
    imag,
    s11_db,
    primary_valley_with_index,
    bandwidths,
    full_sweep=False,
    target_summary=None,
):
    if full_sweep:
        return None
    markers = []
    if primary_valley_with_index:
        idx = primary_valley_with_index["index"]
        markers.append(_smith_marker_payload(
            "center",
            "中心谷值",
            frequencies[idx],
            s11_db[idx],
            real[idx],
            imag[idx],
        ))

    if target_summary:
        target_hz = target_summary.get("target_frequency_hz")
        target_s11 = target_summary.get("target_s11_db")
        target_real = interpolate_series(frequencies, real, target_hz)
        target_imag = interpolate_series(frequencies, imag, target_hz)
        markers.append(
            _smith_marker_payload(
                "target",
                "理想频点",
                target_hz,
                target_s11,
                target_real,
                target_imag,
            )
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
                _smith_marker_payload(
                    f"{key}_{side}",
                    f"{marker_labels.get(key, key)} {'左端点' if side == 'left' else '右端点'}",
                    target,
                    bw.get(f"{side}_s11_db"),
                    interpolate_series(frequencies, real, target),
                    interpolate_series(frequencies, imag, target),
                    return_loss_db=bw.get(f"{side}_return_loss_db"),
                    vswr=bw.get(f"{side}_vswr"),
                )
            )

    return {
        "reference_ohm": REFERENCE_IMPEDANCE_OHM,
        "real": [round(v, 8) for v in real],
        "imag": [round(v, 8) for v in imag],
        "markers": markers,
    }


def _smith_marker_payload(marker_type, label, frequency_hz, s11_db, real, imag, return_loss_db=None, vswr=None):
    marker = {
        "type": marker_type,
        "label": label,
        "frequency_hz": frequency_hz,
        "frequency_mhz": frequency_hz / 1e6 if frequency_hz is not None else None,
        "s11_db": s11_db,
        "return_loss_db": return_loss_db if return_loss_db is not None else return_loss_from_s11_db(s11_db),
        "vswr": vswr if vswr is not None else vswr_from_s11_db(s11_db),
        "real": real,
        "imag": imag,
    }
    impedance = impedance_from_gamma(real, imag)
    if impedance:
        marker.update(impedance)
    return marker


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


def build_target_summary(frequencies, s11_db, config, primary_valley_with_index):
    target_hz = (config or {}).get("target_freq")
    if target_hz is None or not frequencies or not s11_db:
        return None
    target_hz = float(target_hz)
    if target_hz < frequencies[0] or target_hz > frequencies[-1]:
        return None

    target_s11 = interpolate_series(frequencies, s11_db, target_hz)
    target_point = metric_point("理想频点", target_hz, target_s11, "target")
    if not target_point:
        return None

    summary = {
        **target_point,
        "target_frequency_hz": target_hz,
        "target_frequency_mhz": target_hz / 1e6,
        "target_s11_db": target_point["s11_db"],
        "target_return_loss_db": target_point["return_loss_db"],
        "target_vswr": target_point["vswr"],
    }
    if primary_valley_with_index:
        idx = primary_valley_with_index["index"]
        actual_hz = frequencies[idx]
        actual_s11 = s11_db[idx]
        error_hz = actual_hz - target_hz
        target_return_loss = target_point["return_loss_db"]
        actual_return_loss = return_loss_from_s11_db(actual_s11)
        error_percent = (error_hz / target_hz * 100.0) if target_hz else 0.0
        summary.update(
            {
                "actual_frequency_hz": actual_hz,
                "actual_frequency_mhz": actual_hz / 1e6,
                "actual_s11_db": actual_s11,
                "actual_return_loss_db": actual_return_loss,
                "actual_vswr": vswr_from_s11_db(actual_s11),
                "frequency_error_hz": error_hz,
                "frequency_error_mhz": error_hz / 1e6,
                "abs_frequency_error_hz": abs(error_hz),
                "abs_frequency_error_mhz": abs(error_hz) / 1e6,
                "frequency_error_percent": error_percent,
                "abs_frequency_error_percent": abs(error_percent),
                "s11_delta_db": actual_s11 - target_s11,
                "return_loss_delta_db": (
                    round(actual_return_loss - target_return_loss, 4)
                    if actual_return_loss is not None and target_return_loss is not None
                    else None
                ),
            }
        )
        summary["status"] = target_match_status(summary)
    else:
        summary["status"] = "unknown"
    summary["status_label"] = {
        "good": "接近理想",
        "warn": "轻微偏移",
        "bad": "明显偏移",
        "unknown": "暂无实际谷值",
    }.get(summary["status"], "暂无实际谷值")
    return summary


def build_target_window(frequencies, s11_db, config):
    target_hz = (config or {}).get("target_freq")
    if target_hz is None or not frequencies or not s11_db:
        return []
    target_hz = float(target_hz)
    points = []
    for offset_percent in (-3.0, -1.0, 0.0, 1.0, 3.0):
        freq_hz = target_hz * (1.0 + offset_percent / 100.0)
        if freq_hz < frequencies[0] or freq_hz > frequencies[-1]:
            continue
        s11_value = interpolate_series(frequencies, s11_db, freq_hz)
        point = metric_point(
            f"{offset_percent:+.0f}%",
            freq_hz,
            s11_value,
            "target_window",
        )
        if point:
            point["offset_percent"] = offset_percent
            point["offset_mhz"] = (freq_hz - target_hz) / 1e6
            points.append(point)
    return points


def target_match_status(summary):
    if not summary:
        return "unknown"
    error_percent = summary.get("abs_frequency_error_percent")
    target_return_loss = summary.get("target_return_loss_db")
    if error_percent is None:
        return "unknown"
    if error_percent <= 1.0 and (target_return_loss is None or target_return_loss >= 10.0):
        return "good"
    if error_percent <= 3.0 and (target_return_loss is None or target_return_loss >= 6.0):
        return "warn"
    return "bad"


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
    target_summary = build_target_summary(frequencies, s11_db, config, primary_with_index)
    target_window = build_target_window(frequencies, s11_db, config)
    if target_summary:
        points_of_interest.insert(
            1 if primary_with_index else 0,
            metric_point("理想频点", target_summary["target_frequency_hz"], target_summary["target_s11_db"], "target"),
        )

    valleys_payload = []
    for valley in valleys:
        bw = bandwidth_for_valley(frequencies, s11_db, valley)
        item = {k: v for k, v in valley.items() if k != "index"}
        item["bandwidths"] = bw
        valleys_payload.append(item)

    smith = build_smith_payload(
        frequencies,
        real,
        imag,
        s11_db,
        primary_with_index,
        bandwidths,
        full_sweep=full_sweep,
        target_summary=target_summary,
    )
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
        "target_summary": target_summary,
        "target_window": target_window,
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
        chart_pages = [
            _build_na_summary_page(plt, result, user_info or {}, logo_path, font),
            _build_na_s11_page(plt, result, logo_path, font),
            _build_na_vswr_page(plt, result, logo_path, font),
        ]
        if not result.get("is_full_sweep"):
            chart_pages.append(_build_na_smith_page(plt, result, logo_path, font))

        for fig in chart_pages:
            pdf.savefig(fig)
            plt.close(fig)

        for fig in _build_na_detail_pages(plt, result, user_info or {}, logo_path, font):
            pdf.savefig(fig)
            plt.close(fig)

        for fig in _build_na_valley_pages(plt, result, logo_path, font):
            pdf.savefig(fig)
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
    # Match the SA report's A4 portrait footprint while keeping matplotlib as a no-ReportLab fallback.
    fig = plt.figure(figsize=(8.27, 11.69))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis("off")
    if logo_path.exists():
        try:
            image = plt.imread(str(logo_path))
            logo_ax = fig.add_axes([0.065, 0.91, 0.09, 0.065])
            logo_ax.imshow(image)
            logo_ax.axis("off")
        except Exception:
            pass
    ax.text(0.18, 0.952, title, transform=ax.transAxes, fontsize=19, fontproperties=font, weight="bold", color="#172026")
    ax.text(0.18, 0.925, "M5Stack Technology Co., Ltd / N9918A FieldFox", transform=ax.transAxes, fontsize=9.5, fontproperties=font, color="#60717a")
    ax.plot([0.06, 0.94], [0.895, 0.895], transform=ax.transAxes, color="#1f393f", linewidth=1.0)
    return fig, ax




def _build_na_summary_page(plt, result, user_info, logo_path, font):
    fig, ax = _new_report_figure(plt, "N9918A NA 天线测量报告总览", logo_path, font)
    config = result.get("config") or {}
    primary = result.get("primary_valley") or {}
    target = result.get("target_summary") or {}
    bandwidths = result.get("bandwidths") or {}
    measurement_time = result.get("measurement_time") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    ax.text(
        0.06,
        0.855,
        "工程师优先看图：本报告先给出结论卡片和图表，详细测试信息、端点与候选谷值表集中放在末尾，便于复核和归档。",
        transform=ax.transAxes,
        fontsize=10.2,
        fontproperties=font,
        color="#1f393f",
        weight="bold",
    )

    info_line_1 = f"客户：{user_info.get('customer') or '--'}    产品/EUT：{user_info.get('eut') or '--'}    型号：{user_info.get('model') or '--'}"
    info_line_2 = f"工程师：{user_info.get('engineer') or '--'}    测试时间：{measurement_time}"
    info_line_3 = (
        f"预设：{config.get('name') or result.get('preset_key') or 'NA'}    "
        f"范围：{format_hz(config.get('start_freq'))} - {format_hz(config.get('stop_freq'))}    "
        f"点数：{config.get('points', '--')}    IFBW：{format_hz(config.get('ifbw'))}"
    )
    ax.text(
        0.06,
        0.805,
        f"{info_line_1}\n{info_line_2}\n{info_line_3}",
        transform=ax.transAxes,
        fontsize=8.6,
        fontproperties=font,
        color="#455a64",
        linespacing=1.65,
        bbox={"boxstyle": "round,pad=0.55", "fc": "#fff8e8", "ec": "#ccd6d8", "lw": 0.7},
    )

    abs10 = bandwidths.get("absolute_10db") or {}
    cards = [
        ("理想频点", _format_mhz(target.get("target_frequency_mhz")), "预设目标频率", "#23744a"),
        ("中心谷值", _format_mhz(primary.get("frequency_mhz")), "实测最深 S11 谷", "#b7442e"),
        ("频率偏移", _format_frequency_delta(target), target.get("status_label") or "全扫宽不对比", _target_status_color(target)),
        ("中心 S11 / RL", _format_s11_rl(primary.get("s11_db")), "RL = -S11", "#0a6a72"),
        ("中心 VSWR", _format_vswr(primary.get("vswr")), "越接近 1 越好", "#23744a"),
        ("绝对 -10dB 带宽", format_hz(abs10.get("width_hz")), "S11 <= -10dB 区间" if abs10.get("complete") else "未完整跨过阈值", "#d9822b"),
    ]
    _draw_report_card_grid(ax, cards, [0.06, 0.50, 0.88, 0.235], font, columns=3)

    primary_summary = (
        f"中心谷值位于 {_format_mhz(primary.get('frequency_mhz'))}，"
        f"S11/回波损耗为 {_format_s11_rl(primary.get('s11_db'))}，"
        f"VSWR 为 {_format_vswr(primary.get('vswr'))}。"
    )
    target_note = _format_target_comparison_note(target)
    ax.text(
        0.06,
        0.425,
        f"结论摘要：{primary_summary}\n目标对比：{target_note}",
        transform=ax.transAxes,
        fontsize=9.0,
        fontproperties=font,
        color="#172026",
        linespacing=1.6,
        bbox={"boxstyle": "round,pad=0.55", "fc": "#f6f0e2", "ec": "#ccd6d8", "lw": 0.6},
    )

    flow = [
        ("01", "S11 曲线", "看中心谷值、-3/-10dB 端点"),
        ("02", "VSWR 曲线", "看匹配好坏与阈值线"),
        ("03", "Smith Chart", "看 50Ω 归一化阻抗轨迹" if not result.get("is_full_sweep") else "全扫宽不显示 Smith Chart"),
        ("04", "末尾表格", "复核测试信息、端点和谷值"),
    ]
    _draw_report_flow(ax, flow, [0.06, 0.19, 0.88, 0.16], font)
    ax.text(
        0.06,
        0.115,
        "说明：S11 为负值，数值越低表示反射越小；回波损耗 RL=-S11，因此不再单独占列。VSWR 由 |Γ|=10^(S11/20) 换算，1:1 为理想匹配。",
        transform=ax.transAxes,
        fontsize=8.2,
        fontproperties=font,
        color="#60717a",
        linespacing=1.55,
    )
    return fig



def _format_target_table_note(target_summary):
    if not target_summary:
        return "全扫宽或未配置目标频率"
    return (
        f"{target_summary.get('status_label') or '--'}；"
        f"频偏 {_format_frequency_delta(target_summary)}；"
        f"RL差 {_format_db(target_summary.get('return_loss_delta_db'))}"
    )

def _target_status_color(target):
    return {
        "good": "#23744a",
        "warn": "#d9822b",
        "bad": "#b7442e",
    }.get((target or {}).get("status"), "#60717a")


def _draw_report_card_grid(ax, cards, bbox, font, columns=3):
    try:
        from matplotlib.patches import FancyBboxPatch
    except Exception:
        FancyBboxPatch = None

    x, y, width, height = bbox
    columns = max(1, int(columns))
    rows = max(1, (len(cards) + columns - 1) // columns)
    gap_x = 0.018
    gap_y = 0.018
    card_w = (width - gap_x * (columns - 1)) / columns
    card_h = (height - gap_y * (rows - 1)) / rows
    for index, (label, value, caption, color) in enumerate(cards):
        col = index % columns
        row = index // columns
        left = x + col * (card_w + gap_x)
        bottom = y + height - (row + 1) * card_h - row * gap_y
        if FancyBboxPatch:
            patch = FancyBboxPatch(
                (left, bottom),
                card_w,
                card_h,
                boxstyle="round,pad=0.012,rounding_size=0.018",
                transform=ax.transAxes,
                facecolor="#fff8e8",
                edgecolor="#ccd6d8",
                linewidth=0.8,
            )
            ax.add_patch(patch)
        ax.text(left + 0.018, bottom + card_h - 0.028, label, transform=ax.transAxes, fontsize=7.6, fontproperties=font, color="#60717a")
        ax.text(left + 0.018, bottom + card_h * 0.45, value or "--", transform=ax.transAxes, fontsize=13.2, fontproperties=font, color=color, weight="bold")
        ax.text(left + 0.018, bottom + 0.022, caption or "", transform=ax.transAxes, fontsize=7.3, fontproperties=font, color="#455a64")


def _draw_report_flow(ax, steps, bbox, font):
    try:
        from matplotlib.patches import FancyBboxPatch
    except Exception:
        FancyBboxPatch = None

    x, y, width, height = bbox
    gap = 0.014
    item_w = (width - gap * (len(steps) - 1)) / max(1, len(steps))
    for index, (number, title, caption) in enumerate(steps):
        left = x + index * (item_w + gap)
        if FancyBboxPatch:
            patch = FancyBboxPatch(
                (left, y),
                item_w,
                height,
                boxstyle="round,pad=0.012,rounding_size=0.02",
                transform=ax.transAxes,
                facecolor="#f6f0e2",
                edgecolor="#ccd6d8",
                linewidth=0.7,
            )
            ax.add_patch(patch)
        ax.text(left + 0.018, y + height - 0.035, number, transform=ax.transAxes, fontsize=8.0, fontproperties=font, color="#d9822b", weight="bold")
        ax.text(left + 0.018, y + height - 0.074, title, transform=ax.transAxes, fontsize=9.2, fontproperties=font, color="#172026", weight="bold")
        ax.text(left + 0.018, y + 0.025, caption, transform=ax.transAxes, fontsize=7.3, fontproperties=font, color="#60717a", linespacing=1.25)


def _build_na_detail_pages(plt, result, user_info, logo_path, font):
    pages = []
    config = result.get("config") or {}
    primary = result.get("primary_valley") or {}
    target = result.get("target_summary") or {}
    measurement_time = result.get("measurement_time") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    fig, ax = _new_report_figure(plt, "详细表格与测试口径", logo_path, font)
    ax.text(
        0.06,
        0.855,
        "本页集中放置报告首页移出的复核表格。工程师先看前面的图，本页用于确认测试条件、理想/实测偏差和带宽端点。",
        transform=ax.transAxes,
        fontsize=8.8,
        fontproperties=font,
        color="#455a64",
    )
    info_rows = [
        ["客户", user_info.get("customer", ""), "产品名/EUT", user_info.get("eut", "")],
        ["型号", user_info.get("model", ""), "工程师", user_info.get("engineer", "")],
        ["备注", user_info.get("remark", ""), "测试时间", measurement_time],
        ["天线预设", config.get("name", result.get("preset_key", "NA")), "频率范围", f"{format_hz(config.get('start_freq'))} - {format_hz(config.get('stop_freq'))}"],
        ["扫描点数", str(config.get("points", "--")), "IFBW", format_hz(config.get("ifbw"))],
    ]
    _draw_report_table(ax, info_rows, [0.06, 0.69, 0.88, 0.14], font, font_size=8.0)

    metric_rows = [
        ["项目", "频率", "S11 / RL", "VSWR", "说明"],
        ["理想频点", _format_mhz(target.get("target_frequency_mhz")), _format_s11_rl(target.get("target_s11_db")), _format_vswr(target.get("target_vswr")), "预设目标频率处插值"],
        ["实际谷值", _format_mhz(primary.get("frequency_mhz")), _format_s11_rl(primary.get("s11_db")), _format_vswr(primary.get("vswr")), "当前频段最深 S11 谷"],
        ["实际-理想", _format_mhz_delta(target.get("frequency_error_mhz")), _format_s11_rl(target.get("s11_delta_db")), "--", _format_target_table_note(target)],
    ]
    _draw_report_table(ax, metric_rows, [0.06, 0.515, 0.88, 0.145], font, font_size=7.2, header=True)

    bandwidth_rows = [["口径", "左端点", "右端点", "带宽", "状态"]]
    for key, title in _bandwidth_titles():
        bw = (result.get("bandwidths") or {}).get(key) or {}
        bandwidth_rows.append(
            [
                title,
                _format_endpoint_cell(bw, "left"),
                _format_endpoint_cell(bw, "right"),
                format_hz(bw.get("width_hz")),
                "完整" if bw.get("complete") else "不完整/未跨阈值",
            ]
        )
    _draw_report_table(ax, bandwidth_rows, [0.06, 0.225, 0.88, 0.245], font, font_size=6.7, header=True)
    ax.text(
        0.06,
        0.145,
        "表格说明：绝对口径统计 S11<=-3dB 或 <=-10dB 的连续区间；相对口径以中心谷值为基准，\n统计 S11<=谷值+3dB 或 谷值+10dB 的区间。端点由相邻采样点线性插值得到。",
        transform=ax.transAxes,
        fontsize=8.0,
        fontproperties=font,
        color="#60717a",
        linespacing=1.5,
    )
    pages.append(fig)

    fig, ax = _new_report_figure(plt, "关键频点与目标窗口", logo_path, font)
    ax.text(
        0.06,
        0.855,
        "关键频点表用于复核图中标注的中心、理想频点和带宽端点；目标窗口表用于查看理想频点附近 ±1% / ±3% 的匹配变化。",
        transform=ax.transAxes,
        fontsize=8.8,
        fontproperties=font,
        color="#455a64",
    )
    point_rows = [["标记", "频率 MHz", "S11 / RL", "VSWR"]]
    points = result.get("points_of_interest") or []
    if not points:
        point_rows.append(["--", "暂无关键频点", "--", "--"])
    for point in points[:14]:
        point_rows.append(
            [
                point.get("label", ""),
                _format_mhz(point.get("frequency_mhz")),
                _format_s11_rl(point.get("s11_db")),
                _format_vswr(point.get("vswr")),
            ]
        )
    _draw_report_table(ax, point_rows, [0.06, 0.47, 0.88, 0.35], font, font_size=6.9, header=True)

    target_rows = [["偏移", "频率", "S11 / RL", "VSWR"]]
    target_window = result.get("target_window") or []
    if not target_window:
        target_rows.append(["--", "全扫宽或未配置理想频点", "--", "--"])
    for point in target_window:
        target_rows.append(
            [
                _format_percent(point.get("offset_percent")),
                _format_mhz(point.get("frequency_mhz")),
                _format_s11_rl(point.get("s11_db")),
                _format_vswr(point.get("vswr")),
            ]
        )
    _draw_report_table(ax, target_rows, [0.06, 0.20, 0.88, 0.20], font, font_size=7.0, header=True)
    ax.text(
        0.06,
        0.125,
        "表格说明：S11/RL 合并显示，左侧为仪器回波曲线的 S11(dB)，右侧为回波损耗 RL(dB)。\nVSWR 与两者同步换算，用于快速判断天线匹配是否接近 1:1。",
        transform=ax.transAxes,
        fontsize=8.0,
        fontproperties=font,
        color="#60717a",
        linespacing=1.5,
    )
    pages.append(fig)
    return pages


def _build_na_s11_page(plt, result, logo_path, font):
    fig, ax_title = _new_report_figure(plt, "S11 曲线与端点标记", logo_path, font)
    ax = fig.add_axes([0.09, 0.34, 0.84, 0.52])
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
        "target": "#23744a",
        "absolute_3db_left": "#d9822b",
        "absolute_3db_right": "#d9822b",
        "absolute_10db_left": "#b7442e",
        "absolute_10db_right": "#b7442e",
    }
    label_points = []
    for point in result.get("points_of_interest", []):
        point_type = point.get("type")
        if point_type not in colors:
            continue
        if point_type in {"center", "target"}:
            ax.axvline(point.get("frequency_mhz"), color=colors[point_type], linestyle=":" if point_type == "center" else "-.", linewidth=1.1)
        ax.scatter(point.get("frequency_mhz"), point.get("s11_db"), s=38, color=colors[point_type], zorder=5)
        if point_type in {"center", "target", "absolute_3db_left", "absolute_3db_right", "absolute_10db_left", "absolute_10db_right"}:
            label_points.append(
                {
                    "x": point.get("frequency_mhz"),
                    "y": point.get("s11_db"),
                    "type": point_type,
                    "color": colors[point_type],
                    "text": _report_marker_label(point, value_key="s11"),
                }
            )

    ax.set_xlabel("Frequency (MHz)", fontproperties=font)
    ax.set_ylabel("S11 (dB)", fontproperties=font)
    ax.grid(True, alpha=0.25)
    ax.legend(prop=font, loc="best")
    _annotate_report_points(ax, label_points, font)
    ax_title.text(
        0.07,
        0.255,
        "读图说明：中心谷值、理想频点和绝对 -3/-10dB 左右端点直接标在曲线上。\nS11 越负表示回波越小，回波损耗 RL=-S11；详细端点数值见末尾“关键频点与目标窗口”表。",
        transform=ax_title.transAxes,
        fontsize=8.6,
        fontproperties=font,
        color="#60717a",
        linespacing=1.55,
        bbox={"boxstyle": "round,pad=0.45", "fc": "#fff8e8", "ec": "#ccd6d8", "lw": 0.6},
    )
    return fig


def _build_na_vswr_page(plt, result, logo_path, font):
    fig, ax_title = _new_report_figure(plt, "VSWR 驻波比曲线", logo_path, font)
    ax = fig.add_axes([0.09, 0.34, 0.84, 0.52])
    series = result.get("series") or {}
    x_vals = series.get("frequency_mhz") or []
    y_vals = series.get("vswr") or []
    plot_pairs = [(x, y) for x, y in zip(x_vals, y_vals) if y is not None and math.isfinite(float(y))]
    if not plot_pairs:
        ax.text(0.5, 0.5, "暂无 VSWR 数据", transform=ax.transAxes, ha="center", va="center", fontproperties=font)
        return fig

    plot_x, plot_y = zip(*plot_pairs)
    y_cap = _vswr_plot_cap(plot_y)
    clipped_y = [min(float(value), y_cap) for value in plot_y]
    if max(plot_x) / max(min(v for v in plot_x if v > 0), 0.001) > 4:
        ax.semilogx(plot_x, clipped_y, color="#23744a", linewidth=1.5, label="VSWR")
    else:
        ax.plot(plot_x, clipped_y, color="#23744a", linewidth=1.5, label="VSWR")

    threshold_rows = [
        ("S11=-10dB", vswr_from_s11_db(-10), "#b7442e"),
        ("S11=-3dB", vswr_from_s11_db(-3), "#d9822b"),
    ]
    for label, threshold, color in threshold_rows:
        if threshold and threshold <= y_cap:
            ax.axhline(threshold, color=color, linestyle="--", linewidth=1, label=f"{label} / VSWR {threshold:.3f}")

    colors = {
        "center": "#b7442e",
        "target": "#23744a",
        "absolute_3db_left": "#d9822b",
        "absolute_3db_right": "#d9822b",
        "absolute_10db_left": "#b7442e",
        "absolute_10db_right": "#b7442e",
    }
    label_points = []
    for point in result.get("points_of_interest", []):
        point_type = point.get("type")
        vswr = point.get("vswr")
        if point_type not in colors or vswr is None:
            continue
        if point_type in {"center", "target"}:
            ax.axvline(point.get("frequency_mhz"), color=colors[point_type], linestyle=":" if point_type == "center" else "-.", linewidth=1.1)
        ax.scatter(point.get("frequency_mhz"), min(float(vswr), y_cap), s=38, color=colors[point_type], zorder=5)
        if point_type in {"center", "target", "absolute_3db_left", "absolute_3db_right", "absolute_10db_left", "absolute_10db_right"}:
            label_points.append(
                {
                    "x": point.get("frequency_mhz"),
                    "y": min(float(vswr), y_cap),
                    "type": point_type,
                    "color": colors[point_type],
                    "text": _report_marker_label(point, value_key="vswr"),
                }
            )

    ax.set_ylim(1.0, max(1.2, y_cap))
    ax.set_xlabel("Frequency (MHz)", fontproperties=font)
    ax.set_ylabel("VSWR", fontproperties=font)
    ax.grid(True, alpha=0.25)
    ax.legend(prop=font, loc="best")
    _annotate_report_points(ax, label_points, font)
    ax_title.text(
        0.07,
        0.255,
        "读图说明：VSWR 越接近 1 表示匹配越好；图中参考线与 S11=-10dB/-3dB 同步。\n若某些点超过显示上限，会压到图顶端以保持曲线整体可读；详细数值见末尾表格。",
        transform=ax_title.transAxes,
        fontsize=8.6,
        fontproperties=font,
        color="#60717a",
        linespacing=1.55,
        bbox={"boxstyle": "round,pad=0.45", "fc": "#fff8e8", "ec": "#ccd6d8", "lw": 0.6},
    )
    return fig


def _build_na_smith_page(plt, result, logo_path, font):
    fig, ax_title = _new_report_figure(plt, "Smith Chart 与阻抗标记", logo_path, font)
    ax = fig.add_axes([0.12, 0.30, 0.76, 0.56])
    smith = result.get("smith")
    if not smith:
        ax.axis("off")
        ax.text(0.5, 0.5, "全扫宽结果不显示 Smith Chart", transform=ax.transAxes, ha="center", va="center", fontproperties=font)
        return fig

    reference_ohm = smith.get("reference_ohm") or REFERENCE_IMPEDANCE_OHM
    _draw_smith_impedance_grid(ax, reference_ohm, font)
    ax.plot(smith.get("real", []), smith.get("imag", []), color="#0a6a72", linewidth=1.4)

    label_points = []
    for marker in smith.get("markers", []):
        marker_type = marker.get("type")
        if marker_type not in {"center", "target", "absolute_3db_left", "absolute_3db_right"}:
            continue
        color = "#b7442e" if marker_type == "center" else "#23744a" if marker_type == "target" else "#d9822b"
        ax.scatter(marker.get("real"), marker.get("imag"), s=46, color=color, zorder=5)
        label_points.append(
            {
                "x": marker.get("real"),
                "y": marker.get("imag"),
                "type": marker_type,
                "color": color,
                "text": _report_smith_label(marker),
            }
        )

    ax.set_aspect("equal", adjustable="box")
    ax.set_xlim(-1.05, 1.05)
    ax.set_ylim(-1.05, 1.05)
    ax.set_xlabel("Real(Γ)", fontproperties=font)
    ax.set_ylabel("Imag(Γ)", fontproperties=font)
    ax.grid(True, alpha=0.18)
    _annotate_report_points(ax, label_points, font)
    ax_title.text(
        0.08,
        0.225,
        f"读图说明：Smith Chart 使用 SDATA? 复数 Γ 自绘，阻抗网格按 {reference_ohm:g}Ω 归一化。\n中心、理想频点和 -3dB 左右端点直接标出阻抗；完整阻抗/S11/VSWR 数值见末尾表格。",
        transform=ax_title.transAxes,
        fontsize=8.6,
        fontproperties=font,
        color="#60717a",
        linespacing=1.55,
        bbox={"boxstyle": "round,pad=0.45", "fc": "#fff8e8", "ec": "#ccd6d8", "lw": 0.6},
    )
    return fig


def _build_na_valley_pages(plt, result, logo_path, font):
    valleys = result.get("valleys") or []
    if not valleys:
        fig, ax = _new_report_figure(plt, "S11 候选谷值表", logo_path, font)
        ax.text(0.06, 0.72, "暂无谷值数据。", transform=ax.transAxes, fontsize=11, fontproperties=font)
        ax.text(0.06, 0.64, "表格说明：若未找到局部谷值，请检查扫描范围、点数和天线连接状态。", transform=ax.transAxes, fontsize=8.2, fontproperties=font, color="#60717a")
        return [fig]

    pages = []
    chunk_size = 20
    for page_index, start in enumerate(range(0, len(valleys), chunk_size), start=1):
        fig, ax = _new_report_figure(plt, f"S11 候选谷值表 第 {page_index} 页", logo_path, font)
        ax.text(
            0.055,
            0.855,
            "表格说明：候选谷值按频率顺序列出，用于复核是否存在多个谐振点；S11/RL 合并显示。\n绝对 -10dB 与相对 +3dB 带宽用于快速筛选可用谷值。",
            transform=ax.transAxes,
            fontsize=8.1,
            fontproperties=font,
            color="#60717a",
            linespacing=1.45,
        )
        rows = [["#", "频率 MHz", "S11 / RL", "VSWR", "绝对-10dB带宽", "相对+3dB带宽"]]
        for index, valley in enumerate(valleys[start:start + chunk_size], start=start + 1):
            abs10 = (valley.get("bandwidths") or {}).get("absolute_10db") or {}
            rel3 = (valley.get("bandwidths") or {}).get("relative_3db") or {}
            rows.append(
                [
                    str(index),
                    f"{valley.get('frequency_mhz', 0):.6f}",
                    _format_s11_rl(valley.get("s11_db")),
                    _format_vswr(valley.get("vswr")),
                    format_hz(abs10.get("width_hz")),
                    format_hz(rel3.get("width_hz")),
                ]
            )
        table_height = min(0.66, max(0.14, 0.034 * len(rows)))
        _draw_report_table(ax, rows, [0.055, 0.80 - table_height, 0.89, table_height], font, font_size=7.0, header=True)
        pages.append(fig)
    return pages

def _draw_report_table(ax, rows, bbox, font, font_size=8.0, header=False):
    table = ax.table(cellText=rows, bbox=bbox, cellLoc="left")
    table.auto_set_font_size(False)
    table.set_fontsize(font_size)
    for (row, _col), cell in table.get_celld().items():
        cell.set_edgecolor("#ccd6d8")
        cell.set_linewidth(0.55)
        cell.PAD = 0.025
        if font:
            cell.get_text().set_fontproperties(font)
        cell.get_text().set_wrap(True)
        cell.get_text().set_va("center")
        if header and row == 0:
            cell.set_facecolor("#1f393f")
            cell.get_text().set_color("#fff7e8")
        elif row % 2 == 0:
            cell.set_facecolor("#f6f0e2")
        else:
            cell.set_facecolor("#ffffff")
    return table


def _annotate_report_points(ax, label_points, font):
    """Place compact point labels while trying several offsets to avoid overlaps."""
    if not label_points:
        return
    order = {
        "center": 0,
        "target": 1,
        "absolute_3db_left": 2,
        "absolute_3db_right": 3,
        "absolute_10db_left": 4,
        "absolute_10db_right": 5,
    }
    occupied = []
    dpi_scale = ax.figure.dpi / 72.0
    try:
        axis_bounds = ax.get_window_extent().bounds
    except Exception:
        axis_bounds = None

    for point in sorted(label_points, key=lambda item: order.get(item.get("type"), 99)):
        if point.get("x") is None or point.get("y") is None or not point.get("text"):
            continue
        candidates = _annotation_offsets(point.get("type"))
        text_lines = str(point["text"]).split("\n")
        text_width = max(54, max(len(line) for line in text_lines) * 5.8)
        text_height = 12 * len(text_lines) + 8
        anchor_x, anchor_y = ax.transData.transform((point["x"], point["y"]))
        chosen = candidates[-1]
        for dx_pt, dy_pt in candidates:
            dx = dx_pt * dpi_scale
            dy = dy_pt * dpi_scale
            rect = _annotation_rect(anchor_x + dx, anchor_y + dy, text_width, text_height, dx_pt, dy_pt)
            if axis_bounds and not _rect_inside(rect, axis_bounds, margin=2):
                continue
            if any(_rects_overlap(rect, prev) for prev in occupied):
                continue
            chosen = (dx_pt, dy_pt)
            occupied.append(rect)
            break
        else:
            dx = chosen[0] * dpi_scale
            dy = chosen[1] * dpi_scale
            occupied.append(_annotation_rect(anchor_x + dx, anchor_y + dy, text_width, text_height, chosen[0], chosen[1]))

        ax.annotate(
            point["text"],
            (point["x"], point["y"]),
            textcoords="offset points",
            xytext=chosen,
            fontsize=6.7,
            fontproperties=font,
            color="#172026",
            ha="left" if chosen[0] >= 0 else "right",
            va="bottom" if chosen[1] >= 0 else "top",
            bbox={"boxstyle": "round,pad=0.28", "fc": "#fff8e8", "ec": point.get("color", "#60717a"), "lw": 0.8, "alpha": 0.94},
            arrowprops={"arrowstyle": "-", "color": point.get("color", "#60717a"), "lw": 0.7, "shrinkA": 0, "shrinkB": 4},
            zorder=7,
        )


def _annotation_offsets(point_type):
    left_first = [(-12, 16), (-12, -24), (-68, 22), (-68, -30), (14, 28), (14, -36)]
    right_first = [(12, 16), (12, -24), (68, 22), (68, -30), (-14, 28), (-14, -36)]
    center_first = [(14, 16), (-14, 18), (14, -30), (-14, -32), (62, 24), (-62, 24)]
    if point_type and point_type.endswith("_left"):
        return left_first
    if point_type and point_type.endswith("_right"):
        return right_first
    if point_type == "target":
        return [(-14, 18), (-14, -30), (14, 18), (14, -30), (-66, 24), (66, 24)]
    return center_first


def _annotation_rect(anchor_x, anchor_y, width, height, dx_pt, dy_pt):
    left = anchor_x if dx_pt >= 0 else anchor_x - width
    bottom = anchor_y if dy_pt >= 0 else anchor_y - height
    return (left, bottom, left + width, bottom + height)


def _rect_inside(rect, bounds, margin=0):
    left, bottom, right, top = rect
    bx, by, bw, bh = bounds
    return left >= bx + margin and right <= bx + bw - margin and bottom >= by + margin and top <= by + bh - margin


def _rects_overlap(a, b, padding=3):
    return not (a[2] + padding < b[0] or b[2] + padding < a[0] or a[3] + padding < b[1] or b[3] + padding < a[1])


def _report_marker_label(point, value_key="s11"):
    short = _short_marker_label(point.get("type"), point.get("label"))
    freq = point.get("frequency_mhz")
    if value_key == "vswr":
        value = f"VSWR {_format_vswr(point.get('vswr'))}"
    else:
        value = f"S11 {_format_db(point.get('s11_db'))}"
    return f"{short} {freq:.3f}MHz\n{value}" if freq is not None else f"{short}\n{value}"


def _report_smith_label(marker):
    short = _short_marker_label(marker.get("type"), marker.get("label"))
    impedance = marker.get("impedance_label") or "--"
    return f"{short}\n{impedance}"


def _short_marker_label(point_type, fallback=None):
    labels = {
        "center": "中心",
        "target": "理想",
        "absolute_3db_left": "-3L",
        "absolute_3db_right": "-3R",
        "absolute_10db_left": "-10L",
        "absolute_10db_right": "-10R",
    }
    return labels.get(point_type, fallback or point_type or "")


def _draw_smith_impedance_grid(ax, reference_ohm, font):
    try:
        from matplotlib.patches import Circle
        ax.add_artist(Circle((0, 0), 1, color="#1f393f", fill=False, linewidth=1.0))
    except Exception:
        pass

    ax.axhline(0, color="#1f393f", linewidth=0.8, alpha=0.35)
    ax.axvline(0, color="#1f393f", linewidth=0.8, alpha=0.18)
    grid_color = "#60717a"
    resistance_values = [0, 0.2, 0.5, 1, 2, 5]
    reactance_values = [0.2, 0.5, 1, 2, 5]
    r_axis = [i * 0.04 for i in range(0, 501)]
    x_axis = [-10 + i * 0.04 for i in range(0, 501)]

    for resistance in resistance_values:
        points = [_gamma_from_normalized_impedance(resistance, reactance) for reactance in x_axis]
        _plot_gamma_curve(ax, points, grid_color, alpha=0.22)
        label_gamma = _gamma_from_normalized_impedance(resistance, 0)
        if label_gamma:
            label = "短路" if resistance == 0 else f"{resistance * reference_ohm:g}Ω"
            ax.text(label_gamma.real, -0.055, label, fontsize=6.5, color=grid_color, fontproperties=font, ha="center", va="top")

    for reactance in reactance_values:
        for sign, va in ((1, "bottom"), (-1, "top")):
            value = sign * reactance
            points = [_gamma_from_normalized_impedance(resistance, value) for resistance in r_axis]
            _plot_gamma_curve(ax, points, grid_color, alpha=0.18)
            label_gamma = _gamma_from_normalized_impedance(0.2, value)
            if label_gamma:
                label = f"{'+' if sign > 0 else '-'}j{reactance * reference_ohm:g}Ω"
                ax.text(label_gamma.real, label_gamma.imag, label, fontsize=6.5, color=grid_color, fontproperties=font, ha="left", va=va)

    ax.text(-0.98, -0.98, f"Z0={reference_ohm:g}Ω", fontsize=7.2, color="#455a64", fontproperties=font)


def _gamma_from_normalized_impedance(resistance, reactance):
    denominator = complex(resistance + 1.0, reactance)
    if abs(denominator) < 1e-12:
        return None
    gamma = (complex(resistance, reactance) - 1.0) / denominator
    if abs(gamma) > 1.0001:
        return None
    return gamma


def _plot_gamma_curve(ax, points, color, alpha=0.2):
    valid = [point for point in points if point is not None and abs(point) <= 1.0001]
    if len(valid) < 2:
        return
    ax.plot([point.real for point in valid], [point.imag for point in valid], color=color, linewidth=0.55, alpha=alpha)


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


def _format_mhz_delta(value):
    if value is None:
        return "--"
    value = float(value)
    return f"{value:+.3f} MHz"


def _format_db(value):
    if value is None:
        return "--"
    return f"{float(value):.3f} dB"


def _format_s11_rl(s11_db):
    if s11_db is None:
        return "--"
    return f"{float(s11_db):.3f} / {-float(s11_db):.3f} dB"


def _format_vswr(value):
    if value is None:
        return "∞"
    return f"{float(value):.3f}"


def _vswr_plot_cap(values):
    finite_values = [float(value) for value in values if value is not None and math.isfinite(float(value))]
    if not finite_values:
        return 6.0
    return max(2.2, min(10.0, max(finite_values) + 0.4, vswr_from_s11_db(-3) + 0.25))


def format_impedance(resistance_ohm, reactance_ohm):
    if resistance_ohm is None or reactance_ohm is None:
        return "--"
    sign = "+" if float(reactance_ohm) >= 0 else "-"
    return f"{float(resistance_ohm):.2f} {sign} j{abs(float(reactance_ohm)):.2f} Ω"


def _format_percent(value):
    if value is None:
        return "--"
    value = float(value)
    return f"{value:+.0f}%"


def _format_frequency_delta(target_summary):
    if not target_summary:
        return "--"
    error = target_summary.get("frequency_error_mhz")
    percent = target_summary.get("frequency_error_percent")
    if error is None or percent is None:
        return "--"
    return f"{_format_mhz_delta(error)} ({float(percent):+.2f}%)"


def _format_target_comparison_note(target_summary):
    if not target_summary:
        return "理想频点：全扫宽或未配置目标频率时不做理想频点对比。"
    delta = _format_frequency_delta(target_summary)
    rl_delta = _format_db(target_summary.get("return_loss_delta_db"))
    status = target_summary.get("status_label") or "--"
    return f"理想/实际：{status}；频偏 {delta}；RL差 {rl_delta}（正值=实际谷值更深）。"


def _format_endpoint_cell(bw, side):
    return "\n".join(
        [
            format_hz(bw.get(f"{side}_hz")),
            _format_s11_rl(bw.get(f"{side}_s11_db")),
            f"VSWR {_format_vswr(bw.get(f'{side}_vswr'))}",
        ]
    )
