# n9918a_backend.py
try:
    import pyvisa
except ImportError:
    class _PyVisaStub:
        class errors:
            VisaIOError = Exception

        @staticmethod
        def ResourceManager(*_args, **_kwargs):
            raise RuntimeError("pyvisa is not installed. Run `pip install -r requirements.txt`.")

    pyvisa = _PyVisaStub()
import matplotlib.pyplot as plt
import numpy as np
import time
import csv
import os
import math
from datetime import datetime
try:
    from scipy import signal
except ImportError:
    signal = None

# 在文件开头添加平台检测
import platform

# 设置matplotlib字体 - Mac兼容版本
if platform.system() == "Darwin":  # macOS
    plt.rcParams['font.sans-serif'] = ['PingFang SC', 'Heiti SC', 'Arial Unicode MS', 'DejaVu Sans']
else:  # Windows/Linux
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']

plt.rcParams['axes.unicode_minus'] = False


DEFAULT_SA_CORRECTIONS = {
    "cable_loss_db": 0.0,
    "antenna_factor_db": 0.0,
    "switch_loss_db": 0.0,
    "external_preamp_gain_db": 0.0,
}

SCREENING_LIMIT_NOTE = (
    "Screening reference only. Formal EMC judgment must confirm detector, "
    "distance, antenna/cable/switch corrections, site setup, and standard version."
)


def dbm_to_dbuv(dbm_value, impedance_ohm=50.0):
    """Convert dBm to dB microvolts for a resistive system."""
    return float(dbm_value) + 90.0 + 10.0 * math.log10(float(impedance_ohm))


def dbuv_to_dbm(dbuv_value, impedance_ohm=50.0):
    """Convert dB microvolts to dBm for a resistive system."""
    return float(dbuv_value) - 90.0 - 10.0 * math.log10(float(impedance_ohm))


def dbuv_to_microvolts(dbuv_value):
    return 10 ** (float(dbuv_value) / 20.0)


def microvolts_to_dbuv(microvolts):
    return 20.0 * math.log10(max(float(microvolts), 1e-12))


def linear_average_dbuv(values):
    """Average dBuV samples in voltage domain, not directly in dB domain."""
    if not values:
        return 0.0
    linear_values = [dbuv_to_microvolts(value) for value in values]
    return microvolts_to_dbuv(sum(linear_values) / len(linear_values))


def _log_interpolate_limit(freq_mhz, start_mhz, stop_mhz, start_limit, stop_limit):
    if freq_mhz <= start_mhz:
        return float(start_limit)
    if freq_mhz >= stop_mhz:
        return float(stop_limit)
    ratio = (math.log10(freq_mhz) - math.log10(start_mhz)) / (
        math.log10(stop_mhz) - math.log10(start_mhz)
    )
    return float(start_limit) + ratio * (float(stop_limit) - float(start_limit))


def _select_detector_limit(detector_type, quasi_peak, average=None, peak=None):
    detector = (detector_type or "QUASI_PEAK").upper()
    if detector in {"AVERAGE", "AVG", "EAV"} and average is not None:
        return float(average), "AVERAGE"
    if detector in {"PEAK", "PK", "POSITIVE"} and peak is not None:
        return float(peak), "PEAK"
    return float(quasi_peak), "QUASI_PEAK"


def get_emission_limit_info(freq_hz, detector_type="QUASI_PEAK"):
    """
    Return screening FCC/CE limit metadata for the current frequency.

    The returned limits are intentionally labelled as screening references:
    conducted ranges are in dBuV at the receiver input, while radiated ranges
    are in dBuV/m and require the correction chain to include antenna factor.
    """
    freq_mhz = float(freq_hz) / 1e6
    detector = (detector_type or "QUASI_PEAK").upper()

    if 0.150 <= freq_mhz < 30.0:
        if freq_mhz < 0.500:
            fcc_qp = _log_interpolate_limit(freq_mhz, 0.150, 0.500, 66.0, 56.0)
            fcc_avg = _log_interpolate_limit(freq_mhz, 0.150, 0.500, 56.0, 46.0)
        elif freq_mhz < 5.0:
            fcc_qp, fcc_avg = 56.0, 46.0
        else:
            fcc_qp, fcc_avg = 60.0, 50.0
        ce_qp, ce_avg = fcc_qp, fcc_avg
        fcc_limit, fcc_detector = _select_detector_limit(detector, fcc_qp, average=fcc_avg)
        ce_limit, ce_detector = _select_detector_limit(detector, ce_qp, average=ce_avg)
        return {
            "fcc_limit": fcc_limit,
            "ce_limit": ce_limit,
            "fcc_detector": fcc_detector,
            "ce_detector": ce_detector,
            "unit": "dBuV",
            "measurement_type": "conducted_mains_screening",
            "distance_m": None,
            "fcc_source": "FCC Part 15 Class B conducted mains screening",
            "ce_source": "EN 55032/CISPR 32 Class B conducted mains screening",
            "note": SCREENING_LIMIT_NOTE,
        }

    if 30.0 <= freq_mhz < 1000.0:
        if freq_mhz < 88.0:
            fcc_limit = 40.0
        elif freq_mhz < 216.0:
            fcc_limit = 43.5
        elif freq_mhz < 960.0:
            fcc_limit = 46.0
        else:
            fcc_limit = 54.0
        ce_limit = 40.0 if freq_mhz < 230.0 else 47.0
        return {
            "fcc_limit": fcc_limit,
            "ce_limit": ce_limit,
            "fcc_detector": "QUASI_PEAK",
            "ce_detector": "QUASI_PEAK",
            "unit": "dBuV/m",
            "measurement_type": "radiated_3m_screening",
            "distance_m": 3.0,
            "fcc_source": "FCC Part 15 Class B radiated 3m screening",
            "ce_source": "EN 55032/CISPR 32 Class B radiated 3m screening",
            "note": SCREENING_LIMIT_NOTE,
        }

    if 1000.0 <= freq_mhz <= 18000.0:
        fcc_limit, fcc_detector = _select_detector_limit(detector, 54.0, average=54.0, peak=74.0)
        ce_limit, ce_detector = _select_detector_limit(detector, 50.0, average=50.0, peak=70.0)
        return {
            "fcc_limit": fcc_limit,
            "ce_limit": ce_limit,
            "fcc_detector": fcc_detector,
            "ce_detector": ce_detector,
            "unit": "dBuV/m",
            "measurement_type": "radiated_3m_screening_above_1ghz",
            "distance_m": 3.0,
            "fcc_source": "FCC Part 15 Class B radiated above 1GHz screening",
            "ce_source": "EN 55032/CISPR 32 Class B radiated above 1GHz screening",
            "note": SCREENING_LIMIT_NOTE,
        }

    return {
        "fcc_limit": 120.0,
        "ce_limit": 120.0,
        "fcc_detector": detector,
        "ce_detector": detector,
        "unit": "dBuV",
        "measurement_type": "out_of_screening_scope",
        "distance_m": None,
        "fcc_source": "out of configured screening range",
        "ce_source": "out of configured screening range",
        "note": SCREENING_LIMIT_NOTE,
    }


def collapse_contiguous_indices(indices, amplitudes):
    """Collapse adjacent exceeding bins to the highest-amplitude representative."""
    if not indices:
        return []
    sorted_indices = sorted(set(int(index) for index in indices))
    groups = []
    current = [sorted_indices[0]]
    for index in sorted_indices[1:]:
        if index == current[-1] + 1:
            current.append(index)
        else:
            groups.append(current)
            current = [index]
    groups.append(current)
    return [max(group, key=lambda idx: amplitudes[idx]) for group in groups]

class N9918AController:
    """
    N9918A FieldFox SA Controller for EMC Testing
    """
    
    # 预设参数配置
    PRESET_CONFIGS = {
        "EMC_30MHz_1GHz": {
            "name": "EMC测试 (30MHz-1GHz)",
            "start_freq": 30e6,
            "stop_freq": 1e9,
            "n_points": 2001,
            "rbw": 100e3,
            "vbw": 100e3,
            "description": "标准EMC测试参数"
        },
        "LF_9kHz_150kHz": {
            "name": "低频测试 (9kHz-150kHz)",
            "start_freq": 9e3,
            "stop_freq": 150e3,
            "n_points": 1001,
            "rbw": 200,
            "vbw": 1e3,
            "description": "传导发射测试"
        },
        "MF_150kHz_30MHz": {
            "name": "中频测试 (150kHz-30MHz)",
            "start_freq": 150e3,
            "stop_freq": 30e6,
            "n_points": 1501,
            "rbw": 10e3,
            "vbw": 30e3,
            "description": "传导发射测试"
        },
        "HF_1GHz_3GHz": {
            "name": "高频测试 (1GHz-3GHz)",
            "start_freq": 1e9,
            "stop_freq": 3e9,
            "n_points": 1001,
            "rbw": 1e6,
            "vbw": 3e6,
            "description": "辐射发射测试"
        }
    }
    
    def __init__(self, ip_address='192.168.0.124', timeout=10000):
        self.ip_address = ip_address
        self.timeout = timeout
        self.rm = None
        self.device = None
        self.connected = False
        self.start_freq = None
        self.stop_freq = None
        self.n_points = None
        self.rbw = None
        self.vbw = None
        self.current_config = None
        self.amplitude_unit = "DBUV"
        self.sa_corrections = DEFAULT_SA_CORRECTIONS.copy()
        self.last_scpi_errors = []
        
    def connect(self):
        try:
            self.rm = pyvisa.ResourceManager()
            self.device = self.rm.open_resource(f"TCPIP0::{self.ip_address}::inst0::INSTR")
            self.device.timeout = self.timeout
            
            self.device.write("*CLS")
            device_id = self.device.query("*IDN?")
            print(f"Connected to: {device_id}")
            
            # 官方示例使用 *OPC? 等待模式切换完成，避免后续配置命令跑在旧模式下。
            self.device.query("INST:SEL 'SA';*OPC?")
            
            self.connected = True
            print("Successfully connected to N9918A")
            return True
            
        except Exception as e:
            print(f"ERROR: Unable to connect to device - {e}")
            self.connected = False
            return False
    
    def disconnect(self):
        if self.device:
            self.device.close()
        if self.rm:
            self.rm.close()
        self.connected = False
        print("Disconnected from N9918A")

    def set_sa_corrections(self, **corrections):
        """Set receiver-to-limit correction terms used by screening results."""
        for key in DEFAULT_SA_CORRECTIONS:
            if key in corrections and corrections[key] is not None:
                self.sa_corrections[key] = float(corrections[key])
        return self.sa_corrections.copy()

    def correction_total_db(self):
        return (
            self.sa_corrections.get("cable_loss_db", 0.0)
            + self.sa_corrections.get("antenna_factor_db", 0.0)
            + self.sa_corrections.get("switch_loss_db", 0.0)
            - self.sa_corrections.get("external_preamp_gain_db", 0.0)
        )

    def _write_optional(self, command, label):
        try:
            self.device.write(command)
        except Exception as exc:
            print(f"[WARN] {label} command failed ({command}): {exc}")

    def _check_scpi_errors(self, context):
        errors = []
        for _ in range(5):
            try:
                response = str(self.device.query("SYST:ERR?")).strip()
            except Exception as exc:
                print(f"[WARN] 无法读取 SCPI 错误队列 ({context}): {exc}")
                break
            if response.startswith("+0") or response.startswith("0,") or "No error" in response:
                break
            errors.append(response)
        self.last_scpi_errors = errors
        if errors:
            print(f"[WARN] SCPI 错误 ({context}): {' | '.join(errors)}")
        return errors

    def _reset_trace_for_new_sweep(self):
        """Keep trace 1 in Clear/Write so each triggered sweep overwrites history."""
        self.device.write(":TRAC1:TYPE CLRW")
        self._write_optional(":TRAC2:TYPE BLAN", "blank trace 2")
        self._write_optional(":TRAC3:TYPE BLAN", "blank trace 3")
        self._write_optional(":TRAC4:TYPE BLAN", "blank trace 4")
        self._write_optional(":INIT:REST", "restart trace averaging")

    def clear_sa_display_state(self, blank_trace=True):
        """Clear SA status and remove old traces from the FieldFox display."""
        if not self.connected:
            print("ERROR: Device not connected")
            return False

        try:
            self.device.write("*CLS")
            self.device.query("INST:SEL 'SA';*OPC?")
            self.device.write("INIT:CONT OFF")
            self.device.write(":SENS:AMPL:UNIT DBUV")
            self.amplitude_unit = "DBUV"
            self._reset_trace_for_new_sweep()
            if blank_trace:
                self.device.write(":TRAC1:TYPE BLAN")
            errors = self._check_scpi_errors("SA manual clear")
            return not errors
        except Exception as e:
            print(f"ERROR: Failed to clear SA display state - {e}")
            return False

    def _prepare_sa_screening_trace(self, clear_status=False):
        if clear_status:
            self.device.write("*CLS")
        self.device.query("INST:SEL 'SA';*OPC?")
        self.device.write("INIT:CONT OFF")
        self.device.write(":SENS:AMPL:UNIT DBUV")
        self.amplitude_unit = "DBUV"
        self._write_optional(":SENS:BAND:RES:AUTO OFF", "manual RBW")
        self._write_optional(":SENS:BAND:VID:AUTO OFF", "manual VBW")
        self.device.write(":SENS:DET POS")
        self._write_optional(":SENS:AVER:COUN 1", "disable averaging")
        self._reset_trace_for_new_sweep()

    def _query_float(self, command, default=None):
        try:
            value = float(str(self.device.query(command)).strip())
            if math.isfinite(value):
                return value
        except Exception:
            pass
        return default

    def _refresh_actual_sa_settings(self):
        start_freq = self._query_float(":SENS:FREQ:STAR?", self.start_freq)
        stop_freq = self._query_float(":SENS:FREQ:STOP?", self.stop_freq)
        n_points = self._query_float(":SENS:SWE:POIN?", self.n_points)
        rbw = self._query_float(":SENS:BAND:RES?", self.rbw)
        vbw = self._query_float(":SENS:BAND:VID?", self.vbw)
        if start_freq and stop_freq and stop_freq > start_freq:
            self.start_freq = start_freq
            self.stop_freq = stop_freq
        if n_points and n_points >= 2:
            self.n_points = int(round(n_points))
        if rbw and rbw > 0:
            self.rbw = rbw
        if vbw and vbw > 0:
            self.vbw = vbw

    def _estimate_sweep_time(self):
        sweep_time = self._query_float(":SENS:SWE:TIME?", None)
        if sweep_time and sweep_time > 0:
            return sweep_time
        if self.start_freq and self.stop_freq:
            return max(0.5, (self.stop_freq - self.start_freq) / 1e9 * 3.0)
        return 1.0

    def _parse_numeric_csv(self, data):
        return [float(item.strip()) for item in str(data).replace("\n", "").split(",") if item.strip()]

    def _build_frequency_axis(self):
        try:
            x_values = self._parse_numeric_csv(self.device.query(":TRAC1:XVAL?"))
            if self.n_points and len(x_values) == self.n_points:
                return x_values
        except Exception:
            pass
        if not self.start_freq or not self.stop_freq or not self.n_points or self.n_points < 2:
            raise ValueError("SA 频率轴参数不完整，请先配置 start/stop/points。")
        freq_step = (self.stop_freq - self.start_freq) / (self.n_points - 1)
        return [self.start_freq + i * freq_step for i in range(self.n_points)]

    def _read_trace_amplitudes_dbuv(self):
        self.device.write(":TRAC1:DATA?")
        trace_data = self.device.read()
        amplitudes = self._parse_numeric_csv(trace_data)
        if self.n_points and len(amplitudes) != self.n_points:
            raise ValueError(f"SA trace 点数不匹配：期望 {self.n_points}，实际 {len(amplitudes)}。")
        correction_total = self.correction_total_db()
        if correction_total:
            amplitudes = [value + correction_total for value in amplitudes]
        return amplitudes

    def acquire_single_trace(self, reset_trace=True):
        if reset_trace:
            self._reset_trace_for_new_sweep()
        self.device.write("INIT:CONT OFF")
        try:
            self.device.query(":INIT:IMM;*OPC?")
        except Exception:
            self.device.write(":INIT:IMM")
            wait_time = max(self._estimate_sweep_time() * 1.2, 1.0)
            print(f"[WAIT] 等待扫描完成 ({wait_time:.1f}秒)...")
            time.sleep(wait_time)
        frequencies = self._build_frequency_axis()
        amplitudes = self._read_trace_amplitudes_dbuv()
        return frequencies, amplitudes

    def has_emi_option(self):
        try:
            catalog = str(self.device.query("INST:CAT?")).upper()
            return "EMI" in catalog
        except Exception:
            return False
    
    def configure_settings(self, config_name):
        """
        根据预设配置名称配置设备
        """
        if config_name not in self.PRESET_CONFIGS:
            print(f"ERROR: Configuration '{config_name}' not found")
            return False
            
        config = self.PRESET_CONFIGS[config_name]
        return self._configure_device(
            config["start_freq"],
            config["stop_freq"],
            config["n_points"],
            config["rbw"],
            config["vbw"],
            config_name
        )
    
    def _configure_device(self, start_freq, stop_freq, n_points, rbw, vbw, config_name):
        """
        内部配置设备方法
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return False
            
        try:
            print(f"[CONFIG] 配置设备参数: {config_name}")
            self._prepare_sa_screening_trace(clear_status=True)
            
            # Set frequency range
            self.device.write(f":SENS:FREQ:STAR {start_freq}")
            time.sleep(0.2)
            self.device.write(f":SENS:FREQ:STOP {stop_freq}")
            time.sleep(0.2)
            print(f"[FREQ] 频率范围: {start_freq/1e6:.3f}MHz - {stop_freq/1e9:.3f}GHz")
            
            # Set number of points
            self.device.write(f":SENS:SWE:POIN {n_points}")
            time.sleep(0.2)
            print(f"[POINTS] 采样点数: {n_points}")
            
            # Set RBW and VBW
            self._write_optional(":SENS:BAND:RES:AUTO OFF", "manual RBW")
            self.device.write(f":SENS:BAND:RES {rbw}")
            time.sleep(0.5)
            self._write_optional(":SENS:BAND:VID:AUTO OFF", "manual VBW")
            self.device.write(f":SENS:BAND:VID {vbw}")
            time.sleep(0.5)
            print(f"[BAND]  RBW: {rbw}Hz, VBW: {vbw}Hz")
            
            # Positive detector is the safer screening prescan choice for narrow peaks.
            self.device.write(":SENS:DET POS")
            time.sleep(0.2)
            print("[DETECTOR] Detector: Positive Peak")
            
            # Set Internal Amplifier ON
            self.device.write(":SENS:POW:GAIN:STAT ON")
            time.sleep(0.2)
            print("[GAIN] 内部放大器: ON")
            
            # Set Internal Attenuator to 0dB
            self.device.write(":SENS:POW:ATT 0")
            time.sleep(0.2)
            print("[ATT] 内部衰减器: 0dB")
            
            # Store parameters
            self.start_freq = start_freq
            self.stop_freq = stop_freq
            self.n_points = n_points
            self.rbw = rbw
            self.vbw = vbw
            self.current_config = config_name
            self.amplitude_unit = "DBUV"
            self._reset_trace_for_new_sweep()
            self._refresh_actual_sa_settings()
            self._check_scpi_errors("SA configure")
            
            print("[OK] 参数配置完成! (连续扫描已暂停)")
            return True
            
        except Exception as e:
            print(f"ERROR: Failed to configure measurement - {e}")
            return False
    
    def read_trace_data(self):
        """
        Read trace data from the device
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return None, None
            
        try:
            self._prepare_sa_screening_trace(clear_status=True)
            frequencies, amplitudes_dBuv = self.acquire_single_trace(reset_trace=True)
            self._check_scpi_errors("SA single trace")
            return frequencies, amplitudes_dBuv
            
        except Exception as e:
            print(f"ERROR: Failed to read trace data - {e}")
            return None, None
    
    def get_preset_configs(self):
        """获取所有预设配置"""
        return self.PRESET_CONFIGS
    
    def get_current_status(self):
        """获取当前设备状态"""
        if not self.connected:
            return {"status": "未连接"}
        
        config_info = self.PRESET_CONFIGS.get(self.current_config, {}) if self.current_config else {}
        
        return {
            "status": "已连接",
            "ip_address": self.ip_address,
            "current_config": config_info.get("name", self.current_config) if self.current_config else "未配置",
            "start_freq": self.start_freq,
            "stop_freq": self.stop_freq,
            "n_points": self.n_points,
            "rbw": self.rbw,
            "vbw": self.vbw,
            "amplitude_unit": self.amplitude_unit,
            "sa_corrections": self.sa_corrections.copy(),
            "last_scpi_errors": list(self.last_scpi_errors),
        }


    
    def get_emc_measurement_fast(self, duration_seconds=15, should_stop=None):
        """
        快速EMC测量（采集时间序列数据，PC端计算多种模式）
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return {}
        
        total_start_time = time.time()
        print(f"[RUN] 开始快速EMC测量 ({duration_seconds} 秒)")
        print("=" * 50)
        
        try:
            # 1. 收集时间序列数据
            time_series_data = self.collect_emc_time_series(duration_seconds, should_stop=should_stop)
            
            if not time_series_data:
                print("[ERROR] 未能收集到时间序列数据")
                return {}
            
            collection_time = time.time() - total_start_time
            print(f"   [TIME]  数据采集用时: {collection_time:.1f} 秒")
            print(f"   [DATA]  实际采样次数: {len(time_series_data)}")
            
            # 2. PC端计算多种检测器模式
            print(f"\n[CALC] PC端计算EMC检测器模式...")
            calculation_start_time = time.time()
            
            results = {}
            detector_modes = ["PEAK", "QUASI_PEAK", "AVERAGE"]
            
            for mode in detector_modes:
                frequencies, amplitudes = calculate_emc_detector_modes(time_series_data, mode)
                if frequencies is not None and amplitudes is not None:
                    results[mode] = (frequencies, amplitudes)
                    print(f"   [OK] {mode} 模式计算完成")
                    if amplitudes:
                        max_val = max(amplitudes)
                        min_val = min(amplitudes)
                        avg_val = sum(amplitudes) / len(amplitudes)
                        print(f"       最大值: {max_val:.2f} dBμV, 最小值: {min_val:.2f} dBμV, 平均值: {avg_val:.2f} dBμV")
                else:
                    print(f"   [ERROR] {mode} 模式计算失败")
            
            calculation_time = time.time() - calculation_start_time
            total_time = time.time() - total_start_time
            
            # 添加采样数据（用于保存）
            results["sampling_data"] = time_series_data
            
            # 添加采样信息
            results["sampling_info"] = {
                "total_samples": len(time_series_data),
                "sample_duration": duration_seconds,
                "collection_time": collection_time,
                "calculation_time": calculation_time,
                "data_points": len(time_series_data[0]['amplitudes']) if time_series_data else 0,
                "rbw": self.rbw if self.rbw else 100e3,
                "start_time": time_series_data[0]['timestamp'] if time_series_data else 0,
                "end_time": time_series_data[-1]['timestamp'] if time_series_data else 0,
                "amplitude_unit": self.amplitude_unit,
                "corrections": self.sa_corrections.copy(),
                "correction_total_db": self.correction_total_db(),
                "screening_mode": True,
            }
            
            # 添加测量摘要
            results["measurement_summary"] = {
                "total_duration": total_time,
                "actual_measurement_time": duration_seconds,
                "data_points": len(time_series_data[0]['amplitudes']) if time_series_data else 0,
                "total_samples": len(time_series_data),
                "modes_computed": detector_modes,
                "measurement_time": time.strftime("%Y-%m-%d %H:%M:%S"),
                "screening_mode": True,
                "quasi_peak_estimated": True,
                "limit_note": SCREENING_LIMIT_NOTE,
            }
            results["detector_notes"] = {
                "PEAK": "Max of synchronized positive-peak SA sweeps.",
                "AVERAGE": "Voltage-domain screening average, not a formal EMI average detector.",
                "QUASI_PEAK": "Estimated in software from synchronized sweeps; use FieldFox EMI Option 361 for formal QPD.",
            }
            
            print(f"   [TIME]  计算用时: {calculation_time:.1f} 秒")
            print(f"[OK] 所有处理完成! 总用时: {total_time:.1f} 秒")
            
            return results
            
        except Exception as e:
            print(f"ERROR: 快速EMC测量失败 - {e}")
            import traceback
            traceback.print_exc()
            return {}

    def collect_emc_time_series(self, duration_seconds=15, should_stop=None):
        """
        稳定版时间序列数据采集 - 每个样本均等待完整单次 sweep。
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return []
        
        print(f"[LOOP] 开始时间序列数据采集 ({duration_seconds} 秒)")
        original_timeout = self.device.timeout
        try:
            self._prepare_sa_screening_trace(clear_status=True)
            sweep_time = self._estimate_sweep_time()
            sample_interval = max(sweep_time * 1.05, 0.25)
            max_samples = max(1, int(duration_seconds / sample_interval))
            self.device.timeout = max(original_timeout or 0, int((sweep_time + 10.0) * 1000))
            print(f"   [TIME]  仪器 sweep time: {sweep_time:.3f}s")
            print(f"   [TIME]  采样间隔: {sample_interval:.3f}s, 目标采样次数: {max_samples}")

            time_series_data = []
            start_time = time.time()
            sample_count = 0
            last_successful_time = start_time
            consecutive_failures = 0
            max_consecutive_failures = 3
            
            while time.time() - start_time < duration_seconds and sample_count < max_samples:
                if should_stop and should_stop():
                    print("   [STOP]  用户请求停止采样")
                    break

                sample_start_time = time.time()
                try:
                    frequencies, amplitudes = self.acquire_single_trace(reset_trace=True)
                    read_duration = time.time() - sample_start_time
                    if read_duration > max(sweep_time * 3, 8.0):
                        print(f"   [WARN]  单次 sweep+读取耗时异常: {read_duration:.2f}s")

                    sample_count += 1
                    current_time = time.time()
                    time_series_data.append(
                        {
                            'timestamp': current_time - start_time,
                            'frequencies': frequencies,
                            'amplitudes': amplitudes,
                        }
                    )
                    consecutive_failures = 0
                    last_successful_time = current_time

                    progress = (sample_count / max_samples) * 100
                    if sample_count % 5 == 0 or sample_count <= 10:
                        elapsed = current_time - start_time
                        remaining = max(0.0, duration_seconds - elapsed)
                        print(
                            f"   [DATA] 采样 #{sample_count}/{max_samples} ({progress:.1f}%) "
                            f"已用时: {elapsed:.1f}s, 剩余: {remaining:.1f}s"
                        )

                    next_sample_time = sample_start_time + sample_interval
                    sleep_time = max(0.0, next_sample_time - time.time())
                    while sleep_time > 0:
                        if should_stop and should_stop():
                            break
                        time.sleep(min(sleep_time, 0.5))
                        sleep_time = max(0.0, next_sample_time - time.time())

                except pyvisa.errors.VisaIOError as e:
                    consecutive_failures += 1
                    print(f"   [WARN]  VISA通信错误 (第{consecutive_failures}次): {e}")
                    self._check_scpi_errors("SA sampling VISA error")
                    if consecutive_failures >= max_consecutive_failures:
                        print(f"   [ERROR] 连续通信失败{consecutive_failures}次，停止采样")
                        break
                    time.sleep(1)

                except Exception as e:
                    consecutive_failures += 1
                    print(f"   [WARN]  采样失败 (第{consecutive_failures}次): {e}")
                    self._check_scpi_errors("SA sampling error")
                    if consecutive_failures >= max_consecutive_failures:
                        print(f"   [ERROR] 连续失败{consecutive_failures}次，停止采样")
                        break
                    time.sleep(0.5)
                
                # 检查是否长时间无响应
                if time.time() - last_successful_time > max(30.0, sweep_time * 5):
                    print(f"   [ERROR] 设备长时间无响应，停止采样")
                    break
            
            # 停止连续扫描
            try:
                self.device.write("INIT:CONT OFF")
                time.sleep(0.2)
                self.device.timeout = original_timeout  # 恢复原始超时
            except:
                print("   [WARN]  停止扫描时出现异常")
            
            print(f"[OK] 时间序列采集完成! 总采样: {len(time_series_data)} 次")
            
            if time_series_data:
                actual_duration = time_series_data[-1]['timestamp']
                print(f"   [DATA] 实际采样时长: {actual_duration:.1f}s")
                print(f"   [DATA] 平均采样间隔: {actual_duration/len(time_series_data):.2f}s")
            
            return time_series_data
            
        except Exception as e:
            print(f"ERROR: 时间序列采集失败 - {e}")
            try:
                self.device.write("INIT:CONT OFF")
                self.device.timeout = original_timeout
            except:
                pass
            return []

def calculate_emc_detector_modes(time_series_data, detector_type="QUASI_PEAK"):
    """
    改进的EMC检测器模式计算
    """
    if not time_series_data:
        return None, None
    
    print(f"   [DETECTOR] 计算 {detector_type} 模式...")
    
    frequencies = time_series_data[0]['frequencies']
    n_points = len(frequencies)
    n_samples = len(time_series_data)
    
    if n_samples == 0 or n_points == 0:
        return None, None
    
    print(f"       数据维度: {n_samples} 次采样 × {n_points} 个频率点")
    
    result_amplitudes = []
    
    for freq_idx in range(n_points):
        # 收集该频率点的所有采样值
        time_values = []
        for sample in time_series_data:
            if freq_idx < len(sample['amplitudes']):
                time_values.append((sample['timestamp'], sample['amplitudes'][freq_idx]))
        
        if not time_values:
            result_amplitudes.append(0)
            continue
        
        times = [tv[0] for tv in time_values]
        values = [tv[1] for tv in time_values]
        
        # 计算检测器值
        if detector_type == "PEAK":
            detector_value = max(values)
        elif detector_type == "QUASI_PEAK":
            # 软件估算准峰值，正式 QPD 应优先使用仪器 EMI Option 361。
            current_freq = frequencies[freq_idx]
            detector_value = calculate_quasi_peak_value(times, values, current_freq)
        elif detector_type == "AVERAGE":
            detector_value = linear_average_dbuv(values)
        else:
            detector_value = values[-1] if values else 0
        
        result_amplitudes.append(detector_value)
    
    return frequencies, result_amplitudes


def calculate_quasi_peak_value(times, values, frequency_hz=None):
    """
    软件准峰值估算 - 根据频率选择时间常数，并在电压域积分。
    """
    if len(times) <= 1:
        return values[0] if values else 0
    
    # 根据CISPR 16标准选择时间常数（根据频率）
    if frequency_hz is not None:
        freq_mhz = frequency_hz / 1e6
        if freq_mhz < 0.15:  # 150kHz以下
            rise_time = 45e-3    # 45ms
            decay_time = 500e-3  # 500ms
        elif freq_mhz < 30:  # 150kHz - 30MHz
            rise_time = 1e-3     # 1ms
            decay_time = 160e-3  # 160ms
        else:  # 30MHz以上
            rise_time = 1e-3     # 1ms
            decay_time = 550e-6  # 550μs
    else:
        # 默认使用中频参数
        rise_time = 1e-3     # 1ms
        decay_time = 160e-3  # 160ms
    
    # 数据预处理和排序：dBuV 先转线性电压，避免在 dB 域积分。
    time_value_pairs = [
        (float(t), dbuv_to_microvolts(v))
        for t, v in zip(times, values)
    ]
    time_value_pairs.sort(key=lambda x: x[0])
    
    # 计算数据的基本统计信息
    all_values = [v for t, v in time_value_pairs]
    avg_value = sum(all_values) / len(all_values)
    max_value = max(all_values)
    
    # 初始化准峰值为第一个值
    qp_value = time_value_pairs[0][1]
    
    # 逐步计算准峰值
    for i in range(1, len(time_value_pairs)):
        dt = time_value_pairs[i][0] - time_value_pairs[i-1][0]
        current_value = time_value_pairs[i][1]
        
        # 限制时间间隔范围
        if dt <= 0 or dt > 10.0:  # 跳过异常时间间隔
            continue
        
        if current_value > qp_value:
            # 上升过程：快速跟踪较大值
            alpha = 1 - np.exp(-dt / rise_time)
            qp_value = qp_value + alpha * (current_value - qp_value)
        else:
            # 下降过程：缓慢衰减
            decay_factor = np.exp(-dt / decay_time)
            decayed_value = qp_value * decay_factor
            
            # 准峰值不应该低于当前值，也不应该低于平均值的某个比例
            min_allowed = max(current_value, avg_value * 0.7)  # 不低于平均值的70%
            qp_value = max(decayed_value, min_allowed)
    
    # 最终约束：筛查估算值保持在平均与峰值之间。
    qp_value = max(qp_value, avg_value * 0.8)
    qp_value = min(qp_value, max_value)
    
    return microvolts_to_dbuv(qp_value)

def save_emi_measurement_data(frequencies_dict, filename_prefix=None):
    """
    保存EMI测量数据（包含所有采样数据）
    frequencies_dict: 包含所有检测器模式数据的字典
    """
    import json
    from datetime import datetime
    
    measurement_folder = 'measurement_data'
    if not os.path.exists(measurement_folder):
        os.makedirs(measurement_folder)
    
    if filename_prefix is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename_prefix = f"emi_measurement_{timestamp}"
    
    saved_files = []
    
    # 保存每种模式的最终数据
    for mode, data in frequencies_dict.items():
        if mode == "measurement_summary" or mode == "sampling_info":
            continue
            
        if isinstance(data, tuple) and len(data) >= 2:
            frequencies, amplitudes = data[0], data[1]
            
            # 保存CSV格式的最终数据
            csv_filename = f"{filename_prefix}_{mode}_final.csv"
            csv_filepath = os.path.join(measurement_folder, csv_filename)
            
            with open(csv_filepath, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Frequency (Hz)', 'Amplitude (dBμV)'])
                for freq, amp in zip(frequencies, amplitudes):
                    writer.writerow([freq, amp])
            
            saved_files.append(csv_filepath)
            print(f"[SAVE] {mode} 最终数据已保存: {csv_filepath}")
    
    # 保存详细的采样数据（如果存在）
    if "sampling_data" in frequencies_dict:
        sampling_data = frequencies_dict["sampling_data"]
        
        # 保存每次采样的详细数据
        detailed_filename = f"{filename_prefix}_all_samples_detailed.csv"
        detailed_filepath = os.path.join(measurement_folder, detailed_filename)
        
        if sampling_data and len(sampling_data) > 0:
            with open(detailed_filepath, 'w', newline='') as f:
                writer = csv.writer(f)
                
                # 写入表头
                frequencies = sampling_data[0]['frequencies']
                header = ['Sample_Time(s)'] + [f'Freq_{i}_{freq/1e6:.2f}MHz' for i, freq in enumerate(frequencies)]
                writer.writerow(header)
                
                # 写入每次采样的数据
                for sample in sampling_data:
                    row = [f"{sample['timestamp']:.3f}"] + [f"{amp:.2f}" for amp in sample['amplitudes']]
                    writer.writerow(row)
            
            saved_files.append(detailed_filepath)
            print(f"[SAVE] 所有采样详细数据已保存: {detailed_filepath}")
            
            # 保存采样统计信息
            stats_filename = f"{filename_prefix}_sampling_statistics.csv"
            stats_filepath = os.path.join(measurement_folder, stats_filename)
            
            with open(stats_filepath, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Sample_Number', 'Timestamp(s)', 'Min_Value(dBμV)', 'Max_Value(dBμV)', 'Average_Value(dBμV)'])
                
                for i, sample in enumerate(sampling_data):
                    min_val = min(sample['amplitudes'])
                    max_val = max(sample['amplitudes'])
                    avg_val = sum(sample['amplitudes']) / len(sample['amplitudes'])
                    writer.writerow([i+1, f"{sample['timestamp']:.3f}", f"{min_val:.2f}", f"{max_val:.2f}", f"{avg_val:.2f}"])
            
            saved_files.append(stats_filepath)
            print(f"[SAVE] 采样统计信息已保存: {stats_filepath}")
    
    # 保存测量摘要
    if "measurement_summary" in frequencies_dict:
        summary_filename = f"{filename_prefix}_summary.json"
        summary_filepath = os.path.join(measurement_folder, summary_filename)
        
        with open(summary_filepath, 'w') as f:
            json.dump(frequencies_dict["measurement_summary"], f, indent=2)
        
        saved_files.append(summary_filepath)
        print(f"[SAVE] 测量摘要已保存: {summary_filepath}")
    
    # 保存采样信息
    if "sampling_info" in frequencies_dict:
        info_filename = f"{filename_prefix}_sampling_info.json"
        info_filepath = os.path.join(measurement_folder, info_filename)
        
        with open(info_filepath, 'w') as f:
            json.dump(frequencies_dict["sampling_info"], f, indent=2)
        
        saved_files.append(info_filepath)
        print(f"[SAVE] 采样信息已保存: {info_filepath}")
    
    return saved_files

def get_fcc_ce_limits(freq_hz, detector_type="QUASI_PEAK"):
    """
    获取筛查用 FCC 和 CE 参考限值。

    返回值保持兼容旧调用；完整口径请使用 get_emission_limit_info()。
    """
    info = get_emission_limit_info(freq_hz, detector_type=detector_type)
    return info["fcc_limit"], info["ce_limit"]

# 峰值检测函数
def find_peaks_manual(data, distance=5, prominence=3):
    """
    手动实现峰值检测
    """
    peaks = []
    n = len(data)
    
    for i in range(1, n-1):
        is_peak = True
        for j in range(max(0, i-distance), i):
            if data[j] >= data[i]:
                is_peak = False
                break
        if not is_peak:
            continue
        for j in range(i+1, min(n, i+distance+1)):
            if data[j] >= data[i]:
                is_peak = False
                break
        
        if is_peak and data[i] > np.mean(data) + prominence:
            peaks.append(i)
    
    peaks.sort(key=lambda x: data[x], reverse=True)
    return peaks

def post_process_peak_search(
    frequencies,
    amplitudes,
    peak_distance=30,
    min_prominence=2,
    detector_type="QUASI_PEAK",
):
    """
    改进的后处理峰值搜索 - 更智能的峰值检测算法
    """
    if len(amplitudes) < 3:
        return []
    
    # 计算动态参数
    mean_amp = np.mean(amplitudes)
    std_amp = np.std(amplitudes)
    
    # 动态调整显著性阈值
    dynamic_prominence = max(min_prominence, std_amp * 0.2)  # 进一步降低阈值以检测更多峰值
    
    # 动态调整最小高度阈值
    min_height = mean_amp + dynamic_prominence * 0.3  # 进一步降低阈值以检测更多峰值
    
    # 多级峰值检测；开发环境缺少 scipy 时退回手写算法，硬件环境仍建议安装 requirements。
    if signal:
        primary_peaks, _ = signal.find_peaks(
            amplitudes,
            distance=peak_distance,
            prominence=dynamic_prominence,
            height=min_height
        )
        secondary_peaks, _ = signal.find_peaks(
            amplitudes,
            distance=max(5, peak_distance // 4),  # 更小的间隔
            prominence=max(0.3, dynamic_prominence * 0.3),  # 更低的显著性要求
            height=mean_amp + 0.3  # 更低的高度要求
        )
    else:
        primary_peaks = find_peaks_manual(amplitudes, distance=peak_distance, prominence=dynamic_prominence)
        secondary_peaks = find_peaks_manual(
            amplitudes,
            distance=max(5, peak_distance // 4),
            prominence=max(0.3, dynamic_prominence * 0.3),
        )
    
    # 第三级：检测超限连续区域，并用区域内最高点代表，避免一整段超限刷屏。
    threshold_indices = []
    for i in range(len(amplitudes)):
        freq_hz = frequencies[i]
        amp_dbuv = amplitudes[i]
        fcc_limit, ce_limit = get_fcc_ce_limits(freq_hz, detector_type=detector_type)
        
        if amp_dbuv > fcc_limit or amp_dbuv > ce_limit:
            threshold_indices.append(i)
    threshold_peaks = collapse_contiguous_indices(threshold_indices, amplitudes)
    
    # 合并峰值并去重
    all_peaks = list(set(list(primary_peaks) + list(secondary_peaks) + threshold_peaks))
    
    # 如果没有检测到峰值，使用原始方法
    if len(all_peaks) == 0:
        all_peaks = find_peaks_manual(amplitudes, distance=peak_distance//2, prominence=min_prominence*0.3)
    
    # 计算每个峰值的重要性分数
    peak_scores = []
    for idx in all_peaks:
        if idx < 0 or idx >= len(amplitudes):
            continue
            
        amp_dbuv = amplitudes[idx]
        freq_hz = frequencies[idx]
        limit_info = get_emission_limit_info(freq_hz, detector_type=detector_type)
        fcc_limit = limit_info["fcc_limit"]
        ce_limit = limit_info["ce_limit"]
        
        # 计算裕量（相对于限值）
        fcc_margin = amp_dbuv - fcc_limit
        ce_margin = amp_dbuv - ce_limit
        
        # 峰值重要性评分（综合考虑幅度和裕量）
        amplitude_score = (amp_dbuv - mean_amp) / std_amp if std_amp > 0 else 0
        margin_score = max(fcc_margin, ce_margin, 0)  # 只考虑正裕量
        prominence_score = 0
        
        # 计算相对于相邻点的显著性
        if 1 <= idx < len(amplitudes) - 1:
            left_diff = amp_dbuv - amplitudes[idx-1]
            right_diff = amp_dbuv - amplitudes[idx+1]
            prominence_score = min(left_diff, right_diff)
        
        # 对于超过限值的点给予更高的评分
        exceed_bonus = 15 if (fcc_margin > 0 or ce_margin > 0) else 0  # 提高超限点的优先级
        
        # 综合评分
        total_score = amplitude_score * 0.2 + margin_score * 0.5 + prominence_score * 0.2 + exceed_bonus
        peak_scores.append((idx, total_score, amp_dbuv, fcc_margin, ce_margin, fcc_limit, ce_limit))
    
    # 按重要性排序
    peak_scores.sort(key=lambda x: x[1], reverse=True)
    
    # 根据频率范围调整返回的峰值数量
    freq_range_mhz = (frequencies[-1] - frequencies[0]) / 1e6
    if freq_range_mhz < 1:  # 低于1MHz范围
        max_peaks = 25  # 增加峰值数量
    elif freq_range_mhz < 100:  # 1MHz-100MHz范围
        max_peaks = 40  # 增加峰值数量
    else:  # 高于100MHz范围
        max_peaks = 50  # 增加峰值数量
    
    # 确保包含所有超标的峰值
    peak_results = []
    exceed_peaks = []
    normal_peaks = []
    
    for idx, score, amp_dbuv, fcc_margin, ce_margin, fcc_limit, ce_limit in peak_scores:
        freq_hz = frequencies[idx]
        peak_data = {
            'frequency_hz': freq_hz,
            'frequency_mhz': freq_hz / 1e6,
            'amplitude_dbuv': amp_dbuv,
            'fcc_limit': fcc_limit,
            'ce_limit': ce_limit,
            'fcc_margin': fcc_margin,
            'ce_margin': ce_margin,
            'exceed_fcc': fcc_margin > 0,
            'exceed_ce': ce_margin > 0,
            'importance_score': score,
            'limit_unit': limit_info["unit"],
            'limit_measurement_type': limit_info["measurement_type"],
            'fcc_detector': limit_info["fcc_detector"],
            'ce_detector': limit_info["ce_detector"],
            'limit_note': limit_info["note"],
        }
        
        # 确保所有超过限值的点都被包含
        if fcc_margin > 0 or ce_margin > 0:
            exceed_peaks.append(peak_data)
        else:
            normal_peaks.append(peak_data)
    
    # 首先添加所有超标的峰值
    peak_results.extend(exceed_peaks)
    
    # 然后添加重要的正常峰值，直到达到最大数量
    remaining_slots = max_peaks - len(peak_results)
    peak_results.extend(normal_peaks[:remaining_slots])
    
    # 按频率排序以便显示
    peak_results.sort(key=lambda x: x['frequency_hz'])
    
    return peak_results


def save_spectrum_data(frequencies, amplitudes, filename=None):
    """
    保存频谱数据
    """
    measurement_folder = 'measurement_data'
    if not os.path.exists(measurement_folder):
        os.makedirs(measurement_folder)
    
    if filename is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"spectrum_{timestamp}.csv"
    
    filepath = os.path.join(measurement_folder, filename)
    
    with open(filepath, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Frequency (Hz)', 'Amplitude (dBμV)'])
        for freq, amp in zip(frequencies, amplitudes):
            writer.writerow([freq, amp])
    
    return filepath

def save_peak_analysis(peak_results, filename=None):
    """
    保存峰值分析结果
    """
    if not peak_results:
        return None
    
    measurement_folder = 'measurement_data'
    if not os.path.exists(measurement_folder):
        os.makedirs(measurement_folder)
    
    if filename is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"peak_analysis_{timestamp}.csv"
    
    filepath = os.path.join(measurement_folder, filename)
    
    with open(filepath, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow([
            '频率(MHz)', '幅度(dBμV)', 
            'FCC限值(dBμV)', 'CE限值(dBμV)',
            'FCC裕量(dB)', 'CE裕量(dB)',
            'FCC超标', 'CE超标'
        ])
        
        for peak in peak_results:
            writer.writerow([
                f"{peak['frequency_mhz']:.3f}",
                f"{peak['amplitude_dbuv']:.2f}",
                f"{peak['fcc_limit']:.1f}",
                f"{peak['ce_limit']:.1f}",
                f"{peak['fcc_margin']:.2f}",
                f"{peak['ce_margin']:.2f}",
                '是' if peak['exceed_fcc'] else '否',
                '是' if peak['exceed_ce'] else '否'
            ])
    
    return filepath

