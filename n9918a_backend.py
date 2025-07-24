# n9918a_backend.py
import pyvisa
import matplotlib.pyplot as plt
import numpy as np
import time
import csv
import os
from datetime import datetime
from scipy import signal

# 在文件开头添加平台检测
import platform

# 设置matplotlib字体 - Mac兼容版本
if platform.system() == "Darwin":  # macOS
    plt.rcParams['font.sans-serif'] = ['PingFang SC', 'Heiti SC', 'Arial Unicode MS', 'DejaVu Sans']
else:  # Windows/Linux
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']

plt.rcParams['axes.unicode_minus'] = False

class N9918AController:
    """
    N9918A FieldFox Network Analyzer Controller for EMC Testing
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
        
    def connect(self):
        try:
            self.rm = pyvisa.ResourceManager()
            self.device = self.rm.open_resource(f"TCPIP0::{self.ip_address}::inst0::INSTR")
            self.device.timeout = self.timeout
            
            self.device.write("*CLS")
            device_id = self.device.query("*IDN?")
            print(f"Connected to: {device_id}")
            
            self.device.write("INST:SEL 'SA'")
            time.sleep(1)
            
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
    
    def configure_custom_settings(self, start_freq, stop_freq, n_points, rbw, vbw):
        """
        配置自定义参数
        """
        return self._configure_device(start_freq, stop_freq, n_points, rbw, vbw, "Custom")
    
    def _configure_device(self, start_freq, stop_freq, n_points, rbw, vbw, config_name):
        """
        内部配置设备方法
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return False
            
        try:
            print(f"🔧 配置设备参数: {config_name}")
            
            # 关闭连续扫描
            self.device.write("INIT:CONT OFF")
            time.sleep(0.5)
            
            # Set frequency range
            self.device.write(f":SENS:FREQ:STAR {start_freq}")
            time.sleep(0.2)
            self.device.write(f":SENS:FREQ:STOP {stop_freq}")
            time.sleep(0.2)
            print(f"📡 频率范围: {start_freq/1e6:.3f}MHz - {stop_freq/1e9:.3f}GHz")
            
            # Set number of points
            self.device.write(f":SENS:SWE:POIN {n_points}")
            time.sleep(0.2)
            print(f"📈 采样点数: {n_points}")
            
            # Set RBW and VBW
            self.device.write(f":SENS:BAND:RES {rbw}")
            time.sleep(0.5)
            self.device.write(f":SENS:BAND:VID {vbw}")
            time.sleep(0.5)
            print(f"⚙️  RBW: {rbw}Hz, VBW: {vbw}Hz")
            
            # Set Detector to Sample
            self.device.write(":SENS:DET SAMPLE")
            time.sleep(0.2)
            print("🎯 Detector: Sample")
            
            # Set Internal Amplifier ON
            self.device.write(":SENS:POW:GAIN:STAT ON")
            time.sleep(0.2)
            print("🔊 内部放大器: ON")
            
            # Set Internal Attenuator to 0dB
            self.device.write(":SENS:POW:ATT 0")
            time.sleep(0.2)
            print("🔇 内部衰减器: 0dB")
            
            # Store parameters
            self.start_freq = start_freq
            self.stop_freq = stop_freq
            self.n_points = n_points
            self.rbw = rbw
            self.vbw = vbw
            self.current_config = config_name
            
            print("✅ 参数配置完成! (连续扫描已暂停)")
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
            # 触发单次扫描
            self.device.write(":INIT:IMM")
            
            # 获取扫描时间
            try:
                sweep_time = float(self.device.query(":SENS:SWE:TIME?"))
                wait_time = max(sweep_time * 1.2, 1.0)  # 等待1.2倍扫描时间或至少1秒
            except:
                # 如果无法获取扫描时间，使用估算值
                wait_time = max(2.0, (self.stop_freq - self.start_freq) / 1e9 * 3)
            
            print(f"⏳ 等待扫描完成 ({wait_time:.1f}秒)...")
            time.sleep(wait_time)
            
            # Read trace data
            self.device.write(":TRAC:DATA?")
            trace_data = self.device.read()
            amplitudes_dBuv = [float(x) for x in trace_data.split(",")]
            
            # Calculate frequency array
            freq_step = (self.stop_freq - self.start_freq) / (self.n_points - 1)
            frequencies = [self.start_freq + i * freq_step for i in range(self.n_points)]
            
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
            "vbw": self.vbw
        }

# 修正后的EMC标准限值函数
def get_fcc_ce_limits(freq_hz):
    """
    获取FCC和CE标准限值 (单位: dBuV)
    """
    freq_mhz = freq_hz / 1e6
    
    # FCC Part 15 Class B 准峰值限值
    if 0.009 <= freq_mhz < 0.050:      # 9kHz-50kHz
        fcc_limit = 34  # 例如值，实际需要查表
    elif 0.050 <= freq_mhz < 0.150:    # 50kHz-150kHz
        fcc_limit = 40
    elif 0.150 <= freq_mhz < 0.500:    # 150kHz-500kHz
        fcc_limit = 40
    elif 0.500 <= freq_mhz < 1.705:    # 500kHz-1.705MHz
        fcc_limit = 40
    elif 1.705 <= freq_mhz < 30:       # 1.705MHz-30MHz
        fcc_limit = 40
    elif 30 <= freq_mhz < 88:          # 30MHz-88MHz
        fcc_limit = 40
    elif 88 <= freq_mhz < 216:         # 88MHz-216MHz
        fcc_limit = 40
    elif 216 <= freq_mhz < 960:        # 216MHz-960MHz
        fcc_limit = 46
    elif 960 <= freq_mhz <= 10000:     # 960MHz-10GHz
        fcc_limit = 40
    else:
        fcc_limit = 120  # 超出范围设为高值
    
    # EN 55032 Class B 限值 (更准确的分段)
    if 0.009 <= freq_mhz < 0.050:      # 9kHz-50kHz
        ce_limit = 34
    elif 0.050 <= freq_mhz < 0.150:    # 50kHz-150kHz
        ce_limit = 40
    elif 0.150 <= freq_mhz < 0.500:    # 150kHz-500kHz
        ce_limit = 40
    elif 0.500 <= freq_mhz < 1.705:    # 500kHz-1.705MHz
        ce_limit = 40
    elif 1.705 <= freq_mhz < 30:       # 1.705MHz-30MHz
        ce_limit = 40
    elif 30 <= freq_mhz < 230:         # 30MHz-230MHz
        ce_limit = 40
    elif 230 <= freq_mhz < 1000:       # 230MHz-1GHz
        ce_limit = 47
    elif 1000 <= freq_mhz <= 10000:    # 1GHz-10GHz
        ce_limit = 40
    else:
        ce_limit = 120  # 超出范围设为高值
    
    return fcc_limit, ce_limit

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

def post_process_peak_search(frequencies, amplitudes, peak_distance=50, min_prominence=3):
    """
    后处理峰值搜索
    """
    peak_indices, properties = signal.find_peaks(
        amplitudes, 
        distance=peak_distance,
        prominence=min_prominence,
        height=np.mean(amplitudes) + min_prominence
    )
    
    if len(peak_indices) == 0:
        peak_indices = find_peaks_manual(amplitudes, distance=peak_distance, prominence=min_prominence)
    
    peak_indices = peak_indices[:10] if len(peak_indices) > 10 else peak_indices
    
    peak_results = []
    for idx in peak_indices:
        freq_hz = frequencies[idx]
        amp_dbuv = amplitudes[idx]
        fcc_limit, ce_limit = get_fcc_ce_limits(freq_hz)
        
        fcc_margin = amp_dbuv - fcc_limit
        ce_margin = amp_dbuv - ce_limit
        
        peak_results.append({
            'frequency_hz': freq_hz,
            'frequency_mhz': freq_hz / 1e6,
            'amplitude_dbuv': amp_dbuv,
            'fcc_limit': fcc_limit,
            'ce_limit': ce_limit,
            'fcc_margin': fcc_margin,
            'ce_margin': ce_margin,
            'exceed_fcc': fcc_margin > 0,
            'exceed_ce': ce_margin > 0
        })
    
    return peak_results

def plot_emc_spectrum(frequencies, amplitudes, peak_results=None, show_limits=True):
    """
    绘制EMC频谱图 - 自适应窗口大小版本
    """
    # 创建图形，使用相对大小
    fig, ax = plt.subplots(figsize=(12, 8))
    
    # 设置中文字体
    plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans', 'Arial Unicode MS']
    plt.rcParams['axes.unicode_minus'] = False
    
    freq_mhz = [f / 1e6 for f in frequencies]
    
    # 绘制测量数据
    ax.semilogx(freq_mhz, amplitudes, 'b-', linewidth=1, label='测量频谱', alpha=0.8)
    
    # 绘制FCC和CE限值
    if show_limits and frequencies:
        fcc_limits = []
        ce_limits = []
        for freq in frequencies:
            fcc_limit, ce_limit = get_fcc_ce_limits(freq)
            fcc_limits.append(fcc_limit)
            ce_limits.append(ce_limit)
        
        ax.semilogx(freq_mhz, fcc_limits, 'r--', linewidth=1.5, label='FCC Class B', alpha=0.7)
        ax.semilogx(freq_mhz, ce_limits, 'g--', linewidth=1.5, label='CE Class B', alpha=0.7)
    
    # 标记峰值
    if peak_results:
        for peak in peak_results:
            freq_mhz_peak = peak['frequency_mhz']
            amp_dbuv = peak['amplitude_dbuv']
            ax.plot(freq_mhz_peak, amp_dbuv, 'ro', markersize=6, 
                   markeredgecolor='black', markeredgewidth=0.5)
            
            exceed_fcc = peak['exceed_fcc']
            exceed_ce = peak['exceed_ce']
            color = 'red' if exceed_fcc or exceed_ce else 'black'
            
            # 简化的标注，避免重叠
            ax.annotate(f'{freq_mhz_peak:.1f}MHz', 
                       xy=(freq_mhz_peak, amp_dbuv), 
                       xytext=(0, 15), textcoords='offset points',
                       fontsize=7, color=color,
                       ha='center', va='bottom',
                       bbox=dict(boxstyle='round,pad=0.2', facecolor='white', alpha=0.8, edgecolor='gray'))
    
    # 设置标签和标题
    ax.set_xlabel('频率 (MHz)', fontsize=10)
    ax.set_ylabel('幅度 (dBμV)', fontsize=10)
    ax.set_title('EMC频谱分析', fontsize=12, pad=20)
    
    # 网格和图例
    ax.grid(True, which="both", alpha=0.3, linestyle='-', linewidth=0.5)
    ax.legend(loc='upper right', fontsize=9)
    
    # 设置坐标轴范围
    if frequencies:
        ax.set_xlim([min(freq_mhz), max(freq_mhz)])
    
    if amplitudes:
        y_min = min(min(amplitudes), 20) - 10
        y_max = max(max(amplitudes), 80) + 10
        ax.set_ylim([y_min, y_max])
    
    # 优化布局
    plt.tight_layout()
    return fig
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

def print_peak_summary(peak_results):
    """
    打印峰值分析摘要
    """
    if not peak_results:
        print("未检测到峰值")
        return
    
    print("\n📊 峰值分析结果:")
    print("=" * 90)
    print(f"{'频率(MHz)':<12} {'幅度(dBμV)':<12} {'FCC限值':<10} {'CE限值':<10} {'FCC裕量':<10} {'CE裕量':<10} {'状态':<15}")
    print("-" * 90)
    
    for peak in peak_results:
        status = []
        if peak['exceed_fcc']:
            status.append("FCC超标")
        if peak['exceed_ce']:
            status.append("CE超标")
        if not status:
            status = ["合规"]
        
        print(f"{peak['frequency_mhz']:<12.3f} "
              f"{peak['amplitude_dbuv']:<12.2f} "
              f"{peak['fcc_limit']:<10.1f} "
              f"{peak['ce_limit']:<10.1f} "
              f"{peak['fcc_margin']:<10.2f} "
              f"{peak['ce_margin']:<10.2f} "
              f"{', '.join(status):<15}")