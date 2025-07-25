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


    
    def get_emc_measurement_fast(self, duration_seconds=15):
        """
        快速EMC测量（采集时间序列数据，PC端计算多种模式）
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return {}
        
        total_start_time = time.time()
        print(f"🚀 开始快速EMC测量 ({duration_seconds} 秒)")
        print("=" * 50)
        
        try:
            # 1. 收集时间序列数据
            time_series_data = self.collect_emc_time_series(duration_seconds)
            
            if not time_series_data:
                print("❌ 未能收集到时间序列数据")
                return {}
            
            collection_time = time.time() - total_start_time
            print(f"   ⏱️  数据采集用时: {collection_time:.1f} 秒")
            print(f"   📊  实际采样次数: {len(time_series_data)}")
            
            # 2. PC端计算多种检测器模式
            print(f"\n🔬 PC端计算EMC检测器模式...")
            calculation_start_time = time.time()
            
            results = {}
            detector_modes = ["PEAK", "QUASI_PEAK", "AVERAGE"]
            
            for mode in detector_modes:
                frequencies, amplitudes = calculate_emc_detector_modes(time_series_data, mode)
                if frequencies is not None and amplitudes is not None:
                    results[mode] = (frequencies, amplitudes)
                    print(f"   ✅ {mode} 模式计算完成")
                    if amplitudes:
                        max_val = max(amplitudes)
                        min_val = min(amplitudes)
                        avg_val = sum(amplitudes) / len(amplitudes)
                        print(f"       最大值: {max_val:.2f} dBμV, 最小值: {min_val:.2f} dBμV, 平均值: {avg_val:.2f} dBμV")
                else:
                    print(f"   ❌ {mode} 模式计算失败")
            
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
                "end_time": time_series_data[-1]['timestamp'] if time_series_data else 0
            }
            
            # 添加测量摘要
            results["measurement_summary"] = {
                "total_duration": total_time,
                "actual_measurement_time": duration_seconds,
                "data_points": len(time_series_data[0]['amplitudes']) if time_series_data else 0,
                "total_samples": len(time_series_data),
                "modes_computed": detector_modes,
                "measurement_time": time.strftime("%Y-%m-%d %H:%M:%S")
            }
            
            print(f"   ⏱️  计算用时: {calculation_time:.1f} 秒")
            print(f"✅ 所有处理完成! 总用时: {total_time:.1f} 秒")
            
            return results
            
        except Exception as e:
            print(f"ERROR: 快速EMC测量失败 - {e}")
            import traceback
            traceback.print_exc()
            return {}

    def collect_emc_time_series(self, duration_seconds=15):
        """
        稳定版时间序列数据采集 - 支持长时间采样
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return []
        
        print(f"🔄 开始时间序列数据采集 ({duration_seconds} 秒)")
        
        try:
            # 设置为SAMPLE模式
            self.device.write(":SENS:DET SAMP")
            time.sleep(0.2)
            
            # 设置更长的超时时间用于长时间采样
            original_timeout = self.device.timeout
            self.device.timeout = 30000  # 30秒超时
            
            # 开启连续扫描
            self.device.write("INIT:CONT ON")
            time.sleep(0.5)
            
            # 动态调整采样间隔 - 根据测量时长和RBW
            if duration_seconds <= 30:
                sample_interval = 0.3  # 短时间采样用0.3秒，更密集
            elif duration_seconds <= 120:
                sample_interval = 0.8  # 中等时间用0.8秒
            else:
                sample_interval = 1.5  # 长时间采样用1.5秒

            # 根据RBW调整 - RBW越小，需要更长的稳定时间
            if hasattr(self, 'rbw') and self.rbw:
                if self.rbw <= 1000:  # 1kHz以下
                    sample_interval *= 1.5
                elif self.rbw <= 10000:  # 10kHz以下
                    sample_interval *= 1.2

            print(f"   ⏱️  采样间隔: {sample_interval}s (优化后), 目标采样次数: {int(duration_seconds / sample_interval)}")
            
            time_series_data = []
            start_time = time.time()
            next_sample_time = start_time + sample_interval
            sample_count = 0
            max_samples = int(duration_seconds / sample_interval)
            
            print(f"   ⏱️  采样间隔: {sample_interval}s, 目标采样次数: {max_samples}")
            
            # 用于检测卡死的变量
            last_successful_time = start_time
            consecutive_failures = 0
            max_consecutive_failures = 3
            
            while time.time() - start_time < duration_seconds and sample_count < max_samples:
                current_time = time.time()
                
                if current_time >= next_sample_time:
                    try:
                        # 每10次采样后清理一次通信缓冲区
                        if sample_count % 10 == 0 and sample_count > 0:
                            print(f"   🔧 清理通信缓冲区 (采样 #{sample_count})")
                            self.device.write("*CLS")  # 清除状态
                            time.sleep(0.1)
                        
                        # 每20次采样后重新启动连续扫描
                        if sample_count % 20 == 0 and sample_count > 0:
                            print(f"   🔄 重新启动连续扫描 (采样 #{sample_count})")
                            self.device.write("INIT:CONT OFF")
                            time.sleep(0.2)
                            self.device.write("INIT:CONT ON")
                            time.sleep(0.5)
                        
                        # 设置较短的临时超时用于单次读取
                        self.device.timeout = 10000  # 10秒
                        
                        # 读取当前trace数据
                        sample_start_time = time.time()
                        self.device.write(":TRACE:DATA?")
                        trace_data = self.device.read()
                        
                        # 检查读取是否超时
                        read_duration = time.time() - sample_start_time
                        if read_duration > 8:  # 如果读取超过8秒，认为可能有问题
                            print(f"   ⚠️  读取耗时异常: {read_duration:.2f}s")
                        
                        amplitudes = [float(x) for x in trace_data.split(",")]
                        
                        # 验证数据完整性
                        if len(amplitudes) != self.n_points:
                            print(f"   ⚠️  数据点数不匹配: 期望{self.n_points}, 实际{len(amplitudes)}")
                            consecutive_failures += 1
                            if consecutive_failures >= max_consecutive_failures:
                                print(f"   ❌ 连续失败{consecutive_failures}次，停止采样")
                                break
                            continue
                        
                        # 第一次采样时计算频率数组
                        if not time_series_data:
                            freq_step = (self.stop_freq - self.start_freq) / (self.n_points - 1)
                            frequencies = [self.start_freq + i * freq_step for i in range(self.n_points)]
                        else:
                            frequencies = time_series_data[0]['frequencies']
                        
                        # 记录采样
                        sample_record = {
                            'timestamp': current_time - start_time,
                            'frequencies': frequencies,
                            'amplitudes': amplitudes
                        }
                        time_series_data.append(sample_record)
                        sample_count += 1
                        consecutive_failures = 0  # 重置失败计数
                        last_successful_time = current_time
                        
                        # 显示进度
                        progress = (sample_count / max_samples) * 100
                        if sample_count % 5 == 0 or sample_count <= 10:
                            elapsed = current_time - start_time
                            remaining = duration_seconds - elapsed
                            print(f"   📊 采样 #{sample_count}/{max_samples} ({progress:.1f}%) "
                                f"已用时: {elapsed:.1f}s, 剩余: {remaining:.1f}s")
                        
                        # 更新下次采样时间
                        next_sample_time = current_time + sample_interval
                        
                        # 恢复原始超时设置
                        self.device.timeout = original_timeout
                        
                    except pyvisa.errors.VisaIOError as e:
                        consecutive_failures += 1
                        print(f"   ⚠️  VISA通信错误 (第{consecutive_failures}次): {e}")
                        
                        if consecutive_failures >= max_consecutive_failures:
                            print(f"   ❌ 连续通信失败{consecutive_failures}次，尝试重新连接...")
                            # 尝试重新初始化连接
                            try:
                                self.device.write("INIT:CONT OFF")
                                time.sleep(1)
                                self.device.write("*CLS")
                                time.sleep(0.5)
                                self.device.write("INIT:CONT ON")
                                time.sleep(0.5)
                                consecutive_failures = 0
                                print(f"   ✅ 重新连接成功")
                            except:
                                print(f"   ❌ 重新连接失败，停止采样")
                                break
                        
                        # 等待一段时间后重试
                        time.sleep(1)
                        
                    except Exception as e:
                        consecutive_failures += 1
                        print(f"   ⚠️  采样失败 (第{consecutive_failures}次): {e}")
                        
                        if consecutive_failures >= max_consecutive_failures:
                            print(f"   ❌ 连续失败{consecutive_failures}次，停止采样")
                            break
                        
                        time.sleep(0.5)
                
                # 检查是否长时间无响应
                if current_time - last_successful_time > 30:  # 30秒无成功采样
                    print(f"   ❌ 设备长时间无响应，停止采样")
                    break
                
                # 短暂等待
                time.sleep(0.05)
            
            # 停止连续扫描
            try:
                self.device.write("INIT:CONT OFF")
                time.sleep(0.2)
                self.device.timeout = original_timeout  # 恢复原始超时
            except:
                print("   ⚠️  停止扫描时出现异常")
            
            print(f"✅ 时间序列采集完成! 总采样: {len(time_series_data)} 次")
            
            if time_series_data:
                actual_duration = time_series_data[-1]['timestamp']
                print(f"   📊 实际采样时长: {actual_duration:.1f}s")
                print(f"   📊 平均采样间隔: {actual_duration/len(time_series_data):.2f}s")
            
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
    
    print(f"   🎯 计算 {detector_type} 模式...")
    
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
            # 传递频率信息给准峰值计算
            current_freq = frequencies[freq_idx]
            detector_value = calculate_quasi_peak_value(times, values, current_freq)
        elif detector_type == "AVERAGE":
            detector_value = sum(values) / len(values)
        else:
            detector_value = values[-1] if values else 0
        
        result_amplitudes.append(detector_value)
    
    return frequencies, result_amplitudes


def calculate_quasi_peak_value(times, values, frequency_hz=None):
    """
    改进的准峰值计算 - 根据频率选择正确的时间常数
    """
    import numpy as np
    
    if len(times) <= 1:
        return max(0, values[0]) if values else 0
    
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
    
    # 数据预处理和排序
    time_value_pairs = [(float(t), max(0, float(v))) for t, v in zip(times, values)]
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
    
    # 最终约束：准峰值应该在合理范围内
    qp_value = max(qp_value, avg_value * 0.8)  # 至少是平均值的80%
    qp_value = min(qp_value, max_value)        # 不超过峰值
    
    return max(0, qp_value)

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
            print(f"💾 {mode} 最终数据已保存: {csv_filepath}")
    
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
            print(f"💾 所有采样详细数据已保存: {detailed_filepath}")
            
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
            print(f"💾 采样统计信息已保存: {stats_filepath}")
    
    # 保存测量摘要
    if "measurement_summary" in frequencies_dict:
        summary_filename = f"{filename_prefix}_summary.json"
        summary_filepath = os.path.join(measurement_folder, summary_filename)
        
        with open(summary_filepath, 'w') as f:
            json.dump(frequencies_dict["measurement_summary"], f, indent=2)
        
        saved_files.append(summary_filepath)
        print(f"💾 测量摘要已保存: {summary_filepath}")
    
    # 保存采样信息
    if "sampling_info" in frequencies_dict:
        info_filename = f"{filename_prefix}_sampling_info.json"
        info_filepath = os.path.join(measurement_folder, info_filename)
        
        with open(info_filename, 'w') as f:
            json.dump(frequencies_dict["sampling_info"], f, indent=2)
        
        saved_files.append(info_filepath)
        print(f"💾 采样信息已保存: {info_filepath}")
    
    return saved_files

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