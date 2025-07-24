import pyvisa
import matplotlib.pyplot as plt
import numpy as np
import time
import csv
import os
from datetime import datetime
from scipy import signal

# 设置matplotlib字体
plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

class N9918AController:
    """
    N9918A FieldFox Network Analyzer Controller for EMC Testing
    """
    
    def __init__(self, ip_address='192.168.0.124', timeout=10000):
        self.ip_address = ip_address
        self.timeout = timeout
        self.rm = None
        self.device = None
        self.connected = False
        self.start_freq = None
        self.stop_freq = None
        self.n_points = None
        
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
    
    def configure_emc_settings(self, start_freq, stop_freq, n_points=1001):
        """
        Configure EMC test settings with 100kHz RBW/VBW
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return False
            
        try:
            print("🔧 配置EMC测试参数...")
            
            # 关闭连续扫描
            self.device.write("INIT:CONT OFF")
            time.sleep(0.5)
            
            # Set frequency range
            self.device.write(f":SENS:FREQ:STAR {start_freq}")
            time.sleep(0.2)
            self.device.write(f":SENS:FREQ:STOP {stop_freq}")
            time.sleep(0.2)
            print(f"📡 频率范围: {start_freq/1e6:.0f}MHz - {stop_freq/1e9:.1f}GHz")
            
            # Set number of points
            self.device.write(f":SENS:SWE:POIN {n_points}")
            time.sleep(0.2)
            print(f"📈 采样点数: {n_points}")
            
            # Set RBW and VBW to 100kHz (standard EMC value)
            self.device.write(":SENS:BAND:RES 100e3")  # 100kHz RBW
            time.sleep(0.5)
            self.device.write(":SENS:BAND:VID 100e3")  # 100kHz VBW
            time.sleep(0.5)
            print("⚙️  RBW/VBW: 100kHz")
            
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
            
            print("✅ EMC参数配置完成! (连续扫描已暂停)")
            return True
            
        except Exception as e:
            print(f"ERROR: Failed to configure measurement - {e}")
            return False
    
    def read_trace_data(self):
        """
        Read trace data from the device (直接读取dBμV数据)
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return None, None
            
        try:
            # 触发单次扫描
            self.device.write(":INIT:IMM")
            
            # 等待扫描完成
            sweep_time = 2.0  # 固定等待时间
            print(f"⏳ 等待扫描完成 ({sweep_time:.1f}秒)...")
            time.sleep(sweep_time)
            
            # Read trace data (设备直接输出的就是dBμV)
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

def get_fcc_ce_limits(freq_hz):
    """
    获取FCC和CE标准限值 (单位: dBuV)
    """
    freq_mhz = freq_hz / 1e6
    
    # FCC Part 15 Class B 准峰值限值 (简化版)
    if 30 <= freq_mhz <= 88:
        fcc_limit = 40  # 30-88 MHz
    elif 88 <= freq_mhz <= 216:
        fcc_limit = 40  # 88-216 MHz
    elif 216 <= freq_mhz <= 960:
        fcc_limit = 46  # 216-960 MHz
    elif 960 <= freq_mhz <= 10000:  # 10GHz
        fcc_limit = 40  # 960MHz以上
    else:
        fcc_limit = 120  # 超出范围，设为高值
    
    # EN 55032 Class B 限值 (简化版)
    if 30 <= freq_mhz <= 230:
        ce_limit = 40   # 30-230 MHz
    elif 230 <= freq_mhz <= 1000:
        ce_limit = 47   # 230MHz-1GHz
    elif 1000 <= freq_mhz <= 10000:  # 10GHz
        ce_limit = 40   # 1GHz以上
    else:
        ce_limit = 120  # 超出范围，设为高值
    
    return fcc_limit, ce_limit

def find_peaks_manual(data, distance=5, prominence=3):
    """
    手动实现峰值检测
    """
    peaks = []
    n = len(data)
    
    for i in range(1, n-1):
        # 检查是否为局部最大值
        is_peak = True
        # 检查左侧
        for j in range(max(0, i-distance), i):
            if data[j] >= data[i]:
                is_peak = False
                break
        if not is_peak:
            continue
        # 检查右侧
        for j in range(i+1, min(n, i+distance+1)):
            if data[j] >= data[i]:
                is_peak = False
                break
        
        if is_peak and data[i] > np.mean(data) + prominence:
            peaks.append(i)
    
    # 按幅度排序
    peaks.sort(key=lambda x: data[x], reverse=True)
    return peaks

def post_process_peak_search(frequencies, amplitudes, peak_distance=50, min_prominence=3):
    """
    后处理峰值搜索
    """
    # 使用scipy的峰值检测
    peak_indices, properties = signal.find_peaks(
        amplitudes, 
        distance=peak_distance,
        prominence=min_prominence,
        height=np.mean(amplitudes) + min_prominence
    )
    
    # 如果scipy方法失败，使用手动方法
    if len(peak_indices) == 0:
        peak_indices = find_peaks_manual(amplitudes, distance=peak_distance, prominence=min_prominence)
    
    # 获取前10个最高峰值
    peak_indices = peak_indices[:10] if len(peak_indices) > 10 else peak_indices
    
    # 计算每个峰值与标准限值的关系
    peak_results = []
    for idx in peak_indices:
        freq_hz = frequencies[idx]
        amp_dbuv = amplitudes[idx]
        fcc_limit, ce_limit = get_fcc_ce_limits(freq_hz)
        
        # 计算超出限值的dB数
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

def plot_emc_spectrum(frequencies, amplitudes, peak_results=None):
    """
    绘制EMC频谱图，包含FCC和CE限值
    """
    fig, ax = plt.subplots(figsize=(15, 8))
    
    # 转换频率为MHz
    freq_mhz = [f / 1e6 for f in frequencies]
    
    # 绘制测量数据 (对数频率轴)
    ax.semilogx(freq_mhz, amplitudes, 'b-', linewidth=1, label='Measured Spectrum')
    
    # 绘制FCC和CE限值
    fcc_limits = []
    ce_limits = []
    for freq in frequencies:
        fcc_limit, ce_limit = get_fcc_ce_limits(freq)
        fcc_limits.append(fcc_limit)
        ce_limits.append(ce_limit)
    
    ax.semilogx(freq_mhz, fcc_limits, 'r--', linewidth=2, label='FCC Part 15 Class B')
    ax.semilogx(freq_mhz, ce_limits, 'g--', linewidth=2, label='EN 55032 Class B')
    
    # 标记峰值
    if peak_results:
        for peak in peak_results:
            freq_mhz_peak = peak['frequency_mhz']
            amp_dbuv = peak['amplitude_dbuv']
            ax.plot(freq_mhz_peak, amp_dbuv, 'ro', markersize=6)
            # 添加标签
            exceed_fcc = peak['exceed_fcc']
            exceed_ce = peak['exceed_ce']
            color = 'red' if exceed_fcc or exceed_ce else 'black'
            ax.annotate(f"{freq_mhz_peak:.1f}MHz\n{amp_dbuv:.1f}dBμV", 
                       xy=(freq_mhz_peak, amp_dbuv), 
                       xytext=(5, 5), textcoords='offset points',
                       fontsize=8, color=color)
    
    ax.set_xlabel('Frequency (MHz)')
    ax.set_ylabel('Amplitude (dBμV)')
    ax.set_title('EMC Spectrum Analysis with FCC/CE Limits')
    ax.grid(True, which="both", alpha=0.3)
    ax.legend()
    
    # 设置频率轴范围
    ax.set_xlim([min(freq_mhz), max(freq_mhz)])
    
    # 自动调整纵轴范围，使其与设备显示一致
    y_min = min(min(amplitudes), min(fcc_limits), min(ce_limits)) - 10
    y_max = max(max(amplitudes), max(fcc_limits), max(ce_limits)) + 10
    ax.set_ylim([y_min, y_max])
    
    plt.tight_layout()
    plt.show()

def save_peak_analysis(peak_results, filename=None):
    """
    保存峰值分析结果到CSV文件
    """
    if not peak_results:
        print("No peak results to save")
        return
    
    # Create measurement folder if it doesn't exist
    measurement_folder = 'measurement_data'
    if not os.path.exists(measurement_folder):
        os.makedirs(measurement_folder)
    
    # Generate filename if not provided
    if filename is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"peak_analysis_{timestamp}.csv"
    
    filepath = os.path.join(measurement_folder, filename)
    
    with open(filepath, 'w', newline='') as f:
        writer = csv.writer(f)
        
        # Write header
        writer.writerow([
            'Frequency (MHz)', 'Amplitude (dBμV)', 
            'FCC Limit (dBμV)', 'CE Limit (dBμV)',
            'FCC Margin (dB)', 'CE Margin (dB)',
            'Exceed FCC', 'Exceed CE'
        ])
        
        # Write peak data
        for peak in peak_results:
            writer.writerow([
                f"{peak['frequency_mhz']:.3f}",
                f"{peak['amplitude_dbuv']:.2f}",
                f"{peak['fcc_limit']:.1f}",
                f"{peak['ce_limit']:.1f}",
                f"{peak['fcc_margin']:.2f}",
                f"{peak['ce_margin']:.2f}",
                'Yes' if peak['exceed_fcc'] else 'No',
                'Yes' if peak['exceed_ce'] else 'No'
            ])
    
    print(f"💾 峰值分析结果已保存到: {filepath}")
    return filepath

def print_peak_summary(peak_results):
    """
    打印峰值分析摘要
    """
    if not peak_results:
        print("No peaks found")
        return
    
    print("\n📊 峰值分析结果:")
    print("=" * 80)
    print(f"{'频率(MHz)':<12} {'幅度(dBμV)':<12} {'FCC限值':<10} {'CE限值':<10} {'FCC裕量':<10} {'CE裕量':<10} {'状态':<10}")
    print("-" * 80)
    
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
              f"{', '.join(status):<10}")

def main():
    # 创建控制器实例
    controller = N9918AController(ip_address='192.168.20.175', timeout=20000)
    
    print("🔧 开始连接N9918A设备 (EMC测试模式)...")
    print("IP地址: 192.168.20.39")
    print("测试范围: 30MHz ~ 1GHz")
    print("RBW/VBW: 100kHz")
    print("-" * 50)
    
    if not controller.connect():
        print("❌ 连接失败")
        return
    
    try:
        # 配置EMC测试参数
        print("\n⚙️  配置EMC测试参数...")
        success = controller.configure_emc_settings(
            start_freq=30e6,    # 30MHz
            stop_freq=1e9,      # 1GHz
            n_points=2001       # 更高分辨率
        )
        
        if not success:
            print("❌ 参数配置失败")
            controller.disconnect()
            return
        
        # 等待用户确认
        print("\n⚠️  参数已设置完成，设备处于暂停扫描状态")
        print("    请检查设备屏幕上的所有参数是否正确")
        input("    确认无误后，按回车键开始EMC测量... ")
        
        print("\n🔍 开始EMC测量...")
        
        # 读取数据
        frequencies, amplitudes = controller.read_trace_data()
        
        if frequencies is None or amplitudes is None:
            print("❌ 读取数据失败")
            controller.disconnect()
            return
        
        print(f"✅ 成功读取 {len(frequencies)} 个数据点")
        
        # 后处理 - 峰值搜索
        print("🔍 进行峰值分析...")
        peak_results = post_process_peak_search(
            frequencies, amplitudes, 
            peak_distance=50,      # 峰值间最小距离
            min_prominence=3       # 最小突出度
        )
        
        # 显示基本统计信息
        max_amp = max(amplitudes)
        min_amp = min(amplitudes)
        avg_amp = sum(amplitudes) / len(amplitudes)
        max_freq_idx = amplitudes.index(max_amp)
        max_freq = frequencies[max_freq_idx] / 1e6
        
        print(f"\n📈 测量结果统计:")
        print(f"   最大值: {max_amp:.2f} dBμV (在 {max_freq:.2f} MHz)")
        print(f"   最小值: {min_amp:.2f} dBμV")
        print(f"   平均值: {avg_amp:.2f} dBμV")
        print(f"   检测到峰值数: {len(peak_results)}")
        
        # 打印峰值摘要
        print_peak_summary(peak_results)
        
        # 保存完整数据
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 保存频谱数据
        spectrum_filename = f"emc_spectrum_{timestamp}.csv"
        measurement_folder = 'measurement_data'
        if not os.path.exists(measurement_folder):
            os.makedirs(measurement_folder)
        
        spectrum_filepath = os.path.join(measurement_folder, spectrum_filename)
        with open(spectrum_filepath, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Frequency (Hz)', 'Amplitude (dBμV)'])
            for freq, amp in zip(frequencies, amplitudes):
                writer.writerow([freq, amp])
        print(f"💾 完整频谱数据已保存到: {spectrum_filepath}")
        
        # 保存峰值分析结果
        peak_filename = f"peak_analysis_{timestamp}.csv"
        save_peak_analysis(peak_results, peak_filename)
        
        # 显示图形
        print("📊 生成EMC分析图表...")
        plot_emc_spectrum(frequencies, amplitudes, peak_results)
        
        print("\n✅ EMC测试完成!")
        
    except Exception as e:
        print(f"❌ 执行过程中出现错误: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        controller.disconnect()
        print("\n🔌 设备连接已关闭")

if __name__ == "__main__":
    main()