import pyvisa
import matplotlib.pyplot as plt
import csv
from datetime import datetime

def read_current_spectrum():
    try:
        # 连接到设备
        rm = pyvisa.ResourceManager('@py')
        device = rm.open_resource('TCPIP::192.168.20.175::INSTR')
        device.timeout = 10000
        
        print("连接成功!")
        
        # 查询设备信息
        device_id = device.query("*IDN?")
        print(f"设备信息: {device_id}")
        
        # 确保在频谱分析仪模式
        device.write("INST:SEL 'SA'")
        
        # 获取当前的频率设置
        start_freq = float(device.query("SENS:FREQ:STAR?"))
        stop_freq = float(device.query("SENS:FREQ:STOP?"))
        n_points = int(device.query("SENS:SWE:POIN?"))
        
        print(f"频率范围: {start_freq/1e9:.3f} GHz - {stop_freq/1e9:.3f} GHz")
        print(f"采样点数: {n_points}")
        
        # 读取当前显示的频谱数据
        device.write("TRACE:DATA?")
        trace_data = device.read()
        amplitudes = [float(x) for x in trace_data.split(",")]
        
        print(f"读取到 {len(amplitudes)} 个数据点")
        
        # 生成频率数组
        freq_step = (stop_freq - start_freq) / (n_points - 1)
        frequencies = [start_freq + i * freq_step for i in range(n_points)]
        
        # 保存到CSV文件
        filename = f"spectrum_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        with open(filename, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Frequency (Hz)', 'Amplitude (dBm)'])
            for freq, amp in zip(frequencies, amplitudes):
                writer.writerow([freq, amp])
        
        print(f"数据已保存到: {filename}")
        
        # 绘制频谱图
        plt.figure(figsize=(12, 6))
        plt.plot([f/1e9 for f in frequencies], amplitudes)
        plt.xlabel('Frequency (GHz)')
        plt.ylabel('Amplitude (dBm)')
        plt.title('Current Spectrum from N9918A')
        plt.grid(True)
        plt.tight_layout()
        plt.show()
        
        # 显示一些统计信息
        max_amp = max(amplitudes)
        min_amp = min(amplitudes)
        avg_amp = sum(amplitudes) / len(amplitudes)
        
        print(f"最大值: {max_amp:.2f} dBm")
        print(f"最小值: {min_amp:.2f} dBm")
        print(f"平均值: {avg_amp:.2f} dBm")
        
        device.close()
        return frequencies, amplitudes
        
    except Exception as e:
        print(f"错误: {e}")
        return None, None

# 运行函数
if __name__ == "__main__":
    frequencies, amplitudes = read_current_spectrum()