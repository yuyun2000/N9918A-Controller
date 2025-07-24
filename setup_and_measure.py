import pyvisa
import matplotlib.pyplot as plt
import csv
from datetime import datetime
import time

def safe_query(device, command, default_value="Unknown", timeout=10000):
    """安全查询函数，避免超时"""
    try:
        original_timeout = device.timeout
        device.timeout = timeout
        result = device.query(command).strip()
        device.timeout = original_timeout
        return result
    except Exception as e:
        print(f"⚠️  查询 {command} 失败: {e}")
        return default_value

def safe_write(device, command):
    """安全写入函数"""
    try:
        device.write(command)
        return True
    except Exception as e:
        print(f"❌ 写入 {command} 失败: {e}")
        return False

def wait_for_operation_complete(device, timeout=15):
    """等待设备操作完成"""
    try:
        device.write("*OPC?")  # 操作完成查询
        device.read()
        return True
    except:
        time.sleep(2)  # 如果OPC失败，至少等待2秒
        return True

def setup_n9918a():
    try:
        # 连接到设备
        print("🔌 正在连接到设备...")
        rm = pyvisa.ResourceManager('@py')
        device = rm.open_resource('TCPIP::192.168.20.39::INSTR')
        device.timeout = 20000
        
        print("✅ 连接成功!")
        
        # 查询设备信息
        device_id = safe_query(device, "*IDN?")
        print(f"📊 设备信息: {device_id}")
        
        # 清除状态
        safe_write(device, "*CLS")
        time.sleep(1)
        
        # 切换到频谱分析仪模式
        print("🔄 切换到频谱分析仪模式...")
        safe_write(device, "INST:SEL 'SA'")
        time.sleep(3)  # 给更多时间切换模式
        
        # 逐个设置参数，每个设置后都等待
        print("📡 设置频率范围: 30MHz ~ 1GHz")
        safe_write(device, "SENS:FREQ:STAR 30e6")  # 30MHz
        safe_write(device, "SENS:FREQ:STOP 1e9")   # 1GHz
        time.sleep(2)
        
        print("⚙️  设置RBW: 120kHz")
        safe_write(device, "SENS:BWID:RES 120e3")  # RBW 120kHz
        time.sleep(2)
        
        print("⚙️  设置VBW: 120kHz")
        safe_write(device, "SENS:BWID:VID 120e3")  # VBW 120kHz
        time.sleep(3)  # 给VBW更多时间
        
        print("🎯 设置Detector: Sample")
        safe_write(device, "SENS:DET sample")
        time.sleep(2)
        
        print("🔊 设置内部放大器: ON")
        safe_write(device, "SENS:POW:RF:GAIN:STAT ON")
        time.sleep(2)
        
        print("🔇 设置内部衰减器: 0dB")
        safe_write(device, "SENS:POW:RF:ATT 0")
        time.sleep(2)
        
        print("📈 设置采样点数: 1001")
        safe_write(device, "SENS:SWE:POIN 1001")
        safe_write(device, "SENS:AVER:STAT OFF")
        time.sleep(2)
        
        # 强制重新扫描
        print("🔄 触发重新扫描...")
        safe_write(device, "INIT:IMM")
        time.sleep(5)  # 给足够时间完成扫描
        
        print("\n✅ 设备设置完成!")
        print("请仔细检查设备屏幕:")
        print("  - 频率范围: 30MHz ~ 1GHz")
        print("  - RBW: 120kHz")
        print("  - VBW: 120kHz (这个最重要!)")
        print("  - Detector: Sample")
        print("确认所有参数都正确显示后，按回车键继续测量...")
        
        return device, rm
        
    except Exception as e:
        print(f"❌ 设置失败: {e}")
        return None, None

def perform_measurement(device, rm):
    try:
        print("\n🔍 开始测量...")
        
        # 再次触发扫描确保获取最新数据
        print("🔄 触发新扫描...")
        safe_write(device, "INIT:IMM")
        time.sleep(4)  # 等待扫描完成
        
        # 读取频谱数据
        print("📥 正在读取数据...")
        safe_write(device, "TRACE:DATA?")
        
        # 增加读取超时时间
        original_timeout = device.timeout
        device.timeout = 30000
        trace_data = device.read()
        device.timeout = original_timeout
        
        amplitudes = [float(x) for x in trace_data.split(",")]
        
        # 获取当前设置用于频率计算
        start_freq = float(safe_query(device, "SENS:FREQ:STAR?", "30000000"))
        stop_freq = float(safe_query(device, "SENS:FREQ:STOP?", "1000000000"))
        n_points = int(safe_query(device, "SENS:SWE:POIN?", "1001"))
        
        # 生成频率数组
        freq_step = (stop_freq - start_freq) / (n_points - 1)
        frequencies = [start_freq + i * freq_step for i in range(n_points)]
        
        print(f"📊 读取到 {len(amplitudes)} 个数据点")
        
        # 保存数据
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"spectrum_30MHz_1GHz_{timestamp}.csv"
        with open(filename, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(['Frequency (Hz)', 'Amplitude (dBm)'])
            for freq, amp in zip(frequencies, amplitudes):
                writer.writerow([freq, amp])
        
        print(f"💾 数据已保存到: {filename}")
        
        # 绘制频谱图
        try:
            plt.figure(figsize=(12, 6))
            plt.plot([f/1e6 for f in frequencies], amplitudes, 'b-', linewidth=1)
            plt.xlabel('Frequency (MHz)')
            plt.ylabel('Amplitude (dBm)')
            plt.title('Spectrum Measurement: 30MHz ~ 1GHz')
            plt.grid(True, alpha=0.3)
            plt.tight_layout()
            plt.show()
        except Exception as e:
            print(f"⚠️  绘图失败: {e}")
        
        # 显示统计信息
        if amplitudes:
            max_amp = max(amplitudes)
            min_amp = min(amplitudes)
            avg_amp = sum(amplitudes) / len(amplitudes)
            
            max_freq_idx = amplitudes.index(max_amp)
            max_freq = frequencies[max_freq_idx] / 1e6
            
            print(f"\n📈 测量结果统计:")
            print(f"   最大值: {max_amp:.2f} dBm (在 {max_freq:.2f} MHz)")
            print(f"   最小值: {min_amp:.2f} dBm")
            print(f"   平均值: {avg_amp:.2f} dBm")
        
        return frequencies, amplitudes
        
    except Exception as e:
        print(f"❌ 测量失败: {e}")
        return None, None

def main():
    print("🔧 开始设置N9918A参数...")
    print("=" * 50)
    print("目标参数:")
    print("  频率范围: 30MHz ~ 1GHz")
    print("  RBW: 120kHz")
    print("  VBW: 120kHz")
    print("  Detector: Sample")
    print("  内部放大器: ON")
    print("  内部衰减器: 0dB")
    print("=" * 50)
    
    # 设置设备
    device, rm = setup_n9918a()
    
    if device is not None:
        try:
            # 等待用户确认
            input("\n⚠️  请仔细检查设备屏幕上的所有参数，确认无误后按回车键继续测量...")
            
            # 进行测量
            frequencies, amplitudes = perform_measurement(device, rm)
            
            # 关闭连接
            try:
                device.close()
                rm.close()
                print("\n🔌 设备连接已关闭")
            except:
                print("\n⚠️  设备关闭时出现小问题，但不影响结果")
            
            return frequencies, amplitudes
        except KeyboardInterrupt:
            print("\n🛑 用户取消操作")
            try:
                device.close()
                rm.close()
            except:
                pass
            return None, None
    else:
        print("❌ 无法连接到设备")
        return None, None

# 运行主程序
if __name__ == "__main__":
    frequencies, amplitudes = main()