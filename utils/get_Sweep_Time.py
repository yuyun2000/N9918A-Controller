import pyvisa

def read_sweep_time():
    """
    读取频谱仪的扫描时间
    """
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
        
        # 查询扫描时间
        sweep_time = float(device.query("SENS:SWE:TIME?"))
        
        # 将时间转换为毫秒显示
        sweep_time_ms = sweep_time * 1000
        
        print(f"扫描时间: {sweep_time_ms:.0f}ms")
        print(f"扫描时间: {sweep_time:.3f}s")
        
        # 关闭连接
        device.close()
        
        return sweep_time
        
    except Exception as e:
        print(f"错误: {e}")
        return None

# 运行函数
if __name__ == "__main__":
    sweep_time = read_sweep_time()
    if sweep_time is not None:
        print(f"成功读取扫描时间: {sweep_time*1000:.0f}ms")
    else:
        print("读取扫描时间失败")