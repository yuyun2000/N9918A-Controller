# n9918a_na_controller.py
import pyvisa
import time
import numpy as np
from Switch import MiniCircuitsSwitchController

class N9918ANAController:
    """
    N9918A FieldFox Network Analyzer Controller for NA Mode
    """
    
    def __init__(self, ip_address='192.168.20.233', timeout=10000):
        self.ip_address = ip_address
        self.timeout = timeout
        self.rm = None
        self.device = None
        self.switch_controller = None
        self.connected = False
        self.switch_connected = False
    
    def connect(self):
        """连接到N9918A设备"""
        try:
            self.rm = pyvisa.ResourceManager()
            self.device = self.rm.open_resource(f"TCPIP0::{self.ip_address}::inst0::INSTR")
            self.device.timeout = self.timeout
            
            self.device.write("*CLS")
            device_id = self.device.query("*IDN?")
            print(f"Connected to: {device_id}")
            
            # 选择NA模式
            self.device.write("INST:SEL 'NA'")
            time.sleep(1)
            
            self.connected = True
            print("Successfully connected to N9918A in NA mode")
            return True
            
        except Exception as e:
            print(f"ERROR: Unable to connect to device - {e}")
            self.connected = False
            return False
    
    def connect_switch(self):
        """连接切换器"""
        try:
            self.switch_controller = MiniCircuitsSwitchController()
            if self.switch_controller.connect():
                self.switch_connected = True
                print("Successfully connected to switch controller")
                return True
            else:
                print("ERROR: Unable to connect to switch controller")
                return False
        except Exception as e:
            print(f"ERROR: Failed to connect to switch controller - {e}")
            return False
    
    def disconnect(self):
        """断开设备连接"""
        if self.device:
            self.device.close()
        if self.rm:
            self.rm.close()
        if self.switch_controller and self.switch_connected:
            self.switch_controller.disconnect()
        self.connected = False
        self.switch_connected = False
        print("Disconnected from N9918A and switch controller")
    
    
    def perform_calibration(self):
        """
        执行QuickCal 2端口校准（根据官方SCPI文档），自动完成流程。
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return False

        if not self.switch_connected:
            if not self.connect_switch():
                print("ERROR: Cannot connect to switch controller")
                return False

        try:
            print("🚀 开始2端口QuickCal...")

            # Step 1: 按提示连线，用于你第一个校准状态，通常先连THRU或标准负载
            self.switch_controller.set_switch('B', 2)
            time.sleep(0.5)
            self.switch_controller.set_switch('C', 1)
            time.sleep(0.5)

            # 发送QuickCal命令
            print("📡 发送QuickCal命令: CORR:COLL:METH:QCAL:CAL 1,2")
            self.device.write("CORR:COLL:METH:QCAL:CAL 1,2")

            # 仪器进入校准流程，会依次提示（通常屏幕会提示改连线/换到第二步等）
            # 你可以手动观察仪器，或根据自动化流程等待（这里假定所有换线你用sleep或者人工判断）
            # 更高级的做法是用Sweep Complete/Prompt 查询（具体仪器支持什么SCPI请查帮助）

            print("⏳ 等待仪器流程完成（建议根据实际界面或日志适当加等待）")
            while True:
                opc = self.device.query("*OPC?")
                if opc.strip() == '1':
                    break
                time.sleep(1)
            
            print("✅ 2端口QuickCal完成！")
            return True

        except Exception as e:
            print(f"ERROR: QuickCal校准失败 - {e}")
            return False

    def measure_s11(self):
        """
        测量S11参数并返回频率和幅度数据
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return None, None
            
        try:
            print("🔍 开始测量S11参数")
            
            # 设置测量参数
            self.device.write("CALC:PAR:DEF 'S11',S11")
            time.sleep(0.5)
            
            self.device.write("CALC:FORMat MLOG")
            time.sleep(0.5)
            
            # 触发测量
            self.device.write("INIT:IMM")
            time.sleep(1)
            
            # 获取扫描时间
            try:
                sweep_time = float(self.device.query(":SENS:SWE:TIME?"))
                wait_time = max(sweep_time * 1.2, 1.0)
            except:
                wait_time = 2.0
            
            print(f"⏳ 等待测量完成 ({wait_time:.1f}秒)...")
            time.sleep(wait_time)
            
            # 读取频率数据
            self.device.write("CALC:DATA:STIM?")
            freq_data = self.device.read()
            frequencies = [float(x) for x in freq_data.split(",")]
            
            # 读取S11幅度数据
            self.device.write("CALC:DATA:SNP:DATA?")
            s11_data = self.device.read()
            s11_db = [float(x) for x in s11_data.split(",")][::2]  # 只取实部，跳过虚部
            
            print("✅ S11测量完成")
            return frequencies, s11_db
            
        except Exception as e:
            print(f"ERROR: S11测量失败 - {e}")
            return None, None

def main():
    """
    主函数：演示如何使用N9918ANAController
    """
    # 创建控制器实例
    na_controller = N9918ANAController()
    
    try:
        # 连接设备
        if not na_controller.connect():
            print("无法连接到设备")
            return
        
        # 执行校准
        if not na_controller.perform_calibration():
            print("校准失败")
            na_controller.disconnect()
            return
        
        # 测量S11
        frequencies, s11_db = na_controller.measure_s11()
        
        if frequencies is not None and s11_db is not None:
            print(f"成功获取S11数据，共{len(frequencies)}个点")
            print(f"频率范围: {frequencies[0]/1e9:.3f} GHz - {frequencies[-1]/1e9:.3f} GHz")
            print(f"S11幅度范围: {min(s11_db):.2f} dB - {max(s11_db):.2f} dB")
        else:
            print("S11测量失败")
        
    except Exception as e:
        print(f"发生错误: {e}")
    
    finally:
        # 断开连接
        na_controller.disconnect()

if __name__ == "__main__":
    main()
