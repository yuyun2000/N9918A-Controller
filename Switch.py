import clr
clr.AddReference('mcl_RF_Switch_Controller64')
from mcl_RF_Switch_Controller64 import USB_RF_SwitchBox
class MiniCircuitsSwitchController:
    def __init__(self, serial_number=None):
        self.sw = USB_RF_SwitchBox()
        self.serial_number = serial_number
        self.connected = False
    def connect(self):
        """连接设备"""
        if self.serial_number:
            status = self.sw.Connect(self.serial_number)
        else:
            status = self.sw.Connect()
        if status[0] > 0:
            self.connected = True
            print("设备连接成功")
            return True
        else:
            print("设备连接失败")
            return False
    def disconnect(self):
        """断开设备连接"""
        if self.connected:
            self.sw.Disconnect()
            self.connected = False
            print("设备断开连接")
    def get_model_name(self):
        """获取设备型号"""
        if not self.connected:
            raise Exception("设备未连接")
        result = self.sw.Read_ModelName("")
        return result[1] if result[0] == 1 else None
    def get_serial_number(self):
        """获取设备序列号"""
        if not self.connected:
            raise Exception("设备未连接")
        result = self.sw.Read_SN("")
        return result[1] if result[0] == 1 else None
    def set_switch(self, switch_name, position):
        """
        设置某个开关的状态
        :param switch_name: 开关名称，'A', 'B', 'C', 'D'
        :param position: 状态，1 或 2
        """
        if not self.connected:
            raise Exception("设备未连接")
        if position not in [1, 2]:
            raise ValueError("位置只能是 1 或 2")
        # Mini-Circuits 中 0 表示位置 1，1 表示位置 2
        val = 0 if position == 1 else 1
        result = self.sw.Set_Switch(switch_name, val)
        # result 格式: (status, switch_name, value)
        # status 为 1 表示成功
        if result[0] == 1:
            print(f"开关 {switch_name} 设置为位置 {position}")
        else:
            raise Exception(f"设置开关 {switch_name} 失败，返回结果: {result}")
    def get_switch_status(self):
        """
        获取所有开关的状态（bit 位表示状态）
        根据测试结果：bit = 0 表示位置 1，bit = 1 表示位置 2
        """
        if not self.connected:
            raise Exception("设备未连接")
        result = self.sw.GetSwitchesStatus(0)
        if result[0] == 1:
            status_byte = result[1]
            status_dict = {}
            for i, switch in enumerate(['A', 'B', 'C', 'D']):
                # 读取对应 bit 位
                bit = (status_byte >> i) & 1
                # bit = 0 表示位置 1，bit = 1 表示位置 2
                status_dict[switch] = 1 if bit == 0 else 2
            return status_dict
        else:
            raise Exception("读取开关状态失败")
    def get_temperature(self, sensor=1):
        """
        获取设备温度
        :param sensor: 传感器编号，1 或 2（某些型号支持）
        """
        if not self.connected:
            raise Exception("设备未连接")
        try:
            temp = self.sw.GetDeviceTemperature(sensor)
            return temp
        except Exception as e:
            raise Exception(f"获取温度失败: {e}")
    def get_firmware(self):
        """获取固件版本"""
        if not self.connected:
            raise Exception("设备未连接")
        try:
            fw = self.sw.GetFirmware()
            return fw
        except Exception as e:
            raise Exception(f"获取固件版本失败: {e}")
    def get_usb_status(self):
        """获取 USB 连接状态"""
        if not self.connected:
            raise Exception("设备未连接")
        try:
            status = self.sw.GetUSBConnectionStatus()
            return status
        except Exception as e:
            raise Exception(f"获取 USB 状态失败: {e}")
        
# --- 示例使用 ---
if __name__ == "__main__":
    controller = MiniCircuitsSwitchController()  # 如果有多个设备，可传 serial_number
    if controller.connect():
        print("型号:", controller.get_model_name())
        print("序列号:", controller.get_serial_number())
        # 设置开关
        # controller.set_switch('A', 2)
        # controller.set_switch('B', 2)
        # controller.set_switch('C', 2)
        # controller.set_switch('D', 2)
        controller.set_switch('A', 1)
        controller.set_switch('B', 1)
        controller.set_switch('C', 1)
        controller.set_switch('D', 1)
        # 读取状态
        print("当前开关状态:", controller.get_switch_status())
        # 读取温度
        print("温度:", controller.get_temperature())
        # 固件版本
        print("固件版本:", controller.get_firmware())
        controller.disconnect()