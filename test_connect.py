import pyvisa

# 使用pyvisa-py后端
try:
    rm = pyvisa.ResourceManager('@py')
    print("可用设备:")
    resources = rm.list_resources()
    print(resources)
except Exception as e:
    print(f"错误: {e}")