# test_na_controller.py
"""
测试N9918A NA控制器的脚本
"""

from n9918a_na_controller import N9918ANAController

def test_na_controller():
    """测试NA控制器的基本功能"""
    print("🧪 开始测试N9918A NA控制器")
    
    # 创建控制器实例
    na_controller = N9918ANAController()
    
    try:
        # 连接设备
        print("\n1. 连接设备...")
        if not na_controller.connect():
            print("❌ 无法连接到N9918A设备")
            return
        
        print("✅ 成功连接到N9918A设备")
        
        # # 连接切换器
        # print("\n2. 连接切换器...")
        # if not na_controller.connect_switch():
        #     print("❌ 无法连接到切换器")
        #     na_controller.disconnect()
        #     return
        
        # print("✅ 成功连接到切换器")
        
        # 执行校准
        print("\n3. 执行校准...")
        if not na_controller.perform_calibration():
            print("❌ 校准失败")
            na_controller.disconnect()
            return
        
        print("✅ 校准完成")
        
        # 等待校准完全结束
        print("⏳ 等待校准完全结束...")
        time.sleep(5)
        
        # 测量S11
        print("\n4. 测量S11参数...")
        frequencies, s11_db = na_controller.measure_s11()
        
        if frequencies is not None and s11_db is not None:
            print(f"✅ 成功获取S11数据，共{len(frequencies)}个点")
            print(f"   频率范围: {frequencies[0]/1e9:.3f} GHz - {frequencies[-1]/1e9:.3f} GHz")
            print(f"   S11幅度范围: {min(s11_db):.2f} dB - {max(s11_db):.2f} dB")
        else:
            print("❌ S11测量失败")
        
    except Exception as e:
        print(f"❌ 测试过程中发生错误: {e}")
    
    finally:
        # 断开连接
        print("\n5. 断开连接...")
        na_controller.disconnect()
        print("✅ 测试完成")

if __name__ == "__main__":
    test_na_controller()
