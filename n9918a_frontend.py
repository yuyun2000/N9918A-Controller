
# n9918a_frontend.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
import matplotlib
from functools import partial
# 设置matplotlib后端和字体
matplotlib.use('TkAgg')
plt.rcParams['font.sans-serif'] = ['Arial', 'DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import threading
import time
from n9918a_backend import N9918AController, get_fcc_ce_limits, post_process_peak_search
from n9918a_backend import save_emi_measurement_data,save_peak_analysis,save_spectrum_data  # 确保导入正确
import os

from Switch import MiniCircuitsSwitchController

class EMCAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("N9918A EMC Analyzer")
        self.root.geometry("1400x900")
        self.root.minsize(1200, 800)
        
        # 创建后端控制器
        self.controller = N9918AController(ip_address='192.168.20.233')

        self.switch_controller = MiniCircuitsSwitchController()
        
        # 当前数据
        self.current_frequencies = None
        self.current_amplitudes = None
        self.current_peaks = None
        self.emi_results = {}  # 存储完整的EMI测量结果（包含采样数据）
        self.sweep_time = 1.0
        self.measurement_in_progress = False
            
        self.root.after(1000, self.connect_switch)
        
        # 创建界面
        self.create_widgets()
        
        # 初始化状态
        self.update_status()
    
    def auto_set_switch_positions(self):
        """根据频率范围自动设置切换器位置"""
        try:
            # 检查设备是否已连接且已配置
            if not self.controller.connected or not self.controller.current_config:
                return
            
            # 获取当前配置的频率范围
            config = self.controller.get_preset_configs().get(self.selected_preset_key, {})
            if not config:
                return
                
            start_freq = config.get("start_freq", 0)
            stop_freq = config.get("stop_freq", 0)
            freq_mhz = stop_freq / 1e6
            
            # 根据频率范围设置切换器位置
            if freq_mhz < 30:  # 小于30MHz
                # A2+D2
                self.switch_controller.set_switch('A', 2)
                self.switch_controller.set_switch('D', 2)
                # B和C保持默认位置1
                self.switch_controller.set_switch('B', 1)
                self.switch_controller.set_switch('C', 1)
                print(f"自动设置切换器: A2+D2 (频率范围: {start_freq/1e6:.3f}MHz - {stop_freq/1e6:.3f}MHz)")
            elif 30 <= freq_mhz <= 3000:  # 30MHz～3GHz
                # A2+D1
                self.switch_controller.set_switch('A', 2)
                self.switch_controller.set_switch('D', 1)
                # B和C保持默认位置1
                self.switch_controller.set_switch('B', 1)
                self.switch_controller.set_switch('C', 1)
                print(f"自动设置切换器: A2+D1 (频率范围: {start_freq/1e6:.3f}MHz - {stop_freq/1e6:.3f}MHz)")
            else:
                # 默认状态 A1+B1+C1+D1
                self.switch_controller.set_switch('A', 1)
                self.switch_controller.set_switch('B', 1)
                self.switch_controller.set_switch('C', 1)
                self.switch_controller.set_switch('D', 1)
                print(f"自动设置切换器: A1+B1+C1+D1 (默认状态)")
            
            # 更新切换器状态显示
            self.update_switch_status()
            
        except Exception as e:
            print(f"自动设置切换器位置时出错: {e}")
            # 即使出错也要确保切换器状态显示更新
            try:
                self.update_switch_status()
            except:
                pass
    
    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 顶部控制面板
        control_frame = ttk.LabelFrame(main_frame, text="Control Panel")
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # IP地址设置
        ip_frame = ttk.Frame(control_frame)
        ip_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(ip_frame, text="Device IP:").pack(side=tk.LEFT)
        self.ip_var = tk.StringVar(value="192.168.20.233")
        ip_entry = ttk.Entry(ip_frame, textvariable=self.ip_var, width=15)
        ip_entry.pack(side=tk.LEFT, padx=5)
        
        self.connect_btn = ttk.Button(ip_frame, text="Connect", command=self.connect_device)
        self.connect_btn.pack(side=tk.LEFT, padx=5)
        
        self.disconnect_btn = ttk.Button(ip_frame, text="Disconnect", command=self.disconnect_device, state=tk.DISABLED)
        self.disconnect_btn.pack(side=tk.LEFT, padx=5)
        
        # 预设配置选择
        preset_frame = ttk.Frame(control_frame)
        preset_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(preset_frame, text="Test Config:").pack(side=tk.LEFT)
        self.preset_var = tk.StringVar()
        self.preset_combo = ttk.Combobox(preset_frame, textvariable=self.preset_var, width=40, state="readonly")
        self.preset_combo.pack(side=tk.LEFT, padx=5)
        
        self.load_presets()
        
        self.config_btn = ttk.Button(preset_frame, text="Configure", command=self.configure_device, state=tk.DISABLED)
        self.config_btn.pack(side=tk.LEFT, padx=5)
        
        # 测量控制
        measure_frame = ttk.Frame(control_frame)
        measure_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.measure_btn = ttk.Button(measure_frame, text="Single", command=self.start_single_measurement, state=tk.DISABLED)
        self.measure_btn.pack(side=tk.LEFT, padx=5)
        
        self.slow_measure_btn = ttk.Button(measure_frame, text="Slow (15s)", command=lambda: self.start_emi_measurement(15), state=tk.DISABLED)
        self.slow_measure_btn.pack(side=tk.LEFT, padx=5)
        
        self.fast_measure_btn = ttk.Button(measure_frame, text="Fast (5min)", command=lambda: self.start_emi_measurement(300), state=tk.DISABLED)
        self.fast_measure_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_measure_btn = ttk.Button(measure_frame, text="Stop", command=self.stop_measurement, state=tk.DISABLED)
        self.stop_measure_btn.pack(side=tk.LEFT, padx=5)
        
        self.save_btn = ttk.Button(measure_frame, text="Save Data", command=self.save_data, state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT, padx=5)
        
        # AI分析按钮
        self.ai_analysis_btn = ttk.Button(measure_frame, text="AI Analysis", command=self.perform_ai_analysis, state=tk.DISABLED)
        self.ai_analysis_btn.pack(side=tk.LEFT, padx=5)
        
        # 导出PDF按钮
        self.export_pdf_btn = ttk.Button(measure_frame, text="Export PDF", command=self.export_pdf, state=tk.DISABLED)
        self.export_pdf_btn.pack(side=tk.LEFT, padx=5)
        
        # 进度条
        self.progress_var = tk.StringVar(value="Ready")
        self.progress_label = ttk.Label(measure_frame, textvariable=self.progress_var)
        self.progress_label.pack(side=tk.RIGHT, padx=5)
        
        # 当前参数显示
        params_frame = ttk.LabelFrame(control_frame, text="Current Parameters")
        params_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.params_text = tk.Text(params_frame, height=3, width=100)
        self.params_text.pack(fill=tk.X, padx=5, pady=5)
        
        # 状态显示
        status_frame = ttk.LabelFrame(main_frame, text="Device Status")
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.status_text = tk.Text(status_frame, height=6, width=80)
        self.status_text.pack(fill=tk.BOTH, padx=5, pady=5)
        
        # 添加切换器控制面板
        switch_frame = ttk.LabelFrame(main_frame, text="Switch Controller")
        switch_frame.pack(fill=tk.X, pady=(0, 10))

        # 控制按钮
        switch_ctrl_frame = ttk.Frame(switch_frame)
        switch_ctrl_frame.pack(fill=tk.X, padx=5, pady=5)

        self.connect_switch_btn = ttk.Button(switch_ctrl_frame, text="Connect Switch", command=self.connect_switch)
        self.connect_switch_btn.pack(side=tk.LEFT, padx=5)

        self.disconnect_switch_btn = ttk.Button(switch_ctrl_frame, text="Disconnect Switch", command=self.disconnect_switch, state=tk.DISABLED)
        self.disconnect_switch_btn.pack(side=tk.LEFT, padx=5)

        # 开关按钮 A/B/C/D
        switch_btn_frame = ttk.Frame(switch_frame)
        switch_btn_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(switch_btn_frame, text="Switch Positions:").pack(anchor='w')

        # 保存开关按钮引用
        self.switch_buttons = {}

        for switch in ['A', 'B', 'C', 'D']:
            frame = ttk.Frame(switch_btn_frame)
            frame.pack(side=tk.LEFT, padx=10)

            label = ttk.Label(frame, text=f"Switch {switch}")
            label.pack()

            btn1 = ttk.Button(frame, text="Pos 1", width=6, command=lambda s=switch: self.set_switch_position(s, 1))
            btn1.pack(side=tk.LEFT, padx=2)
            
            btn2 = ttk.Button(frame, text="Pos 2", width=6, command=lambda s=switch: self.set_switch_position(s, 2))
            btn2.pack(side=tk.LEFT, padx=2)

            self.switch_buttons[switch] = (btn1, btn2)

        # 状态显示区域
        switch_status_frame = ttk.LabelFrame(switch_frame, text="Switch Status")
        switch_status_frame.pack(fill=tk.X, padx=5, pady=5)

        self.switch_status_text = tk.Text(switch_status_frame, height=6, width=80)
        self.switch_status_text.pack(fill=tk.BOTH, padx=5, pady=5)
        # 创建可调整的主显示区域
        self.create_main_display(main_frame)
        
        # 添加用户信息输入区域
        self.create_user_info_panel(main_frame)
        
        # 添加调整比例的滑块
        self.create_ratio_control(main_frame)

    #---------------------------------切换器相关函数
    def connect_switch(self):
        """连接切换器"""
        try:
            if self.switch_controller.connect():
                self.connect_switch_btn.config(state=tk.DISABLED)
                self.disconnect_switch_btn.config(state=tk.NORMAL)
                self.update_switch_status()
                messagebox.showinfo("Success", "Switch connected successfully!")
            else:
                messagebox.showerror("Error", "Failed to connect to switch.")
        except Exception as e:
            messagebox.showerror("Error", f"Error connecting to switch:\n{e}")

    def disconnect_switch(self):
        """断开切换器"""
        try:
            self.switch_controller.disconnect()
            self.connect_switch_btn.config(state=tk.NORMAL)
            self.disconnect_switch_btn.config(state=tk.DISABLED)
            self.switch_status_text.delete(1.0, tk.END)
            self.switch_status_text.insert(tk.END, "Switch disconnected.\n")
            messagebox.showinfo("Info", "Switch disconnected.")
        except Exception as e:
            messagebox.showerror("Error", f"Error disconnecting switch:\n{e}")

    def set_switch_position(self, switch, position):
        """设置某个开关的位置"""
        try:
            self.switch_controller.set_switch(switch, position)
            self.update_switch_status()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to set switch {switch} to position {position}:\n{e}")

    def update_switch_status(self):
        """更新切换器状态显示"""
        try:
            self.switch_status_text.delete(1.0, tk.END)
            
            # 获取开关状态
            switch_status = self.switch_controller.get_switch_status()
            model = self.switch_controller.get_model_name()
            serial = self.switch_controller.get_serial_number()
            firmware = self.switch_controller.get_firmware()
            temp = self.switch_controller.get_temperature()
            usb_status = self.switch_controller.get_usb_status()

            self.switch_status_text.insert(tk.END, f"Model: {model}\n")
            self.switch_status_text.insert(tk.END, f"Serial: {serial}\n")
            self.switch_status_text.insert(tk.END, f"Firmware: {firmware}\n")
            self.switch_status_text.insert(tk.END, f"Temperature: {temp} °C\n")
            self.switch_status_text.insert(tk.END, f"USB Status: {usb_status}\n\n")
            self.switch_status_text.insert(tk.END, "Switch Positions:\n")
            for switch, pos in switch_status.items():
                self.switch_status_text.insert(tk.END, f"  {switch}: Position {pos}\n")

            # 更新按钮状态
            for switch, pos in switch_status.items():
                btn1, btn2 = self.switch_buttons[switch]
                btn1.config(state=tk.NORMAL if pos != 1 else tk.DISABLED)
                btn2.config(state=tk.NORMAL if pos != 2 else tk.DISABLED)

        except Exception as e:
            self.switch_status_text.delete(1.0, tk.END)
            self.switch_status_text.insert(tk.END, f"Error reading switch status:\n{e}\n")

    #------------------------------------频谱仪相关

    def start_emi_measurement(self, duration):
        """开始快速EMC测量（只采集一次，PC端计算多种模式）"""
        self.measurement_in_progress = True
        
        def emi_measure_task():
            try:
                # 禁用所有测量按钮
                self.root.after(0, lambda: self.set_all_measure_buttons(False))
                self.root.after(0, lambda: self.progress_var.set(f"Fast EMI Measurement ({duration}s)..."))
                
                # 获取快速EMC测量数据
                results = self.controller.get_emc_measurement_fast(duration)
                
                if results:
                    self.root.after(0, lambda: self.on_emi_measurement_complete(results))
                else:
                    self.root.after(0, self.on_measurement_failed)
                    
            except Exception as e:
                self.root.after(0, partial(self.on_measurement_error, str(e)))
            finally:
                self.measurement_in_progress = False
                self.root.after(0, lambda: self.set_all_measure_buttons(True))
                self.root.after(0, lambda: self.progress_var.set("Ready"))
        
        threading.Thread(target=emi_measure_task, daemon=True).start()

    def set_all_measure_buttons(self, enabled):
        """设置所有测量按钮状态"""
        state = tk.NORMAL if enabled and self.controller.connected else tk.DISABLED
        stop_state = tk.NORMAL if not enabled and self.measurement_in_progress else tk.DISABLED
        
        self.measure_btn.config(state=state)
        self.slow_measure_btn.config(state=state)
        self.fast_measure_btn.config(state=state)
        self.stop_measure_btn.config(state=stop_state)

    def stop_measurement(self):
        """停止测量"""
        try:
            if self.controller.connected and self.controller.device:
                self.controller.device.write("INIT:CONT OFF")
                self.measurement_in_progress = False
                self.progress_var.set("Measurement stopped")
                self.set_all_measure_buttons(True)
                messagebox.showinfo("Info", "Measurement stopped")
        except Exception as e:
            print(f"Error stopping measurement: {e}")

    def update_emi_plot(self):
        """更新EMI图形显示"""
        if not self.emi_results:
            return
        
        try:
            # 清除旧图形
            self.ax.clear()
            
            # 颜色和线型映射
            styles = {
                "PEAK": ('red', '-', 1.5),           # 红色实线
                "QUASI_PEAK": ('blue', '--', 1.2),   # 蓝色虚线
                "AVERAGE": ('green', '-.', 1.0)      # 绿色点划线
            }
            
            # 绘制每种检测器模式的数据
            for mode, data in self.emi_results.items():
                if mode in ["measurement_summary", "sampling_info"]:
                    continue
                    
                if isinstance(data, tuple) and len(data) >= 2:
                    frequencies, amplitudes = data[0], data[1]
                    if frequencies is None or amplitudes is None:
                        continue
                        
                    freq_mhz = [f/1e6 for f in frequencies]
                    color, linestyle, linewidth = styles.get(mode, ('black', '-', 1.0))
                    
                    # 绘制线条
                    self.ax.semilogx(freq_mhz, amplitudes, 
                                color=color, linewidth=linewidth, 
                                linestyle=linestyle, alpha=0.7, label=f'{mode}')
                    
                    # 为PEAK模式标记峰值点
                    if mode == "PEAK":
                        peaks = post_process_peak_search(frequencies, amplitudes)
                        for peak in peaks:
                            freq_mhz_peak = peak['frequency_mhz']
                            amp_dbuv = peak['amplitude_dbuv']
                            exceed_fcc = peak['exceed_fcc']
                            exceed_ce = peak['exceed_ce']
                            marker_color = 'red' if exceed_fcc or exceed_ce else 'orange'
                            self.ax.plot(freq_mhz_peak, amp_dbuv, 'D',  # 菱形标记
                                    color=marker_color, markersize=8, 
                                    markeredgecolor='black', markeredgewidth=1)
            
            # 绘制FCC和CE限值
            if self.current_frequencies:
                fcc_limits = []
                ce_limits = []
                for freq in self.current_frequencies:
                    fcc_limit, ce_limit = get_fcc_ce_limits(freq)
                    fcc_limits.append(fcc_limit)
                    ce_limits.append(ce_limit)
                
                freq_mhz_limits = [f/1e6 for f in self.current_frequencies]
                self.ax.semilogx(freq_mhz_limits, fcc_limits, 
                            'magenta', linewidth=2, linestyle=':', 
                            alpha=0.8, label='FCC Limit')
                self.ax.semilogx(freq_mhz_limits, ce_limits, 
                            'cyan', linewidth=2, linestyle=':', 
                            alpha=0.8, label='CE Limit')
            
            # 设置标签和标题
            self.ax.set_xlabel('Frequency (MHz)', fontsize=10)
            self.ax.set_ylabel('Amplitude (dBμV)', fontsize=10)
            self.ax.set_title('Fast EMI Spectrum Analysis (Computed Modes)', fontsize=12)
            self.ax.grid(True, which="both", alpha=0.3)
            self.ax.legend(fontsize=8, loc='upper right')
            
            # 改进横轴标签显示
            self.improve_xaxis_labels_detailed()
            
            # 设置坐标轴范围
            if self.current_frequencies:
                freq_mhz = [f/1e6 for f in self.current_frequencies]
                self.ax.set_xlim([min(freq_mhz), max(freq_mhz)])
            
            # 设置纵轴范围
            if self.emi_results:
                all_amplitudes = []
                for mode, data in self.emi_results.items():
                    if mode in ["measurement_summary", "sampling_info"]:
                        continue
                    if isinstance(data, tuple) and len(data) >= 2 and data[1]:
                        all_amplitudes.extend(data[1])
                if all_amplitudes:
                    # 添加限值数据
                    if 'fcc_limits' in locals():
                        all_amplitudes.extend(fcc_limits)
                        all_amplitudes.extend(ce_limits)
                    if all_amplitudes:
                        y_min = min(all_amplitudes) - 10
                        y_max = max(all_amplitudes) + 10
                        self.ax.set_ylim([y_min, y_max])
            
            # 更新画布
            self.fig.tight_layout()
            self.canvas.draw()
            
        except Exception as e:
            print(f"Plot error: {e}")
            import traceback
            traceback.print_exc()
            self.ax.clear()
            self.ax.set_title("Error displaying plot")
            self.ax.grid(True, alpha=0.3)
            self.canvas.draw()

    def improve_xaxis_labels_detailed(self):
        """改进横轴标签显示 - 更密集的标签"""
        if not hasattr(self, 'current_frequencies') or self.current_frequencies is None:
            return
            
        freq_mhz = [f/1e6 for f in self.current_frequencies]
        min_freq = min(freq_mhz)
        max_freq = max(freq_mhz)
        
        # 生成更密集的频率标签
        def generate_ticks(min_f, max_f):
            ticks = []
            labels = []
            
            # 根据范围选择合适的步进
            if max_f <= 1:  # kHz范围 (优化9k-150k)
                if min_f >= 0.009 and max_f <= 0.150:  # 9k-150k范围
                    # 特别优化9k-150k范围
                    major_ticks = [0.009, 0.01, 0.03, 0.05, 0.1, 0.150]
                    major_labels = ['9k', '10k', '30k', '50k', '100k', '150k']
                    
                    minor_ticks = [0.015, 0.02, 0.025, 0.04, 0.06, 0.07, 0.08, 0.09, 0.12, 0.13]
                    minor_labels = ['15k', '20k', '25k', '40k', '60k', '70k', '80k', '90k', '120k', '130k']
                    
                    # 添加主要刻度
                    for tick, label in zip(major_ticks, major_labels):
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(label)
                    
                    # 添加中间刻度
                    for tick, label in zip(minor_ticks, minor_labels):
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(label)
                    
                    # 排序
                    combined = list(zip(ticks, labels))
                    combined.sort(key=lambda x: x[0])
                    ticks, labels = zip(*combined) if combined else ([], [])
                    ticks, labels = list(ticks), list(labels)
                else:
                    # 原来的逻辑
                    base_ticks = [0.01, 0.03, 0.1, 0.3, 1.0]
                    for tick in base_ticks:
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            if tick >= 1:
                                labels.append(f'{tick:.0f}M')
                            elif tick >= 0.001:
                                labels.append(f'{tick*1000:.0f}k')
                            else:
                                labels.append(f'{tick*1000000:.0f}')
            elif max_f <= 30:  # MHz范围 (优化150k-30M)
                if min_f >= 0.150 and max_f <= 30:  # 150k-30M范围
                    # 特别优化150k-30M范围
                    major_ticks = [0.150, 0.3, 0.5, 1.0, 2.0, 5.0, 10.0, 20.0, 30.0]
                    major_labels = ['150k', '300k', '500k', '1M', '2M', '5M', '10M', '20M', '30M']
                    
                    minor_ticks = [0.2, 0.4, 0.6, 0.7, 0.8, 0.9, 1.5, 2.5, 3.0, 4.0, 6.0, 7.0, 8.0, 9.0, 15.0, 25.0]
                    minor_labels = ['200k', '400k', '600k', '700k', '800k', '900k', '1.5M', '2.5M', '3M', '4M', '6M', '7M', '8M', '9M', '15M', '25M']
                    
                    # 添加主要刻度
                    for tick, label in zip(major_ticks, major_labels):
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(label)
                    
                    # 添加中间刻度
                    for tick, label in zip(minor_ticks, minor_labels):
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(label)
                    
                    # 排序
                    combined = list(zip(ticks, labels))
                    combined.sort(key=lambda x: x[0])
                    ticks, labels = zip(*combined) if combined else ([], [])
                    ticks, labels = list(ticks), list(labels)
                else:
                    # 原来的逻辑
                    current = max(0.1, min_f)
                    while current <= max_f:
                        ticks.append(current)
                        if current >= 1:
                            labels.append(f'{current:.0f}M')
                        else:
                            labels.append(f'{current*1000:.0f}k')
                        if current < 0.1:
                            current += 0.01
                        elif current < 1:
                            current += 0.1
                        else:
                            current += 0.5
                        # 避免无限循环
                        if current > max_f * 2:
                            break
            elif max_f <= 100:  # MHz范围
                # 更智能的刻度生成
                range_freq = max_f - min_f
                if range_freq <= 20:  # 小范围，更密集
                    step = 1.0
                elif range_freq <= 50:
                    step = 2.0
                else:
                    step = 5.0
                    
                current = max(step, min_f)
                while current <= max_f:
                    if current >= min_f:
                        ticks.append(current)
                        labels.append(f'{current:.0f}M')
                    current += step
            elif max_f <= 1000:  # GHz范围 (30M-1G是主要关注点)
                # 特别优化30M-1G范围
                range_freq = max_f - min_f
                if min_f >= 30 and max_f <= 1000:  # 30M-1G范围
                    # 在这个范围内生成更密集的刻度
                    major_ticks = [30, 50, 100, 200, 300, 500, 1000]
                    for tick in major_ticks:
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            if tick >= 1000:
                                labels.append(f'{tick/1000:.1f}G')
                            else:
                                labels.append(f'{tick}M')
                    
                    # 添加中间刻度
                    minor_ticks = [40, 60, 70, 80, 90, 150, 250, 400, 600, 800]
                    for tick in minor_ticks:
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(f'{tick}M')
                    
                    # 排序
                    combined = list(zip(ticks, labels))
                    combined.sort(key=lambda x: x[0])
                    ticks, labels = zip(*combined) if combined else ([], [])
                    ticks, labels = list(ticks), list(labels)
                else:
                    # 原来的逻辑
                    step = max(50.0, range_freq / 8)
                    current = max(step, min_f)
                    while current <= max_f:
                        ticks.append(current)
                        if current >= 1000:
                            labels.append(f'{current/1000:.1f}G')
                        else:
                            labels.append(f'{current:.0f}M')
                        current += step
            else:  # 更高频率
                range_freq = max_f - min_f
                step = max(100.0, range_freq / 6)
                current = max(step, min_f)
                while current <= max_f:
                    ticks.append(current)
                    if current >= 1000:
                        labels.append(f'{current/1000:.1f}G')
                    else:
                        labels.append(f'{current:.0f}M')
                    current += step
            
            return ticks, labels
        
        # 生成标签
        ticks, labels = generate_ticks(min_freq, max_freq)
        
        # 设置自定义标签
        if ticks:
            self.ax.set_xticks(ticks)
            self.ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=8)
        
        # 确保纵轴有数字标签
        self.ax.yaxis.set_major_locator(plt.MaxNLocator(nbins=10, integer=False))
        self.ax.tick_params(axis='y', which='major', labelsize=9)
        self.ax.tick_params(axis='x', which='major', labelsize=8)

    def update_emi_peak_display(self):
        """更新EMI峰值显示"""
        self.peak_text.delete(1.0, tk.END)
        
        if not self.emi_results:
            self.peak_text.insert(tk.END, "No EMI measurement data available\n")
            return
        
        # 只显示QUASI_PEAK模式的结果
        mode = "QUASI_PEAK"
        if mode in self.emi_results:
            data = self.emi_results[mode]
            if isinstance(data, tuple) and len(data) >= 2:
                frequencies, amplitudes = data[0], data[1]
                if amplitudes is not None:
                    self.peak_text.insert(tk.END, f"\n{mode} Mode Results:\n")
                    self.peak_text.insert(tk.END, "="*100 + "\n")
                    header = f"{'No':<4} {'Freq [MHz]':<12} {'Amplitude [dBμV]':<18} {'FCC Limit [dBμV]':<18} {'FCC Margin [dB]':<18} {'CE Limit [dBμV]':<18} {'CE Margin [dB]':<18} {'Status':<15}\n"
                    separator = "-"*130 + "\n"
                    
                    self.peak_text.insert(tk.END, header)
                    self.peak_text.insert(tk.END, separator)
                    
                    # 分析峰值
                    peaks = post_process_peak_search(frequencies, amplitudes)
                    
                    # 按照要求排序：首先是超标的排前面，然后是余量最少的，最后是余量最多的
                    exceed_peaks = [p for p in peaks if p['exceed_fcc']]
                    normal_peaks = [p for p in peaks if not p['exceed_fcc']]
                    
                    # 对超标的峰值按余量排序（余量越少越前）
                    exceed_peaks.sort(key=lambda x: x['fcc_margin'],reverse=True)
                    
                    # 对正常的峰值按余量排序（余量越少越前）
                    normal_peaks.sort(key=lambda x: x['fcc_margin'],reverse=True)
                    
                    # 合并所有峰值，确保左下图中显示的所有peak都在表格中标出来
                    all_peaks = exceed_peaks + normal_peaks
                    
                    # 显示所有峰值（添加序号）
                    for i, peak in enumerate(all_peaks, 1):
                        status = []
                        if peak['exceed_fcc']:
                            status.append("FCC Fail")
                        if peak['exceed_ce']:
                            status.append("CE Fail")
                        if not status:
                            status = ["Pass"]
                        
                        line = f"{i:<4} "
                        line += f"{peak['frequency_mhz']:<12.3f} "
                        line += f"{peak['amplitude_dbuv']:<18.2f} "
                        line += f"{peak['fcc_limit']:<18.1f} "
                        line += f"{peak['fcc_margin']:<18.2f} "
                        line += f"{peak['ce_limit']:<18.1f} "
                        line += f"{peak['ce_margin']:<18.2f} "
                        line += f"{', '.join(status):<15}\n"
                        
                        self.peak_text.insert(tk.END, line)
    def on_emi_measurement_complete(self, results):
        """EMI测量完成回调 - 快速版本"""
        self.emi_results = results
        self.progress_var.set("Fast EMI Measurement completed")
        self.save_btn.config(state=tk.NORMAL)
        self.ai_analysis_btn.config(state=tk.NORMAL)  # 启用AI分析按钮
        # 检查是否是15秒或5分钟测试，如果是则启用导出PDF按钮
        if "measurement_summary" in results:
            summary = results["measurement_summary"]
            duration = summary.get("actual_measurement_time", 0)
            if duration == 15 or duration == 300:  # 15秒或300秒(5分钟)
                self.export_pdf_btn.config(state=tk.NORMAL)
        
        # 显示测量摘要
        if "measurement_summary" in results:
            summary = results["measurement_summary"]
            summary_msg = "Fast EMI Measurement Completed!\n\n"
            summary_msg += f"Total Time: {summary['total_duration']:.1f} seconds\n"
            summary_msg += f"Measurement Time: {summary['actual_measurement_time']} seconds\n"
            summary_msg += f"Data Points: {summary['data_points']}\n"
            summary_msg += f"Modes Computed: {', '.join(summary['modes_computed'])}\n"
            summary_msg += f"Completed at: {summary['measurement_time']}"
            messagebox.showinfo("Fast EMI Measurement Summary", summary_msg)
        
        # 使用PEAK数据作为主显示
        computed_modes = ["PEAK", "QUASI_PEAK", "AVERAGE"]
        for mode in computed_modes:
            if mode in results:
                frequencies, amplitudes = results[mode]
                if mode == "PEAK":
                    self.current_frequencies = frequencies
                    self.current_amplitudes = amplitudes
                    self.current_peaks = post_process_peak_search(frequencies, amplitudes)
                break
        else:
            # 如果没有找到标准模式，使用第一个可用的数据
            for mode, data in results.items():
                if mode not in ["measurement_summary", "sampling_info"]:
                    if isinstance(data, tuple) and len(data) >= 2:
                        frequencies, amplitudes = data[0], data[1]
                        self.current_frequencies = frequencies
                        self.current_amplitudes = amplitudes
                        self.current_peaks = post_process_peak_search(frequencies, amplitudes)
                        break
        
        # 更新图形显示多种检测器模式
        self.update_emi_plot()
        
        # 更新峰值显示
        self.update_emi_peak_display()
        
        # 显示测量结果摘要
        self.show_measurement_summary()
    

    def show_measurement_summary(self):
        """显示测量结果摘要"""
        if not self.emi_results:
            return
        
        summary = "Fast EMI Measurement Results Summary:\n"
        summary += "="*60 + "\n"
        
        # 显示采样信息
        if "sampling_info" in self.emi_results:
            info = self.emi_results["sampling_info"]
            summary += f"采集信息:\n"
            summary += f"  采样次数: {info['total_samples']}\n"
            summary += f"  测量时间: {info['sample_duration']} 秒\n"
            summary += f"  数据点数: {info['data_points']}\n"
            summary += "-"*40 + "\n"
        
        # 显示各模式结果
        computed_modes = ["PEAK", "QUASI_PEAK", "AVERAGE"]
        for mode in computed_modes:
            if mode in self.emi_results:
                data = self.emi_results[mode]
                if isinstance(data, tuple) and len(data) >= 2:
                    frequencies, amplitudes = data[0], data[1]
                    if amplitudes:
                        max_val = max(amplitudes)
                        min_val = min(amplitudes)
                        avg_val = sum(amplitudes) / len(amplitudes)
                        summary += f"{mode} 模式:\n"
                        summary += f"  最大值: {max_val:.2f} dBμV\n"
                        summary += f"  最小值: {min_val:.2f} dBμV\n"
                        summary += f"  平均值: {avg_val:.2f} dBμV\n"
                        summary += "-"*40 + "\n"
        
        print(summary)  # 在控制台也打印摘要

    def start_single_measurement(self):
        """开始单次测量"""
        def single_measure_task():
            try:
                self.root.after(0, lambda: self.set_all_measure_buttons(False))
                self.root.after(0, lambda: self.progress_var.set("Single Measurement..."))
                
                frequencies, amplitudes = self.controller.read_trace_data()
                if frequencies is not None and amplitudes is not None:
                    peaks = post_process_peak_search(frequencies, amplitudes)
                    self.root.after(0, lambda: self.on_measurement_complete(frequencies, amplitudes, peaks))
                else:
                    self.root.after(0, self.on_measurement_failed)
                    
            except Exception as ex:  # 改变变量名避免冲突
                self.root.after(0, lambda: self.on_measurement_error(str(ex)))
            finally:
                self.root.after(0, lambda: self.set_all_measure_buttons(True))
                self.root.after(0, lambda: self.progress_var.set("Ready"))
        
        threading.Thread(target=single_measure_task, daemon=True).start()




    def set_measurement_buttons_state(self, enabled):
        """设置测量按钮状态"""
        state = tk.NORMAL if enabled else tk.DISABLED
        self.measure_btn.config(state=state)
        self.slow_measure_btn.config(state=state)
        self.fast_measure_btn.config(state=state)
        self.stop_measure_btn.config(state=tk.DISABLED if enabled else tk.NORMAL)

    
    def create_main_display(self, parent):
        """创建主显示区域"""
        # 创建可调整的PanedWindow
        self.main_paned = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # 左侧图形显示区域
        self.plot_frame = ttk.LabelFrame(self.main_paned, text="Spectrum Plot")
        self.main_paned.add(self.plot_frame, weight=3)
        
        # 右侧峰值显示区域
        self.peak_frame = ttk.LabelFrame(self.main_paned, text="Peak Analysis Results")
        self.main_paned.add(self.peak_frame, weight=2)
        
        # 创建图形画布
        self.create_plot_canvas(self.plot_frame)
        
        # 创建峰值显示区域
        self.create_peak_display(self.peak_frame)
    
    def create_user_info_panel(self, parent):
        """创建用户信息输入面板"""
        # 创建用户信息框架
        user_info_frame = ttk.LabelFrame(parent, text="User Information for PDF Report")
        user_info_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 创建输入框框架
        input_frame = ttk.Frame(user_info_frame)
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 客户名称
        customer_frame = ttk.Frame(input_frame)
        customer_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(customer_frame, text="Customer:").pack(anchor='w')
        self.customer_var = tk.StringVar(value="M5Stack")
        ttk.Entry(customer_frame, textvariable=self.customer_var, width=15).pack()
        
        # EUT
        eut_frame = ttk.Frame(input_frame)
        eut_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(eut_frame, text="EUT:").pack(anchor='w')
        self.eut_var = tk.StringVar(value="产品A")
        ttk.Entry(eut_frame, textvariable=self.eut_var, width=15).pack()
        
        # 型号
        model_frame = ttk.Frame(input_frame)
        model_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(model_frame, text="Model:").pack(anchor='w')
        self.model_var = tk.StringVar(value="Model-X")
        ttk.Entry(model_frame, textvariable=self.model_var, width=15).pack()
        
        # 工程师
        engineer_frame = ttk.Frame(input_frame)
        engineer_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(engineer_frame, text="Engineer:").pack(anchor='w')
        self.engineer_var = tk.StringVar(value="张工程师")
        ttk.Entry(engineer_frame, textvariable=self.engineer_var, width=15).pack()
        
        # 备注
        remark_frame = ttk.Frame(input_frame)
        remark_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(remark_frame, text="Remark:").pack(anchor='w')
        self.remark_var = tk.StringVar(value="首次测试")
        ttk.Entry(remark_frame, textvariable=self.remark_var, width=15).pack()
    
    def create_ratio_control(self, parent):
        """创建比例控制滑块"""
        ratio_frame = ttk.Frame(parent)
        ratio_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Label(ratio_frame, text="Adjust Layout Ratio:").pack(side=tk.LEFT, padx=(0, 10))
        
        self.ratio_var = tk.DoubleVar(value=60)  # 默认60%给图形
        ratio_scale = ttk.Scale(ratio_frame, from_=30, to=70, variable=self.ratio_var, 
                               orient=tk.HORIZONTAL, command=self.on_ratio_change)
        ratio_scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        self.ratio_label = ttk.Label(ratio_frame, text="Graph: 60% | Data: 40%")
        self.ratio_label.pack(side=tk.LEFT)
    
    def on_ratio_change(self, value):
        """比例改变时的回调"""
        # 注意：tkinter PanedWindow的动态权重调整比较复杂
        # 这里主要是更新显示标签
        ratio = int(float(value))
        self.ratio_label.config(text=f"Graph: {ratio}% | Data: {100-ratio}%")
    
    def create_plot_canvas(self, parent):
        """创建图形画布"""
        # 创建图形容器
        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建matplotlib图形
        self.fig, self.ax = plt.subplots(figsize=(10, 6))
        self.ax.set_title("Waiting for measurement data...", fontsize=12)
        self.ax.grid(True, alpha=0.3)
        self.ax.set_xlabel('Frequency (MHz)', fontsize=10)
        self.ax.set_ylabel('Amplitude (dBμV)', fontsize=10)
        
        # 创建画布
        self.canvas = FigureCanvasTkAgg(self.fig, canvas_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 设置图形适应
        self.fig.tight_layout()
    
    def create_peak_display(self, parent):
        """创建峰值显示区域"""
        # 创建容器
        container = ttk.Frame(parent)
        container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建文本框和滚动条
        self.peak_text = tk.Text(container, height=15, width=60)
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self.peak_text.yview)
        self.peak_text.configure(yscrollcommand=scrollbar.set)
        
        self.peak_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 初始化显示
        self.peak_text.insert(tk.END, "No peak data available\n")
    
    def load_presets(self):
        """加载预设配置"""
        presets = self.controller.get_preset_configs()
        preset_names = [f"{config['name']} - {config['description']}" for config in presets.values()]
        preset_keys = list(presets.keys())
        
        self.preset_combo['values'] = preset_names
        if preset_names:
            self.preset_combo.current(0)
            self.selected_preset_key = preset_keys[0]
        
        # 绑定选择事件
        def on_preset_select(event):
            selected_index = self.preset_combo.current()
            if selected_index >= 0:
                self.selected_preset_key = preset_keys[selected_index]
        
        self.preset_combo.bind('<<ComboboxSelected>>', on_preset_select)
    
    def update_status(self):
        """更新状态显示"""
        self.status_text.delete(1.0, tk.END)
        status = self.controller.get_current_status()
        
        for key, value in status.items():
            if key == "start_freq" and value is not None:
                value = f"{value/1e6:.3f} MHz"
            elif key == "stop_freq" and value is not None:
                value = f"{value/1e6:.3f} MHz"
            elif key == "rbw" and value is not None:
                value = f"{value} Hz"
            elif key == "vbw" and value is not None:
                value = f"{value} Hz"
            self.status_text.insert(tk.END, f"{key}: {value}\n")
    
    def update_params_display(self):
        """更新参数显示"""
        self.params_text.delete(1.0, tk.END)
        status = self.controller.get_current_status()
        
        if status.get("status") == "已连接" and status.get("current_config") != "未配置":
            config = self.controller.get_preset_configs().get(self.selected_preset_key, {})
            if config:
                params_info = f"Config: {config.get('name', 'N/A')} | "
                if status.get('start_freq'):
                    params_info += f"Freq: {status.get('start_freq')/1e6:.3f}-{status.get('stop_freq')/1e6:.3f} MHz | "
                if status.get('rbw'):
                    params_info += f"RBW: {status.get('rbw')} Hz | VBW: {status.get('vbw')} Hz | "
                params_info += f"Points: {status.get('n_points', 0)}"
                self.params_text.insert(tk.END, params_info)
    
    def get_sweep_time(self):
        """获取扫描时间"""
        try:
            if self.controller.connected and self.controller.device:
                # 尝试获取设备的扫描时间
                sweep_time_str = self.controller.device.query(":SENS:SWE:TIME?")
                sweep_time = float(sweep_time_str)
                return max(sweep_time, 0.5)  # 至少等待0.5秒
        except:
            pass
        # 默认返回基于频率范围的估算时间
        if self.controller.start_freq and self.controller.stop_freq:
            freq_range = (self.controller.stop_freq - self.controller.start_freq) / 1e9
            return max(1.0, freq_range * 2)
        return 2.0
    
    def connect_device(self):
        """连接设备"""
        ip_address = self.ip_var.get()
        self.controller.ip_address = ip_address
        
        def connect_task():
            try:
                self.connect_btn.config(state=tk.DISABLED, text="Connecting...")
                self.root.update()
                
                if self.controller.connect():
                    self.root.after(0, self.on_connected)
                else:
                    self.root.after(0, self.on_connect_failed)
            except Exception as e:
                self.root.after(0, lambda: self.on_connect_error(str(e)))
        
        threading.Thread(target=connect_task, daemon=True).start()
    
    def on_connected(self):
        """连接成功回调"""
        self.connect_btn.config(state=tk.DISABLED)
        self.disconnect_btn.config(state=tk.NORMAL)
        self.config_btn.config(state=tk.NORMAL)
        self.update_status()
        messagebox.showinfo("Success", "Device connected successfully!")
    
    def on_connect_failed(self):
        """连接失败回调"""
        self.connect_btn.config(state=tk.NORMAL, text="Connect")
        messagebox.showerror("Error", "Failed to connect to device!")
    
    def on_connect_error(self, error_msg):
        """连接错误回调"""
        self.connect_btn.config(state=tk.NORMAL, text="Connect")
        messagebox.showerror("Error", f"Error connecting to device:\n{error_msg}")
    
    def disconnect_device(self):
        """断开设备连接"""
        self.controller.disconnect()
        self.connect_btn.config(state=tk.NORMAL)
        self.disconnect_btn.config(state=tk.DISABLED)
        self.config_btn.config(state=tk.DISABLED)
        self.measure_btn.config(state=tk.DISABLED)
        self.save_btn.config(state=tk.DISABLED)
        self.export_pdf_btn.config(state=tk.DISABLED)  # 重置导出PDF按钮状态
        self.update_status()
        self.params_text.delete(1.0, tk.END)
        self.peak_text.delete(1.0, tk.END)
        # 清除图形
        self.ax.clear()
        self.ax.set_title("Waiting for measurement data...")
        self.ax.grid(True, alpha=0.3)
        self.ax.set_xlabel('Frequency (MHz)', fontsize=10)
        self.ax.set_ylabel('Amplitude (dBμV)', fontsize=10)
        self.canvas.draw()
        messagebox.showinfo("Info", "Device disconnected")
    
    def configure_device(self):
        """配置设备"""
        if not self.controller.connected:
            messagebox.showwarning("Warning", "Please connect to device first!")
            return
        
        def config_task():
            try:
                self.config_btn.config(state=tk.DISABLED, text="Configuring...")
                self.root.update()
                
                if self.controller.configure_settings(self.selected_preset_key):
                    self.root.after(0, self.on_configured)
                else:
                    self.root.after(0, self.on_config_failed)
            except Exception as e:
                self.root.after(0, lambda: self.on_config_error(str(e)))
        
        threading.Thread(target=config_task, daemon=True).start()
    
    def on_configured(self):
        """配置成功回调"""
        self.config_btn.config(state=tk.NORMAL, text="Configure")
        self.measure_btn.config(state=tk.NORMAL)
        self.update_status()
        self.update_params_display()
        
        # 自动设置切换器位置
        self.auto_set_switch_positions()
        
        messagebox.showinfo("Success", "Device configured successfully!")
    
    def on_config_failed(self):
        """配置失败回调"""
        self.config_btn.config(state=tk.NORMAL, text="Configure")
        messagebox.showerror("Error", "Failed to configure device!")
    
    def on_config_error(self, error_msg):
        """配置错误回调"""
        self.config_btn.config(state=tk.NORMAL, text="Configure")
        messagebox.showerror("Error", f"Error configuring device:\n{error_msg}")
    

    def on_measurement_complete(self, frequencies, amplitudes, peaks, multi_results=None):
        """测量完成回调 - 增强版本"""
        self.set_measurement_buttons_state(True)
        self.progress_var.set("Measurement completed")
        self.save_btn.config(state=tk.NORMAL)
        self.ai_analysis_btn.config(state=tk.NORMAL)  # 启用AI分析按钮
        self.export_pdf_btn.config(state=tk.DISABLED)  # 禁用导出PDF按钮（仅15秒和5分钟测试可以导出）
        
        # 保存数据
        self.current_frequencies = frequencies
        self.current_amplitudes = amplitudes
        self.current_peaks = peaks
        self.multi_detector_results = multi_results  # 保存多种检测器结果
        
        # 更新图形
        self.update_plot()
        
        # 更新峰值显示
        self.update_peak_display()
        
        # 如果有多种检测器结果，显示额外信息
        if multi_results:
            info_msg = "Measurement completed with multiple detector modes:\n"
            for mode, (freq, amp) in multi_results.items():
                if freq is not None and amp is not None:
                    max_val = max(amp)
                    info_msg += f"{mode}: Max = {max_val:.2f} dBμV\n"
            messagebox.showinfo("Success", info_msg)
        else:
            messagebox.showinfo("Success", "Measurement completed!")
    
    def on_measurement_failed(self):
        """测量失败回调"""
        self.measure_btn.config(state=tk.NORMAL, text="Measure")
        messagebox.showerror("Error", "Measurement failed!")
    
    def on_measurement_error(self, error_msg):
        """测量错误回调"""
        self.measure_btn.config(state=tk.NORMAL, text="Measure")
        messagebox.showerror("Error", f"Error during measurement:\n{error_msg}")
    
    def update_plot(self):
        """更新图形显示 - 改进横轴标签"""
        if self.current_frequencies is None or self.current_amplitudes is None:
            return
        
        try:
            # 清除旧图形
            self.ax.clear()
            
            # 转换频率为MHz
            freq_mhz = [f/1e6 for f in self.current_frequencies]
            
            # 绘制测量数据
            self.ax.semilogx(freq_mhz, self.current_amplitudes, 
                        'b-', linewidth=1.2, label='Measured Spectrum', alpha=0.8)
            
            # 绘制FCC和CE限值
            fcc_limits = []
            ce_limits = []
            for freq in self.current_frequencies:
                fcc_limit, ce_limit = get_fcc_ce_limits(freq)
                fcc_limits.append(fcc_limit)
                ce_limits.append(ce_limit)
            
            self.ax.semilogx(freq_mhz, fcc_limits, 
                        'r--', linewidth=1.5, label='FCC Part 15 Class B', alpha=0.7)
            self.ax.semilogx(freq_mhz, ce_limits, 
                        'g--', linewidth=1.5, label='EN 55032 Class B', alpha=0.7)
            
            # 标记峰值
            if self.current_peaks:
                for peak in self.current_peaks:
                    freq_mhz_peak = peak['frequency_mhz']
                    amp_dbuv = peak['amplitude_dbuv']
                    exceed_fcc = peak['exceed_fcc']
                    exceed_ce = peak['exceed_ce']
                    color = 'red' if exceed_fcc or exceed_ce else 'orange'
                    self.ax.plot(freq_mhz_peak, amp_dbuv, 'o', color=color, 
                            markersize=6, markeredgecolor='black', markeredgewidth=0.5)
            
            # 设置标签和标题
            self.ax.set_xlabel('Frequency (MHz)', fontsize=10)
            self.ax.set_ylabel('Amplitude (dBμV)', fontsize=10)
            self.ax.set_title('EMC Spectrum Analysis', fontsize=12)
            self.ax.grid(True, which="both", alpha=0.3)
            self.ax.legend(fontsize=9, loc='upper right')
            
            # 改进横轴标签显示
            self.improve_xaxis_labels()
            
            # 设置坐标轴范围
            if freq_mhz:
                self.ax.set_xlim([min(freq_mhz), max(freq_mhz)])
            
            if self.current_amplitudes:
                all_y_values = list(self.current_amplitudes) + fcc_limits + ce_limits
                y_min = min(all_y_values) - 10
                y_max = max(all_y_values) + 10
                self.ax.set_ylim([y_min, y_max])
            
            # 更新画布
            self.fig.tight_layout()
            self.canvas.draw()
            
        except Exception as e:
            print(f"Plot error: {e}")
            self.ax.clear()
            self.ax.set_title("Error displaying plot")
            self.ax.grid(True, alpha=0.3)
            self.canvas.draw()
    
    def improve_xaxis_labels(self):
        """改进横轴标签显示"""
        # 获取当前频率范围
        if not hasattr(self, 'current_frequencies') or self.current_frequencies is None:
            return
            
        freq_mhz = [f/1e6 for f in self.current_frequencies]
        min_freq = min(freq_mhz)
        max_freq = max(freq_mhz)
        
        # 生成更智能的频率标签
        def generate_smart_ticks(min_f, max_f):
            ticks = []
            labels = []
            
            # 根据范围选择合适的步进
            if max_f <= 1:  # kHz范围 (优化9k-150k)
                if min_f >= 0.009 and max_f <= 0.150:  # 9k-150k范围
                    # 特别优化9k-150k范围
                    major_ticks = [0.009, 0.01, 0.03, 0.05, 0.1, 0.150]
                    major_labels = ['9k', '10k', '30k', '50k', '100k', '150k']
                    
                    minor_ticks = [0.015, 0.02, 0.025, 0.04, 0.06, 0.07, 0.08, 0.09, 0.12, 0.13]
                    minor_labels = ['15k', '20k', '25k', '40k', '60k', '70k', '80k', '90k', '120k', '130k']
                    
                    # 添加主要刻度
                    for tick, label in zip(major_ticks, major_labels):
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(label)
                    
                    # 添加中间刻度
                    for tick, label in zip(minor_ticks, minor_labels):
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(label)
                    
                    # 排序
                    combined = list(zip(ticks, labels))
                    combined.sort(key=lambda x: x[0])
                    ticks, labels = zip(*combined) if combined else ([], [])
                    ticks, labels = list(ticks), list(labels)
                else:
                    # 原来的逻辑
                    base_ticks = [0.01, 0.03, 0.1, 0.3, 1.0]
                    base_labels = ['10k', '30k', '100k', '300k', '1M']
                    for tick, label in zip(base_ticks, base_labels):
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(label)
            elif max_f <= 30:  # MHz范围 (优化150k-30M)
                if min_f >= 0.150 and max_f <= 30:  # 150k-30M范围
                    # 特别优化150k-30M范围
                    major_ticks = [0.150, 0.3, 0.5, 1.0, 2.0, 5.0, 10.0, 20.0, 30.0]
                    major_labels = ['150k', '300k', '500k', '1M', '2M', '5M', '10M', '20M', '30M']
                    
                    minor_ticks = [0.2, 0.4, 0.6, 0.7, 0.8, 0.9, 1.5, 2.5, 3.0, 4.0, 6.0, 7.0, 8.0, 9.0, 15.0, 25.0]
                    minor_labels = ['200k', '400k', '600k', '700k', '800k', '900k', '1.5M', '2.5M', '3M', '4M', '6M', '7M', '8M', '9M', '15M', '25M']
                    
                    # 添加主要刻度
                    for tick, label in zip(major_ticks, major_labels):
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(label)
                    
                    # 添加中间刻度
                    for tick, label in zip(minor_ticks, minor_labels):
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(label)
                    
                    # 排序
                    combined = list(zip(ticks, labels))
                    combined.sort(key=lambda x: x[0])
                    ticks, labels = zip(*combined) if combined else ([], [])
                    ticks, labels = list(ticks), list(labels)
                else:
                    # 原来的逻辑
                    base_ticks = [0.1, 0.3, 1.0, 3.0, 10.0, 30.0]
                    base_labels = ['100k', '300k', '1M', '3M', '10M', '30M']
                    for tick, label in zip(base_ticks, base_labels):
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(label)
            elif max_f <= 1000:  # MHz范围 (重点优化30M-1G)
                # 特别处理30M-1G范围
                if min_f >= 30 and max_f <= 1000:
                    # 主要刻度
                    major_ticks = [30, 50, 100, 200, 300, 500, 1000]
                    major_labels = ['30M', '50M', '100M', '200M', '300M', '500M', '1G']
                    
                    # 中间刻度
                    minor_ticks = [40, 60, 70, 80, 90, 150, 250, 400, 600, 800]
                    minor_labels = ['40M', '60M', '70M', '80M', '90M', '150M', '250M', '400M', '600M', '800M']
                    
                    # 添加主要刻度
                    for tick, label in zip(major_ticks, major_labels):
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(label)
                    
                    # 添加中间刻度
                    for tick, label in zip(minor_ticks, minor_labels):
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(label)
                    
                    # 排序
                    combined = list(zip(ticks, labels))
                    combined.sort(key=lambda x: x[0])
                    ticks, labels = zip(*combined) if combined else ([], [])
                    ticks, labels = list(ticks), list(labels)
                else:
                    # 其他MHz范围
                    base_ticks = [0.1, 0.3, 1.0, 3.0, 10.0, 30.0, 100.0, 300.0, 1000.0]
                    base_labels = ['100k', '300k', '1M', '3M', '10M', '30M', '100M', '300M', '1G']
                    for tick, label in zip(base_ticks, base_labels):
                        if min_f <= tick <= max_f:
                            ticks.append(tick)
                            labels.append(label)
            elif max_f <= 3000:  # GHz范围
                range_freq = max_f - min_f
                if range_freq <= 200:  # 小范围更密集
                    step = 50.0
                elif range_freq <= 500:
                    step = 100.0
                else:
                    step = 200.0
                
                current = max(step, min_f)
                while current <= max_f:
                    ticks.append(current)
                    if current >= 1000:
                        labels.append(f'{current/1000:.1f}G')
                    else:
                        labels.append(f'{current:.0f}M')
                    current += step
            else:  # 更高频率
                range_freq = max_f - min_f
                step = max(200.0, range_freq / 6)
                current = max(step, min_f)
                while current <= max_f:
                    ticks.append(current)
                    if current >= 1000:
                        labels.append(f'{current/1000:.1f}G')
                    else:
                        labels.append(f'{current:.0f}M')
                    current += step
            
            return ticks, labels
        
        # 生成标签
        ticks, labels = generate_smart_ticks(min_freq, max_freq)
        
        # 设置自定义标签
        if ticks:
            self.ax.set_xticks(ticks)
            self.ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=9)
        
        # 确保纵轴有数字标签
        self.ax.yaxis.set_major_locator(plt.MaxNLocator(nbins=10, integer=False))
        self.ax.tick_params(axis='y', which='major', labelsize=9)
        self.ax.tick_params(axis='x', which='major', labelsize=9)

    def update_peak_display(self):
        """更新峰值显示"""
        self.peak_text.delete(1.0, tk.END)
        
        if not self.current_peaks:
            self.peak_text.insert(tk.END, "No peaks detected\n")
            return
        
        # 添加表头（显示FCC和CE标准，添加单位标识，添加序号列）
        header = f"{'No':<4} {'Freq [MHz]':<12} {'Amplitude [dBμV]':<18} {'FCC Limit [dBμV]':<18} {'FCC Margin [dB]':<18} {'CE Limit [dBμV]':<18} {'CE Margin [dB]':<18} {'Status':<15}\n"
        separator = "-" * 130 + "\n"
        
        self.peak_text.insert(tk.END, header)
        self.peak_text.insert(tk.END, separator)
        
        # 按照要求排序：首先是超标的排前面，然后是余量最少的，最后是余量最多的
        exceed_peaks = [p for p in self.current_peaks if p['exceed_fcc']]
        normal_peaks = [p for p in self.current_peaks if not p['exceed_fcc']]
        
        # 对超标的峰值按余量排序（余量越少越前）
        exceed_peaks.sort(key=lambda x: x['fcc_margin'],reverse=True)
        
        # 对正常的峰值按余量排序（余量越少越前）
        normal_peaks.sort(key=lambda x: x['fcc_margin'],reverse=True)
        
        # 合并所有峰值，确保左下图中显示的所有peak都在表格中标出来
        all_peaks = exceed_peaks + normal_peaks
        
        # 显示所有峰值（添加序号）
        for i, peak in enumerate(all_peaks, 1):
            status = []
            if peak['exceed_fcc']:
                status.append("FCC Fail")
            if peak['exceed_ce']:
                status.append("CE Fail")
            if not status:
                status = ["Pass"]
            
            line = f"{i:<4} "
            line += f"{peak['frequency_mhz']:<12.3f} "
            line += f"{peak['amplitude_dbuv']:<18.2f} "
            line += f"{peak['fcc_limit']:<18.1f} "
            line += f"{peak['fcc_margin']:<18.2f} "
            line += f"{peak['ce_limit']:<18.1f} "
            line += f"{peak['ce_margin']:<18.2f} "
            line += f"{', '.join(status):<15}\n"
            
            self.peak_text.insert(tk.END, line)
    
    def save_data(self):
        """保存数据 - 增强版本"""
        if not self.emi_results and not self.current_frequencies:
            messagebox.showwarning("Warning", "No data to save!")
            return
        
        try:

            saved_files = []
            
            # 如果有EMI测量数据，保存完整数据
            if self.emi_results:
                # 显示采样数据统计
                if "sampling_data" in self.emi_results:
                    sampling_data = self.emi_results["sampling_data"]
                    if sampling_data:
                        info_msg = f"准备保存数据:\n"
                        info_msg += f"采样次数: {len(sampling_data)}\n"
                        if sampling_data:
                            info_msg += f"每次采样点数: {len(sampling_data[0]['amplitudes'])}\n"
                            # 检查数据是否不同
                            if len(sampling_data) > 1:
                                first_sample = sampling_data[0]['amplitudes'][:10]  # 前10个点
                                last_sample = sampling_data[-1]['amplitudes'][:10]  # 前10个点
                                if first_sample != last_sample:
                                    info_msg += "✅ 检测到采样数据有变化\n"
                                else:
                                    info_msg += "⚠️  注意: 前几次采样数据相同\n"
                        messagebox.showinfo("数据检查", info_msg)
                
                files = save_emi_measurement_data(self.emi_results)
                saved_files.extend(files)
                
                # 也保存峰值分析
                if self.current_peaks:
                    peak_file = save_peak_analysis(self.current_peaks)
                    saved_files.append(peak_file)
            else:
                # 保存单次测量数据
                if self.current_frequencies and self.current_amplitudes:
                    spectrum_file = save_spectrum_data(self.current_frequencies, self.current_amplitudes)
                    saved_files.append(spectrum_file)
                    
                    if self.current_peaks:
                        peak_file = save_peak_analysis(self.current_peaks)
                        saved_files.append(peak_file)
            
            # 显示保存结果
            if saved_files:
                file_list = "\n".join([f"• {os.path.basename(f)}" for f in saved_files])
                full_paths = "\n".join(saved_files)
                messagebox.showinfo("Success", f"数据保存成功!\n\n保存的文件:\n{file_list}\n\n完整路径:\n{full_paths}")
            else:
                messagebox.showinfo("Info", "没有数据被保存。")
                
        except Exception as e:
            messagebox.showerror("Error", f"保存数据时出错:\n{e}")
            import traceback
            traceback.print_exc()

    def perform_ai_analysis(self):
        """执行AI分析"""
        try:
            # 检查是否有数据可以分析
            if not self.emi_results and not self.current_peaks:
                messagebox.showwarning("Warning", "No data available for AI analysis!")
                return
            
            # 禁用AI分析按钮，防止重复点击
            self.ai_analysis_btn.config(state=tk.DISABLED)
            self.progress_var.set("AI Analysis in progress...")
            
            # 在新线程中执行AI分析
            def ai_analysis_task():
                try:
                    # 导入chat.py中的ChatBot
                    from chat import ChatBot, sys_prompt
                    
                    # 创建ChatBot实例
                    bot = ChatBot(
                        api_key="b8800336-579b-4322-b2e9-ca0f4443db71",
                        base_url="https://ark.cn-beijing.volces.com/api/v3",
                        model="ep-20250708144105-dqzdw",
                        system_message=sys_prompt
                    )
                    
                    # 准备输入数据
                    # 获取频率范围
                    start_freq = 0
                    stop_freq = 0
                    if self.controller and self.controller.current_config:
                        config = self.controller.get_preset_configs().get(self.selected_preset_key, {})
                        if config:
                            start_freq = config.get("start_freq", 0)
                            stop_freq = config.get("stop_freq", 0)
                    
                    # 获取测量时长
                    duration = 0
                    if "measurement_summary" in self.emi_results:
                        summary = self.emi_results["measurement_summary"]
                        duration = summary.get("actual_measurement_time", 0)
                    
                    # 构建输入文本
                    input_text = f"频段:{start_freq/1e6:.3f}MHz-{stop_freq/1e6:.3f}MHz 测量时长：{duration}s 测量数据：\n"
                    
                    # 直接使用右下角表格中已经计算和排序好的数据
                    # 获取表格中的文本内容
                    table_content = self.peak_text.get("1.0", tk.END)
                    
                    # 提取表格内容部分（去掉开头的标题行）
                    lines = table_content.strip().split('\n')
                    if len(lines) > 3:  # 确保有足够的行
                        # 添加QUASI_PEAK模式标题
                        input_text += "QUASI_PEAK Mode Results:\n"
                        
                        # 添加分隔线和表头
                        input_text += "="*100 + "\n"
                        
                        # 找到表头行的索引
                        header_index = -1
                        separator_index = -1
                        for i, line in enumerate(lines):
                            if "No" in line and "Freq [MHz]" in line and "Amplitude [dBμV]" in line:
                                header_index = i
                                break
                        
                        # 如果找到了表头，提取表头和数据行
                        if header_index != -1:
                            # 添加表头
                            input_text += lines[header_index] + "\n"
                            
                            # 找到分隔线
                            for i in range(header_index + 1, len(lines)):
                                if "-" in lines[i] and len(lines[i]) > 50:  # 分隔线通常很长且包含很多"-"
                                    separator_index = i
                                    input_text += lines[i] + "\n"
                                    break
                            
                            # 添加数据行（从分隔线之后开始）
                            if separator_index != -1:
                                for i in range(separator_index + 1, len(lines)):
                                    # 检查是否是有效的数据行（不是空行且包含数字）
                                    if lines[i].strip() and any(c.isdigit() for c in lines[i]):
                                        input_text += lines[i] + "\n"
                    
                    # 调用AI分析
                    response = bot.chat_no_stream(input_text)
                    msg_obj = response.choices[0].message
                    ai_result = msg_obj.content if hasattr(msg_obj, "content") else msg_obj.get("content", "")
                    
                    # 在主线程中显示结果
                    self.root.after(0, lambda: self.show_ai_analysis_result(ai_result))
                except Exception as e:
                    self.root.after(0, lambda: self.on_ai_analysis_error(str(e)))
                finally:
                    # 重新启用AI分析按钮
                    self.root.after(0, lambda: self.ai_analysis_btn.config(state=tk.NORMAL))
                    self.root.after(0, lambda: self.progress_var.set("Ready"))
            
            # 启动AI分析线程
            threading.Thread(target=ai_analysis_task, daemon=True).start()
        except Exception as e:
            messagebox.showerror("Error", f"启动AI分析时出错:\n{e}")
            self.ai_analysis_btn.config(state=tk.NORMAL)
            self.progress_var.set("Ready")

    def show_ai_analysis_result(self, result):
        """显示AI分析结果"""
        # 保存AI分析结果用于PDF导出
        self._last_ai_result = result
        
        # 创建新窗口显示结果
        result_window = tk.Toplevel(self.root)
        result_window.title("AI Analysis Result")
        result_window.geometry("800x600")
        result_window.minsize(600, 400)
        
        # 创建文本框和滚动条
        text_frame = ttk.Frame(result_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        text_widget = tk.Text(text_frame, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 插入AI分析结果
        text_widget.insert(tk.END, result)
        text_widget.config(state=tk.DISABLED)  # 设置为只读
        
        # 添加关闭按钮
        close_btn = ttk.Button(result_window, text="Close", command=result_window.destroy)
        close_btn.pack(pady=10)

    def on_ai_analysis_error(self, error_msg):
        """AI分析错误回调"""
        messagebox.showerror("AI Analysis Error", f"AI分析过程中出错:\n{error_msg}")

    def perform_ai_analysis_for_pdf(self):
        """为PDF导出执行AI分析"""
        try:
            # 检查是否有数据可以分析
            if not self.emi_results and not self.current_peaks:
                messagebox.showwarning("Warning", "No data available for AI analysis!")
                return
            
            # 禁用导出PDF按钮，防止重复点击
            self.export_pdf_btn.config(state=tk.DISABLED)
            self.progress_var.set("AI Analysis for PDF in progress...")
            
            # 在新线程中执行AI分析
            def ai_analysis_task():
                try:
                    # 导入chat.py中的ChatBot
                    from chat import ChatBot, sys_prompt
                    
                    # 创建ChatBot实例
                    bot = ChatBot(
                        api_key="b8800336-579b-4322-b2e9-ca0f4443db71",
                        base_url="https://ark.cn-beijing.volces.com/api/v3",
                        model="ep-20250708144105-dqzdw",
                        system_message=sys_prompt
                    )
                    
                    # 准备输入数据
                    # 获取频率范围
                    start_freq = 0
                    stop_freq = 0
                    if self.controller and self.controller.current_config:
                        config = self.controller.get_preset_configs().get(self.selected_preset_key, {})
                        if config:
                            start_freq = config.get("start_freq", 0)
                            stop_freq = config.get("stop_freq", 0)
                    
                    # 获取测量时长
                    duration = 0
                    if "measurement_summary" in self.emi_results:
                        summary = self.emi_results["measurement_summary"]
                        duration = summary.get("actual_measurement_time", 0)
                    
                    # 构建输入文本
                    input_text = f"频段:{start_freq/1e6:.3f}MHz-{stop_freq/1e6:.3f}MHz 测量时长：{duration}s 测量数据：\n"
                    
                    # 直接使用右下角表格中已经计算和排序好的数据
                    # 获取表格中的文本内容
                    table_content = self.peak_text.get("1.0", tk.END)
                    
                    # 提取表格内容部分（去掉开头的标题行）
                    lines = table_content.strip().split('\n')
                    if len(lines) > 3:  # 确保有足够的行
                        # 添加QUASI_PEAK模式标题
                        input_text += "QUASI_PEAK Mode Results:\n"
                        
                        # 添加分隔线和表头
                        input_text += "="*100 + "\n"
                        
                        # 找到表头行的索引
                        header_index = -1
                        separator_index = -1
                        for i, line in enumerate(lines):
                            if "No" in line and "Freq [MHz]" in line and "Amplitude [dBμV]" in line:
                                header_index = i
                                break
                        
                        # 如果找到了表头，提取表头和数据行
                        if header_index != -1:
                            # 添加表头
                            input_text += lines[header_index] + "\n"
                            
                            # 找到分隔线
                            for i in range(header_index + 1, len(lines)):
                                if "-" in lines[i] and len(lines[i]) > 50:  # 分隔线通常很长且包含很多"-"
                                    separator_index = i
                                    input_text += lines[i] + "\n"
                                    break
                            
                            # 添加数据行（从分隔线之后开始）
                            if separator_index != -1:
                                for i in range(separator_index + 1, len(lines)):
                                    # 检查是否是有效的数据行（不是空行且包含数字）
                                    if lines[i].strip() and any(c.isdigit() for c in lines[i]):
                                        input_text += lines[i] + "\n"
                    
                    # 调用AI分析
                    response = bot.chat_no_stream(input_text)
                    msg_obj = response.choices[0].message
                    ai_result = msg_obj.content if hasattr(msg_obj, "content") else msg_obj.get("content", "")
                    
                    # 保存AI分析结果用于PDF导出
                    self._last_ai_result = ai_result
                    
                    # 在主线程中继续执行PDF导出
                    self.root.after(0, self.export_pdf)
                except Exception as e:
                    self.root.after(0, lambda: self.on_ai_analysis_error(str(e)))
                finally:
                    # 重新启用导出PDF按钮
                    self.root.after(0, lambda: self.export_pdf_btn.config(state=tk.NORMAL))
                    self.root.after(0, lambda: self.progress_var.set("Ready"))
            
            # 启动AI分析线程
            threading.Thread(target=ai_analysis_task, daemon=True).start()
        except Exception as e:
            messagebox.showerror("Error", f"启动AI分析时出错:\n{e}")
            self.export_pdf_btn.config(state=tk.NORMAL)
            self.progress_var.set("Ready")

    def export_pdf(self):
        """导出PDF报告"""
        try:
            # 检查是否有数据可以导出
            if not self.emi_results and not self.current_peaks:
                messagebox.showwarning("Warning", "No data available for PDF export!")
                return
            
            # 自动执行AI分析（如果还没有结果）
            if not hasattr(self, '_last_ai_result') or not self._last_ai_result:
                self.perform_ai_analysis_for_pdf()
                return  # AI分析完成后会重新调用export_pdf
            
            # 保存波形图为PNG文件
            graph_filename = "temp_spectrum_graph.png"
            self.fig.savefig(graph_filename, dpi=150, bbox_inches='tight')
            
            # 获取用户输入的信息
            customer = self.customer_var.get()
            eut = self.eut_var.get()
            model = self.model_var.get()
            engineer = self.engineer_var.get()
            remark = self.remark_var.get()
            
            # 获取频率范围
            start_freq = 0
            stop_freq = 0
            if self.controller and self.controller.current_config:
                config = self.controller.get_preset_configs().get(self.selected_preset_key, {})
                if config:
                    start_freq = config.get("start_freq", 0)
                    stop_freq = config.get("stop_freq", 0)
            
            # 获取测量时长
            duration = 0
            if "measurement_summary" in self.emi_results:
                summary = self.emi_results["measurement_summary"]
                duration = summary.get("actual_measurement_time", 0)
            
            # 构建mode信息
            mode = f"{start_freq/1e6:.3f}MHz-{stop_freq/1e6:.3f}MHz_{duration}s"
            
            # 获取表格数据（只取前15个点）
            table_content = self.peak_text.get("1.0", tk.END)
            print(table_content)
 
            # 获取AI分析结果
            summary_text = getattr(self, '_last_ai_result', '')
            
            # 生成PDF文件名
            filename = f"{eut}-{mode}.pdf"
            
            # 导入PDF生成模块
            from utils.create_pdf import generate_test_report
            
            # 项目信息
            project_info = {
                'customer': customer,
                'eut': eut,
                'model': model,
                'mode': mode,
                'engineer': engineer,
                'remark': remark
            }
            
            # 生成PDF报告
            generate_test_report(
                filename=filename,
                logo_path="./assets/m5logo2022.png",
                project_info=project_info,
                test_graph_path=graph_filename,
                spectrum_data=table_content,
                summary_text=summary_text
            )
            
            messagebox.showinfo("Success", f"PDF报告已生成: {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"导出PDF时出错:\n{e}")
            import traceback
            traceback.print_exc()

def main():
    root = tk.Tk()
    app = EMCAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
