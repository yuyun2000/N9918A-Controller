
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

class EMCAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("N9918A EMC Analyzer")
        self.root.geometry("1400x900")
        self.root.minsize(1200, 800)
        
        # 创建后端控制器
        self.controller = N9918AController(ip_address='192.168.20.39')
        
        # 当前数据
        self.current_frequencies = None
        self.current_amplitudes = None
        self.current_peaks = None
        self.emi_results = {}  # 存储完整的EMI测量结果（包含采样数据）
        self.sweep_time = 1.0
        self.measurement_in_progress = False
            
        # 创建界面
        self.create_widgets()
        
        # 初始化状态
        self.update_status()
    
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
        self.ip_var = tk.StringVar(value="192.168.20.39")
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
        
        # 创建可调整的主显示区域
        self.create_main_display(main_frame)
        
        # 添加调整比例的滑块
        self.create_ratio_control(main_frame)


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
            if max_f <= 1:  # kHz范围
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
            elif max_f <= 10:  # MHz范围（更密集）
                # 生成更密集的MHz标签
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
                current = max(1.0, min_f)
                while current <= max_f:
                    ticks.append(current)
                    labels.append(f'{current:.0f}M')
                    current += max(1.0, (max_f - min_f) / 10)
            elif max_f <= 1000:  # GHz范围
                current = max(10.0, min_f)
                while current <= max_f:
                    ticks.append(current)
                    if current >= 1000:
                        labels.append(f'{current/1000:.1f}G')
                    else:
                        labels.append(f'{current:.0f}M')
                    current += max(50.0, (max_f - min_f) / 8)
            else:  # 更高频率
                current = max(100.0, min_f)
                while current <= max_f:
                    ticks.append(current)
                    if current >= 1000:
                        labels.append(f'{current/1000:.1f}G')
                    else:
                        labels.append(f'{current:.0f}M')
                    current += (max_f - min_f) / 6
            
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
        
        # 显示每种检测器模式的结果
        for mode, data in self.emi_results.items():
            if mode in ["measurement_summary", "sampling_info"]:
                continue
                
            if isinstance(data, tuple) and len(data) >= 2:
                frequencies, amplitudes = data[0], data[1]
                if amplitudes is None:
                    continue
                    
                self.peak_text.insert(tk.END, f"\n{mode} Mode Results:\n")
                self.peak_text.insert(tk.END, "="*80 + "\n")
                header = f"{'Freq (MHz)':<12} {'Amplitude':<12} {'FCC Limit':<12} {'CE Limit':<12} {'FCC Margin':<12} {'CE Margin':<12} {'Status':<15}\n"
                separator = "-"*80 + "\n"
                
                self.peak_text.insert(tk.END, header)
                self.peak_text.insert(tk.END, separator)
                
                # 分析峰值
                peaks = post_process_peak_search(frequencies, amplitudes)
                for peak in peaks[:5]:  # 只显示前5个峰值
                    status = []
                    if peak['exceed_fcc']:
                        status.append("FCC Fail")
                    if peak['exceed_ce']:
                        status.append("CE Fail")
                    if not status:
                        status = ["Pass"]
                    
                    line = f"{peak['frequency_mhz']:<12.3f} "
                    line += f"{peak['amplitude_dbuv']:<12.2f} "
                    line += f"{peak['fcc_limit']:<12.1f} "
                    line += f"{peak['ce_limit']:<12.1f} "
                    line += f"{peak['fcc_margin']:<12.2f} "
                    line += f"{peak['ce_margin']:<12.2f} "
                    line += f"{', '.join(status):<15}\n"
                    
                    self.peak_text.insert(tk.END, line)
    def on_emi_measurement_complete(self, results):
        """EMI测量完成回调 - 快速版本"""
        self.emi_results = results
        self.progress_var.set("Fast EMI Measurement completed")
        self.save_btn.config(state=tk.NORMAL)
        
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

    def stop_measurement(self):
        """停止测量"""
        try:
            if self.controller.connected and self.controller.device:
                self.controller.device.write("INIT:CONT OFF")
                self.progress_var.set("Measurement stopped")
                self.set_measurement_buttons_state(True)
        except Exception as e:
            print(f"Error stopping measurement: {e}")
    
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
        messagebox.showinfo("Success", "Device configured successfully!")
    
    def on_config_failed(self):
        """配置失败回调"""
        self.config_btn.config(state=tk.NORMAL, text="Configure")
        messagebox.showerror("Error", "Failed to configure device!")
    
    def on_config_error(self, error_msg):
        """配置错误回调"""
        self.config_btn.config(state=tk.NORMAL, text="Configure")
        messagebox.showerror("Error", f"Error configuring device:\n{error_msg}")
    
    def start_measurement(self):
        """开始测量"""
        def measure_task():
            try:
                self.measure_btn.config(state=tk.DISABLED, text="Measuring...")
                self.root.update()
                
                # 获取扫描时间
                self.sweep_time = self.get_sweep_time()
                print(f"Estimated sweep time: {self.sweep_time:.1f} seconds")
                
                frequencies, amplitudes = self.controller.read_trace_data()
                if frequencies is not None and amplitudes is not None:
                    # 峰值分析
                    peaks = post_process_peak_search(frequencies, amplitudes)
                    self.root.after(0, lambda: self.on_measurement_complete(frequencies, amplitudes, peaks))
                else:
                    self.root.after(0, self.on_measurement_failed)
            except Exception as e:
                self.root.after(0, lambda: self.on_measurement_error(str(e)))
        
        threading.Thread(target=measure_task, daemon=True).start()
    
    def on_measurement_complete(self, frequencies, amplitudes, peaks, multi_results=None):
        """测量完成回调 - 增强版本"""
        self.set_measurement_buttons_state(True)
        self.progress_var.set("Measurement completed")
        self.save_btn.config(state=tk.NORMAL)
        
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
        
        # 根据频率范围设置合适的标签
        if max_freq <= 150:  # kHz范围
            ticks = [0.01, 0.03, 0.1, 0.3, 1.0, 3.0, 10.0, 30.0, 100.0, 150.0]
            tick_labels = ['10k', '30k', '100k', '300k', '1M', '3M', '10M', '30M', '100M', '150M']
        elif max_freq <= 1000:  # MHz范围
            ticks = [0.1, 0.3, 1.0, 3.0, 10.0, 30.0, 100.0, 300.0, 1000.0]
            tick_labels = ['100k', '300k', '1M', '3M', '10M', '30M', '100M', '300M', '1G']
        elif max_freq <= 3000:  # GHz范围
            ticks = [1.0, 3.0, 10.0, 30.0, 100.0, 300.0, 1000.0, 3000.0]
            tick_labels = ['1M', '3M', '10M', '30M', '100M', '300M', '1G', '3G']
        else:  # 更高频率
            ticks = [10.0, 30.0, 100.0, 300.0, 1000.0, 3000.0, 10000.0]
            tick_labels = ['10M', '30M', '100M', '300M', '1G', '3G', '10G']
        
        # 过滤出在当前范围内的标签
        valid_ticks = []
        valid_labels = []
        for tick, label in zip(ticks, tick_labels):
            if min_freq <= tick <= max_freq:
                valid_ticks.append(tick)
                valid_labels.append(label)
        
        # 设置自定义标签
        if valid_ticks:
            self.ax.set_xticks(valid_ticks)
            self.ax.set_xticklabels(valid_labels, rotation=45, ha='right')
        
        # 确保纵轴有数字标签
        self.ax.yaxis.set_major_locator(plt.MaxNLocator(integer=True))
        self.ax.tick_params(axis='y', which='major', labelsize=9)
        self.ax.tick_params(axis='x', which='major', labelsize=9)

    def update_peak_display(self):
        """更新峰值显示"""
        self.peak_text.delete(1.0, tk.END)
        
        if not self.current_peaks:
            self.peak_text.insert(tk.END, "No peaks detected\n")
            return
        
        # 添加表头
        header = f"{'Freq (MHz)':<12} {'Amp (dBμV)':<12} {'FCC Limit':<12} {'CE Limit':<12} {'FCC Margin':<12} {'CE Margin':<12} {'Status':<15}\n"
        separator = "-" * 95 + "\n"
        
        self.peak_text.insert(tk.END, header)
        self.peak_text.insert(tk.END, separator)
        
        for peak in self.current_peaks:
            status = []
            if peak['exceed_fcc']:
                status.append("FCC Fail")
            if peak['exceed_ce']:
                status.append("CE Fail")
            if not status:
                status = ["Pass"]
            
            line = f"{peak['frequency_mhz']:<12.3f} "
            line += f"{peak['amplitude_dbuv']:<12.2f} "
            line += f"{peak['fcc_limit']:<12.1f} "
            line += f"{peak['ce_limit']:<12.1f} "
            line += f"{peak['fcc_margin']:<12.2f} "
            line += f"{peak['ce_margin']:<12.2f} "
            line += f"{', '.join(status):<15}\n"
            
            self.peak_text.insert(tk.END, line)
    
    def save_data(self):
        """保存数据 - 增强版本"""
        if not self.emi_results and not self.current_frequencies:
            messagebox.showwarning("Warning", "No data to save!")
            return
        
        try:
            from n9918a_backend import save_emi_measurement_data,save_peak_analysis,save_spectrum_data  # 确保导入正确
            import os
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

def main():
    root = tk.Tk()
    app = EMCAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()