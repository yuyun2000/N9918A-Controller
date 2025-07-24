
# n9918a_frontend.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
import matplotlib
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
        self.sweep_time = 1.0  # 默认扫描时间
        
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
        
        self.measure_btn = ttk.Button(measure_frame, text="Measure", command=self.start_measurement, state=tk.DISABLED)
        self.measure_btn.pack(side=tk.LEFT, padx=5)
        
        self.save_btn = ttk.Button(measure_frame, text="Save Data", command=self.save_data, state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT, padx=5)
        
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
    
    def on_measurement_complete(self, frequencies, amplitudes, peaks):
        """测量完成回调"""
        self.measure_btn.config(state=tk.NORMAL, text="Measure")
        self.save_btn.config(state=tk.NORMAL)
        
        # 保存数据
        self.current_frequencies = frequencies
        self.current_amplitudes = amplitudes
        self.current_peaks = peaks
        
        # 更新图形
        self.update_plot()
        
        # 更新峰值显示
        self.update_peak_display()
        
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
        """更新图形显示 - 修复FCC/CE标准线显示"""
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
            
            # 绘制FCC和CE限值 - 确保在整个频率范围内都计算
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
        """保存数据"""
        if self.current_frequencies is None or self.current_amplitudes is None:
            messagebox.showwarning("Warning", "No data to save!")
            return
        
        try:
            # 保存频谱数据
            from n9918a_backend import save_spectrum_data, save_peak_analysis
            
            spectrum_file = save_spectrum_data(self.current_frequencies, self.current_amplitudes)
            if self.current_peaks:
                peak_file = save_peak_analysis(self.current_peaks)
                messagebox.showinfo("Success", f"Data saved:\n{spectrum_file}\n{peak_file}")
            else:
                messagebox.showinfo("Success", f"Spectrum data saved:\n{spectrum_file}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error saving data:\n{e}")

def main():
    root = tk.Tk()
    app = EMCAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()