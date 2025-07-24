import pyvisa
# import matplotlib.pyplot as plt
import time
import csv
import os
from datetime import datetime

class N9918AController:
    """
    N9918A FieldFox Network Analyzer Controller
    Provides methods to connect, configure, and read measurement data from N9918A
    """
    
    def __init__(self, ip_address='192.168.0.124', timeout=10000):
        """
        Initialize N9918A controller
        
        Args:
            ip_address (str): IP address of the N9918A device
            timeout (int): Connection timeout in milliseconds
        """
        self.ip_address = ip_address
        self.timeout = timeout
        self.rm = None
        self.device = None
        self.connected = False
        
        # Measurement parameters
        self.start_freq = None
        self.stop_freq = None
        self.n_points = None
        
    def connect(self):
        """
        Connect to N9918A device
        
        Returns:
            bool: True if connection successful, False otherwise
        """
        try:
            self.rm = pyvisa.ResourceManager()
            self.device = self.rm.open_resource(f"TCPIP0::{self.ip_address}::inst0::INSTR")
            self.device.timeout = self.timeout
            
            # Clear status and query device identification
            self.device.write("*CLS")
            device_id = self.device.query("*IDN?")
            print(f"Connected to: {device_id}")
            
            # Set to Spectrum Analyzer mode
            self.device.write("INST:SEL 'SA'")
            self.device.write("*OPC?")
            self.device.read()
            
            self.connected = True
            print("Successfully connected to N9918A")
            return True
            
        except Exception as e:
            print(f"ERROR: Unable to connect to device - {e}")
            self.connected = False
            return False
    
    def disconnect(self):
        """Disconnect from the device"""
        if self.device:
            self.device.close()
        if self.rm:
            self.rm.close()
        self.connected = False
        print("Disconnected from N9918A")
    
    def configure_sa_measurement(self, start_freq, stop_freq, n_points):
        """
        Configure spectrum analyzer measurement parameters
        
        Args:
            start_freq (float): Start frequency in Hz
            stop_freq (float): Stop frequency in Hz
            n_points (int): Number of measurement points
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return False
            
        try:
            # Set frequency range
            self.device.write(f"SENS:FREQ:STAR {start_freq}")
            self.device.write(f"SENS:FREQ:STOP {stop_freq}")
            
            # Set number of points
            self.device.write(f"SENS:SWE:POIN {n_points}")
            
            # Store parameters
            self.start_freq = start_freq
            self.stop_freq = stop_freq
            self.n_points = n_points
            
            print(f"Configured SA: {start_freq/1e9:.6f} - {stop_freq/1e9:.6f} GHz, {n_points} points")
            return True
            
        except Exception as e:
            print(f"ERROR: Failed to configure measurement - {e}")
            return False
    
    def read_trace_data(self):
        """
        Read trace data from the device
        
        Returns:
            tuple: (frequencies, amplitudes) or (None, None) if error
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return None, None
            
        try:
            # Read trace data
            self.device.write("TRACE:DATA?")
            trace_data = self.device.read()
            amplitudes = [float(x) for x in trace_data.split(",")]
            
            # Calculate frequency array
            freq_step = (self.stop_freq - self.start_freq) / (self.n_points - 1)
            frequencies = [self.start_freq + i * freq_step for i in range(self.n_points)]
            
            return frequencies, amplitudes
            
        except Exception as e:
            print(f"ERROR: Failed to read trace data - {e}")
            return None, None
    
    def plot_measurement(self, frequencies, amplitudes, title="N9918A Measurement"):
        """
        Plot measurement data
        
        Args:
            frequencies (list): Frequency array
            amplitudes (list): Amplitude array
            title (str): Plot title
        """
        if frequencies is None or amplitudes is None:
            print("ERROR: No data to plot")
            return
            
        plt.figure(figsize=(10, 6))
        plt.plot(frequencies, amplitudes)
        
        # Determine frequency unit
        if max(frequencies) > 1e9:
            freq_unit = "GHz"
            freq_data = [f / 1e9 for f in frequencies]
        elif max(frequencies) > 1e6:
            freq_unit = "MHz"
            freq_data = [f / 1e6 for f in frequencies]
        elif max(frequencies) > 1e3:
            freq_unit = "kHz"
            freq_data = [f / 1e3 for f in frequencies]
        else:
            freq_unit = "Hz"
            freq_data = frequencies
            
        plt.plot(freq_data, amplitudes)
        plt.xlabel(f'Frequency [{freq_unit}]')
        plt.ylabel('Level [dBm]')
        plt.title(title)
        plt.grid(True)
        plt.tight_layout()
        plt.show()
    
    def save_measurement(self, frequencies, amplitudes, filename, site_name="", n_samples=1, interval=0.2):
        """
        Save measurement data to CSV file
        
        Args:
            frequencies (list): Frequency array
            amplitudes (list): Amplitude array
            filename (str): Output filename
            site_name (str): Site name for metadata
            n_samples (int): Number of samples taken
            interval (float): Time interval between samples
        """
        if frequencies is None or amplitudes is None:
            print("ERROR: No data to save")
            return
            
        # Create measurement folder if it doesn't exist
        measurement_folder = 'measurement_data'
        if not os.path.exists(measurement_folder):
            os.makedirs(measurement_folder)
            
        filepath = os.path.join(measurement_folder, filename)
        
        with open(filepath, 'w', newline='') as f:
            writer = csv.writer(f)
            
            # Write metadata
            writer.writerow(['Create Date', datetime.now().strftime("%Y%m%d_%H%M%S")])
            writer.writerow(['Site', site_name])
            writer.writerow(['Start Freq (Hz)', str(self.start_freq)])
            writer.writerow(['Stop Freq (Hz)', str(self.stop_freq)])
            writer.writerow(['# Points', str(self.n_points)])
            writer.writerow(['# Samples', str(n_samples)])
            writer.writerow(['Interval (s)', str(interval)])
            
            # Write data header
            writer.writerow(['Frequency (Hz)', 'Amplitude (dBm)'])
            
            # Write measurement data
            for freq, amp in zip(frequencies, amplitudes):
                writer.writerow([freq, amp])
                
        print(f"Measurement data saved to: {filepath}")
    
    def continuous_measurement(self, site_name, n_samples=50, interval=0.2):
        """
        Perform continuous measurement with multiple samples
        
        Args:
            site_name (str): Site name for the measurement
            n_samples (int): Number of samples to take
            interval (float): Time interval between samples in seconds
        """
        if not self.connected:
            print("ERROR: Device not connected")
            return
            
        filename = f"{site_name}.csv"
        measurement_folder = 'measurement_data'
        filepath = os.path.join(measurement_folder, filename)
        
        if os.path.exists(filepath):
            print(f"ERROR: File {filepath} already exists")
            return
            
        print(f"Starting continuous measurement: {n_samples} samples, {interval}s interval")
        
        # Create file and write headers
        if not os.path.exists(measurement_folder):
            os.makedirs(measurement_folder)
            
        with open(filepath, 'w', newline='') as f:
            writer = csv.writer(f)
            
            # Write metadata
            measurement_time = time.time()
            writer.writerow(['Create Date', datetime.now().strftime("%Y%m%d_%H%M%S")])
            writer.writerow(['Site', site_name])
            writer.writerow(['Start Freq (Hz)', str(self.start_freq)])
            writer.writerow(['Stop Freq (Hz)', str(self.stop_freq)])
            writer.writerow(['# Points', str(self.n_points)])
            writer.writerow(['# Samples', str(n_samples)])
            writer.writerow(['Interval (s)', str(interval)])
            writer.writerow(['Start Time', str(measurement_time)])
            
            # Write data header
            writer.writerow(['Sample', 'Time (s)', 'Max Level (dBm)'] + [f'Point_{i}' for i in range(self.n_points)])
            
            # Perform measurements
            for sample in range(n_samples):
                start_time = time.time()
                
                # Read trace data
                frequencies, amplitudes = self.read_trace_data()
                if frequencies is None:
                    print(f"ERROR: Failed to read data for sample {sample + 1}")
                    continue
                    
                max_level = max(amplitudes)
                elapsed_time = start_time - measurement_time
                
                # Write data row
                row = [sample + 1, elapsed_time, max_level] + amplitudes
                writer.writerow(row)
                
                print(f"Sample {sample + 1}/{n_samples}, Max: {max_level:.2f} dBm")
                
                # Plot latest measurement
                if sample == n_samples - 1:  # Plot only the last measurement
                    self.plot_measurement(frequencies, amplitudes, f"{site_name} - Final Measurement")
                
                # Wait for next sample
                elapsed = time.time() - start_time
                sleep_time = interval - elapsed
                if sleep_time > 0:
                    time.sleep(sleep_time)
                    
        print(f"Continuous measurement completed. Data saved to: {filepath}")