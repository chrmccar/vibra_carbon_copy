import sys

# Check for required modules
required = {
    'ttkbootstrap': 'ttkbootstrap',
    'PIL': 'pillow',
    'serial': 'pyserial',
    'tkinter': None  
}

missing = []
for module, pipname in required.items():
    try:
        __import__(module)
    except ImportError:
        missing.append((module, pipname))

if missing:
    print("\nMissing required packages or modules:")
    for module, pipname in missing:
        if module == "tkinter":
            print(
                "\n* tkinter is required for this program to run.\n"
                "On Windows and Mac, it is included with Python installers.\n"
                "On Linux, you may need to install it separately:\n"
                "    sudo apt-get install python3-tk\n"
            )
        elif pipname:
            print(f"  - {module} (install with: pip install {pipname})")
    print("\nRecommended: Python 3.8 or newer.")
    sys.exit(1)

if sys.version_info < (3, 7):
    print("Python 3.7 or newer required (3.8+ recommended).")
    sys.exit(1)

import serial
import time
import csv
import os
import re
import threading
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import serial.tools.list_ports
import tkinter.messagebox
import tkinter
from tkinter import simpledialog

def open_serial_connection(port_name):
    ser = serial.Serial(port_name, 9600, timeout=1)
    ser.reset_input_buffer()
    ser.reset_output_buffer()
    ser.flush()
    return ser

class BalanceLoggerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Balance Logger")

        main_frame = ttk.Frame(root, padding=10)
        main_frame.pack(fill=BOTH, expand=True)

        self.data_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
        try:
            os.makedirs(self.data_folder, exist_ok=True)
        except Exception as e:
            tkinter.messagebox.showerror("Error", f"Could not create data folder:\n{self.data_folder}\n\n{e}")
            self.root.destroy()
            return

        # --- Top Buttons (centered)
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(10, 20))

        self.add_balance_button = ttk.Button(button_frame, text="Add Balance", command=self.add_balance, bootstyle="primary")
        self.add_balance_button.pack(side=LEFT, padx=5)

        self.remove_balance_button = ttk.Button(button_frame, text="Remove Balance", command=self.remove_balance, bootstyle="primary")
        self.remove_balance_button.pack(side=LEFT, padx=5)

        self.start_button = ttk.Button(button_frame, text="Start Recording", command=self.start_measurements, bootstyle="primary")
        self.start_button.pack(side=LEFT, padx=5)

        self.stop_button = ttk.Button(button_frame, text="Stop Recording and Exit", command=self.stop_measurements, bootstyle="danger")
        self.stop_button.pack(side=LEFT, padx=5)

        # --- Settings Frame
        settings_labelframe = ttk.Labelframe(main_frame, text="Settings", padding=10)
        settings_labelframe.pack(fill=X, pady=(0, 20))

        self.freq_var = ttk.IntVar(value=10)
        self.filename_var = ttk.StringVar(value="test_run")
        self.command_var = ttk.StringVar(value="Immediate Output (Stable or Unstable Readings)")

        settings_grid = ttk.Frame(settings_labelframe)
        settings_grid.pack(fill=X)

        ttk.Label(settings_grid, text="Output Frequency (seconds):").grid(row=0, column=0, sticky=E, padx=5, pady=5)
        ttk.Entry(settings_grid, textvariable=self.freq_var, width=10).grid(row=0, column=1, sticky=W, pady=5)

        ttk.Label(settings_grid, text="Filename:").grid(row=1, column=0, sticky=E, padx=5, pady=5)
        ttk.Entry(settings_grid, textvariable=self.filename_var, width=20).grid(row=1, column=1, sticky=W, pady=5)

        ttk.Label(settings_grid, text="Command to Send:").grid(row=3, column=0, sticky=E, padx=5, pady=5)
        command_options = ["Immediate Output (Stable or Unstable Readings)", "Immediate Output (Stable Readings Only)"]
        self.command_combobox = ttk.Combobox(settings_grid, textvariable=self.command_var, values=command_options, width=50, state="readonly")
        self.command_combobox.grid(row=3, column=1, sticky=W, pady=5)

        # --- Balances Frame
        balances_labelframe = ttk.Labelframe(main_frame, text="Balances", padding=10)
        balances_labelframe.pack(fill=X, pady=(0, 20))

        self.balance_frame = ttk.Frame(balances_labelframe)
        self.balance_frame.pack(fill=X)

        self.save_path_label = ttk.Label(
            main_frame,
            text=f"Data is being saved to: {self.data_folder}",
            foreground="red"
        )
        self.save_path_label.pack(pady=(0, 10))

        # --- Log Output
        self.log = ttk.ScrolledText(main_frame, width=80, height=20, font=("Consolas", 10))
        self.log.pack(fill=BOTH, expand=True)

        # --- Other Variables
        self.ser_objects = []
        self.running = False
        self.setup_balances()


    def stop_measurements(self):
        self.running = False
        self.log_message("Stopping measurements...")
        self.root.after(1000, self.root.destroy)

    def setup_balances(self):
        for widget in self.balance_frame.winfo_children():
            widget.destroy()

        self.port_vars = []
        self.name_vars = []
        self.preview_vars = []

        self.add_balance()  # Start with one balance at launch

    def log_message(self, message):
        self.log.configure(state='normal')
        self.log.insert('end', f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
        self.log.see('end')
        self.log.configure(state='disabled')

    def start_measurements(self):
        for preview_var, preview_entry in self.preview_vars:
            preview_entry.configure(state="disabled")

        if self.running:
            ttk.messagebox.showinfo("Already running", "Measurements are already running.")
            return

        self.ser_objects = []

        for port_var in self.port_vars:
            selected = port_var.get()
            port = self.get_device_from_selection(selected)
            try:
                ser = serial.Serial(port, 9600, timeout=1)
                ser.reset_input_buffer()
                ser.reset_output_buffer()
                ser.flush()

                command = self.get_actual_command()
                ser.write((command + "\r\n").encode())
                time.sleep(0.2)

                self.ser_objects.append(ser)
            except Exception as e:
                self.log_message(f"Warning: Could not pre-ping {port}: {e}")

        time.sleep(1)

        filename = self.filename_var.get().strip()
        invalid_chars = r'[\\\/:*?"<>|]'
        if not filename or " " in filename or re.search(invalid_chars, filename):
            self.log_message("Invalid filename. Do not use spaces or special characters \\ / : * ? \" < > |")
            return

        filepath = os.path.join(self.data_folder, f"{filename}.csv")
        while os.path.exists(filepath):
            new_filename = simpledialog.askstring(
                "File Exists",
                "Filename already exists.\nPlease enter a new filename:",
                parent=self.root
            )
            if not new_filename:
                self.log_message("Recording canceled. No new filename provided.")
                return
            filepath = os.path.join(self.data_folder, f"{new_filename}.csv")
            self.filename_var.set(new_filename)

        self.filepath = filepath
        self.running = True
        threading.Thread(target=self.measure_loop, daemon=True).start()

    def measure_loop(self):
        try:
            filepath = self.filepath
            frequency = self.freq_var.get()

            balance_names = []
            for p, n in zip(self.port_vars, self.name_vars):
                if n.get().strip():
                    balance_names.append(n.get().strip())
                else:
                    selected = p.get()
                    if "(" in selected and ")" in selected:
                        port_clean = selected.split("(")[0].strip()
                        balance_names.append(port_clean)
                    else:
                        balance_names.append(selected)

            fieldnames = ['Timestamp'] + [f'{name}_Weight' for name in balance_names]

            with open(filepath, 'w', newline='') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()

            self.log_message("Starting measurements...")

            while self.running:
                print_data = {"Timestamp": time.strftime('%Y-%m-%d %H:%M:%S')}
                for i, ser in enumerate(self.ser_objects):
                    try:
                        weight = self.get_weight(ser)
                        if weight is None:
                            print_data[f'{balance_names[i]}_Weight'] = "NA"
                        else:
                            print_data[f'{balance_names[i]}_Weight'] = weight
                        self.log_message(f"Measured {balance_names[i]}: {print_data[f'{balance_names[i]}_Weight']}")
                    except Exception as e:
                        self.log_message(f"Error reading {balance_names[i]}: {e}")
                        print_data[f'{balance_names[i]}_Weight'] = "Error"

                with open(filepath, 'a', newline='') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                    writer.writerow(print_data)

                time.sleep(frequency)

        except Exception as e:
            self.log_message(f"Error: {str(e)}")

        finally:
            for ser in self.ser_objects:
                try:
                    ser.close()
                except:
                    pass
            self.running = False

    def add_balance(self):
        i = len(self.port_vars)

        ttk.Label(self.balance_frame, text=f"Balance {i+1} Port:").grid(row=i, column=0, sticky=E, padx=5, pady=2)

        port_var = ttk.StringVar()

        # Get list of all available ports
        ports = serial.tools.list_ports.comports()
        all_port_choices = []
        for port in ports:
            description = port.description
            device = port.device
            display_str = f"{device} - {description}"
            all_port_choices.append(display_str)

        # Filter out already-selected devices
        selected_devices = set()
        for p in self.port_vars:
            sel = p.get()
            if sel:
                selected_devices.add(self.get_device_from_selection(sel))

        available_port_choices = []
        for choice in all_port_choices:
            device = self.get_device_from_selection(choice)
            if device not in selected_devices:
                available_port_choices.append(choice)

        if available_port_choices:
            port_var.set(available_port_choices[0])
        else:
            port_var.set("No Ports Available")

        combobox_width = max(30, max([len(c) for c in all_port_choices]) if all_port_choices else 30)
        port_combobox = ttk.Combobox(
            self.balance_frame,
            textvariable=port_var,
            values=available_port_choices,
            width=combobox_width,
            state="readonly"
        )
        port_combobox.grid(row=i, column=1, sticky=W, padx=5)

        ttk.Label(self.balance_frame, text="Name (optional):").grid(row=i, column=2, sticky=E, padx=5, pady=2)
        name_var = ttk.StringVar()
        ttk.Entry(self.balance_frame, textvariable=name_var, width=15).grid(row=i, column=3, sticky=W, padx=5)

        ttk.Label(self.balance_frame, text="Preview:").grid(row=i, column=4, sticky=E, padx=5, pady=2)
        preview_var = ttk.StringVar(value="--")
        preview_entry = ttk.Entry(self.balance_frame, textvariable=preview_var, width=15, state="readonly")
        preview_entry.grid(row=i, column=5, sticky=W, padx=5)

        ttk.Button(
            self.balance_frame,
            text="Ping",
            command=lambda i=i: self.ping_balance(i),
            bootstyle="secondary"
        ).grid(row=i, column=6, sticky=W, padx=5)

        self.port_vars.append(port_var)
        self.name_vars.append(name_var)
        self.preview_vars.append((preview_var, preview_entry))


    def remove_balance(self):
        if len(self.port_vars) <= 1:
            tkinter.messagebox.showinfo("Cannot Remove", "At least one balance must remain.")
            return

        # Remove last balance widgets
        for widget in self.balance_frame.grid_slaves(row=len(self.port_vars)-1):
            widget.destroy()

        # Remove last balance data
        self.port_vars.pop()
        self.name_vars.pop()
        self.preview_vars.pop()

        # Now update the dropdowns for all remaining balances
        self.refresh_balance_dropdowns()

    def get_weight(self, ser):
        command = self.get_actual_command()
        
        ser.write((command + "\r\n").encode()) 
        
        start_time = time.time()
        while time.time() - start_time < 2:
            line = ser.readline().decode('utf-8').rstrip()
            if line.startswith('+') or line.startswith('-'):
                weight_data = line.split()
                try:
                    signed_value = weight_data[0] + weight_data[1]
                    return float(signed_value)
                except (IndexError, ValueError):
                    return None
        return None

    def ping_balance(self, index):
        try:
            selected = self.port_vars[index].get()
            if "(" in selected and ")" in selected:
                #port = selected.split("(")[-1].replace(")", "")
                port = selected.split(" - ")[0]
            else:
                port = selected

            command = self.get_actual_command()

            weight = None
            with serial.Serial(port, 9600, timeout=1) as ser: 
                ser.flush()
                ser.write((command + "\r\n").encode())

                start_time = time.time()
                while time.time() - start_time < 0.5:
                    line = ser.readline().decode('utf-8').rstrip()
                    if line.startswith('+') or line.startswith('-'):
                        weight_data = line.split()
                        signed_value = weight_data[0] + weight_data[1]
                        weight = float(signed_value)
                        break

            if weight is not None:
                self.preview_vars[index][0].set(f"{weight:.2f} g")
            else:
                self.preview_vars[index][0].set("No Balance Detected")

        except Exception:
            self.preview_vars[index][0].set("No Balance Detected")

    def refresh_balance_dropdowns(self):
        # Recomputes dropdown lists when a device is removed
        selected_devices = set()
        for port_var in self.port_vars:
            sel = port_var.get()
            if sel:
                selected_devices.add(self.get_device_from_selection(sel))

        ports = serial.tools.list_ports.comports()
        all_port_choices = []
        for port in ports:
            description = port.description
            device = port.device
            display_str = f"{device} - {description}"
            all_port_choices.append(display_str)

        for i, port_var in enumerate(self.port_vars):
            current_selected = port_var.get()
            current_device = self.get_device_from_selection(current_selected)
            available_choices = []
            for choice in all_port_choices:
                device = self.get_device_from_selection(choice)
                if device == current_device or device not in selected_devices:
                    available_choices.append(choice)
            combobox = self.balance_frame.grid_slaves(row=i, column=1)[0]  
            combobox.configure(values=available_choices)


    def get_device_from_selection(self, selection):
        return selection.split(" - ")[0] if " - " in selection else selection

    def get_actual_command(self):
        selected_text = self.command_var.get()
        if "Stable or Unstable" in selected_text:
            return "O8"  # Output regardless of stability
        elif "Stable Readings Only" in selected_text:
            return "S"   # Only output if stable
        else:
            return "O8"  # Default fallback

if __name__ == "__main__":
    app = ttk.Window(themename="flatly")
    app.state('zoomed')
    BalanceLoggerApp(app)
    app.mainloop()

