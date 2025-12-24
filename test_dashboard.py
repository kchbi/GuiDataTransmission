import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import serial
import serial.tools.list_ports
import subprocess
import threading
import queue
import shlex # Import for safely splitting command strings

class TestDashboardApp:
    def __init__(self, root, config, assembly_name):
        self.root = root
        self.config = config
        self.assembly_name = assembly_name
        
        # Check if the selected assembly exists in the config
        if self.assembly_name not in self.config.get('assemblies', {}):
            messagebox.showerror("Config Error", f"Assembly '{self.assembly_name}' not found in config.json.")
            self.root.destroy()
            return
            
        self.assembly_config = self.config['assemblies'][self.assembly_name]

        self.root.title(f"Match Network Analyzer - {self.assembly_name}")
        self.root.geometry("1020x720")

        # Serial Communication State
        self.serial_port = None
        self.serial_thread = None
        self.is_running = False
        self.gui_queue = queue.Queue()
        
        # This dictionary maps a parameter name to its Treeview item ID for efficient updates
        self.item_map = {}

        self.create_widgets()
        self.populate_table()
        self.update_serial_ports()
        
        self.root.after(100, self.process_queue)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    # In test_dashboard.py, replace the entire create_widgets method with this one.
    def reset_table(self):
        """
        Resets the 'Measured' column and removes pass/fail tags from the table.
        """
        self.log_to_monitor("Resetting test results table for new test.\n")
        # Iterate through all the item IDs we have stored in our item_map
        for item_id in self.item_map.values():
            # Set the 'Measured' column value back to '---'
            self.data_table.set(item_id, "Measured", "---")
            # Remove any existing tags (like 'pass' or 'fail') from the item
            self.data_table.item(item_id, tags=())

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # --- ADJUSTED COLUMN WEIGHTS FOR BETTER LAYOUT ---
        # Give the right column (Table) twice the space of the left column (Controls).
        main_frame.columnconfigure(0, weight=1, minsize=280) 
        main_frame.columnconfigure(0, weight=1)
        
        main_frame.rowconfigure(0, weight=1)

        # --- Left Pane: Controls ---
        control_pane = self._create_control_panel(main_frame)
        control_pane.grid(row=0, column=0, sticky="nswe", padx=(0, 10))

        # --- Right Pane: Test Results Table ---
        data_pane = self._create_data_table(main_frame)
        data_pane.grid(row=0, column=1, sticky="nswe")
        
        # --- Configure Color Tags for Pass/Fail ---
        style = ttk.Style()
        style.map("Treeview", background=[('selected', '#347083')]) # Set a custom selection color
        self.data_table.tag_configure('fail', background='#FFC0CB') # Pink
        self.data_table.tag_configure('pass', background='#90EE90') # LightGreen

    def _create_control_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Controls", padding=10)
        frame.pack_propagate(False) # Prevent frame from shrinking
        
        # --- Connection Controls ---
        conn_frame = ttk.Frame(frame)
        conn_frame.pack(fill="x", pady=5)
        ttk.Label(conn_frame, text="Port:").pack(side="left")
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(conn_frame, textvariable=self.port_var, state="readonly", width=15)
        self.port_combo.pack(side="left", fill="x", expand=True, padx=5)
        self.connect_button = ttk.Button(conn_frame, text="Connect", command=self.toggle_serial_connection)
        self.connect_button.pack(side="left")

        # --- Main Application Controls ---
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)

        self.back_button = ttk.Button(frame, text=" Change Matching Network", command=self.on_closing)
        self.back_button.pack(pady=(0, 10), ipady=5, padx=10)

        self.perform_test_button = ttk.Button(frame, text="Perform Test", command=self.perform_test)
        self.perform_test_button.pack(ipady=10, pady=5, padx=10)

        self.stop_test_button = ttk.Button(frame ,text="Stop Test", command=self.stop_test)
        self.stop_test_button.pack(ipady=10, pady=5, padx=10)
        
        # --- Manual Motor Control Frame ---
        motor_control_frame = ttk.LabelFrame(frame, text="Corner Positions", padding=10)
        motor_control_frame.pack(fill="x", pady=10, padx=10)

        motor_control_frame.columnconfigure(0, weight=1)
        motor_control_frame.columnconfigure(1, weight=1)

        self.corner1_button = ttk.Button(motor_control_frame, text="Min-Min", command=self.Go_To_Corner1)
        self.corner1_button.grid(row=0, column=0, sticky="ew", padx=5, pady=2)

        self.corner2_button = ttk.Button(motor_control_frame, text="Max-Min", command=self.Go_To_Corner2)
        self.corner2_button.grid(row=0, column=1, sticky="ew", padx=5, pady=2)

        self.corner3_button = ttk.Button(motor_control_frame, text="Min-Max", command=self.Go_To_Corner3)
        self.corner3_button.grid(row=1, column=0, sticky="ew", padx=5, pady=2)

        self.corner4_button = ttk.Button(motor_control_frame, text="Max-Max", command=self.Go_To_Corner4)
        self.corner4_button.grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        
        # --- Flash Button ---
        self.flash_button = ttk.Button(frame, text=f"Flash '{self.assembly_name}'", command=self.flash_firmware)
        self.flash_button.pack(pady=(10, 5), padx=10)

        # --- Observation Window ---
        ttk.Label(frame, text="Observation Window:").pack(anchor="w", pady=(10, 2))
        self.log_text = scrolledtext.ScrolledText(frame, height=10, state='disabled', wrap=tk.WORD)
        self.log_text.pack(fill="both", expand=True)
        
        return frame

    def _create_data_table(self, parent):
        frame = ttk.LabelFrame(parent, text="Test Results", padding=10)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        cols = ("Test Parameter", "Min", "Measured", "Max")
        self.data_table = ttk.Treeview(frame, columns=cols, show="headings")
        
        for col in cols:
            self.data_table.heading(col, text=col)
            self.data_table.column(col, width=100, anchor="center")
        self.data_table.column("Test Parameter", width=250, anchor="w")

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.data_table.yview)
        vsb.grid(row=0, column=1, sticky='ns')
        self.data_table.configure(yscrollcommand=vsb.set)
        self.data_table.grid(row=0, column=0, sticky="nsew")
        
        return frame

    def populate_table(self):
        """Fills the table with parameter names and limits from the config."""
        for param in self.assembly_config.get('parameters', []):
            name = param.get('name', 'N/A')
            min_val = param.get('min', 'N/A')
            max_val = param.get('max', 'N/A')
            item_id = self.data_table.insert("", "end", values=(name, min_val, "---", max_val))
            self.item_map[name] = item_id
            
    def update_table(self, data_values):
        """Updates the 'Measured' column and applies pass/fail colors."""
        params = self.assembly_config.get('parameters', [])
        if len(data_values) != len(params):
            self.log_to_monitor(f"Warning: Data length mismatch. Expected {len(params)}, got {len(data_values)}.\n")
            return
            
        for i, value_str in enumerate(data_values):
            param_config = params[i]
            param_name = param_config.get('name')
            
            if not param_name or param_name not in self.item_map:
                continue # Skip if parameter name is missing or not in table

            try:
                measured_val = float(value_str) / 1000.0
                min_val = float(param_config.get('min'))
                max_val = float(param_config.get('max'))
                item_id = self.item_map[param_name]
                
                self.data_table.set(item_id, "Measured", f"{measured_val:.3f}")
                
                if min_val <= measured_val <= max_val:
                    self.data_table.item(item_id, tags=('pass',))
                else:
                    self.data_table.item(item_id, tags=('fail',))
            except (ValueError, TypeError) as e:
                self.log_to_monitor(f"Error updating '{param_name}': Could not parse value '{value_str}'.\n")

    def perform_test(self):
        """Sends a generic 'start test' command to the MCU."""
        print("DEBUG: 'Perform Test' button clicked.")
        
        # --- MODIFICATION START ---
        # Call the new method to clear the table before sending the command.
        self.reset_table()
        # --- MODIFICATION END ---
        
        self._send_command("CMD:PERFORM_TEST")
        self.log_to_monitor("Command: PERFORM_TEST sent.\n")
    
    def stop_test(self):
        """Sends a Command to "Stop the test which sets the g_test_is_running = 0 " Stopping the test in the main while Loop """
        print("DEBUG:'Stop Test' button clicked.")
        self._send_command("CMD:STOP_TEST")
        self.log_to_monitor("Command: STOP_TEST sent. \n")
    def Go_To_Corner1(self):
        """Sends a Commands which sets g_test_is_running = 2 This commands the System to take the Motors to Minimum and Minimum Position """
        print("DEBUG:'Corner 1' button Clicked ")
        self._send_command("CMD:CORNER1")
        self.log_to_monitor("Command: Corner1 Sent. \n" )
    def Go_To_Corner2(self):
        """Sends a Commands which sets g_test_is_running = 3 This commands the System to take the Motors to Maximum and Minimum Position """
        print("DEBUG:'Corner 2' button Clicked ")
        self._send_command("CMD:CORNER2")
        self.log_to_monitor("Command: Corner2 Sent. \n" )
    def Go_To_Corner3(self):
        """Sends a Commands which sets g_test_is_running = 4 This commands the System to take the Motors to Maximum and Minimum Position """
        print("DEBUG:'Corner 3' button Clicked ")
        self._send_command("CMD:CORNER3")
        self.log_to_monitor("Command: Corner3 Sent. \n" )
    def Go_To_Corner4(self):
        """Sends a Commands which sets g_test_is_running = 5 This commands the System to take the Motors to Minimum and Maximum Position """
        print("DEBUG:'Corner 4' button Clicked ")
        self._send_command("CMD:CORNER4")
        self.log_to_monitor("Command: Corner4 Sent" )

# In test_dashboard.py, replace the whole flash_firmware method.

    def flash_firmware(self):
        """Flashes the firmware using the flexible command from the config file."""
        try:
            firmware_path = self.assembly_config['firmwareFile']
            flash_command_template = self.assembly_config['flashCommand']
        except KeyError as e:
            messagebox.showerror("Config Error", f"Missing key in config.json for '{self.assembly_name}': {e}")
            return

        final_command_str = flash_command_template.replace('{firmware_path}', firmware_path)
        command_list = shlex.split(final_command_str)
        self.log_to_monitor(f"Executing: {final_command_str}\n")
        
        def run_flash_in_thread():
            self.flash_button.config(state="disabled", text="Flashing...")
            try:
                # --- THE FIX ---
                # We add encoding='utf-8' and errors='ignore' to the subprocess call itself.
                # This handles non-standard characters from the flashing tool's output.
                process = subprocess.run(
                    command_list, 
                    capture_output=True, 
                    text=True, 
                    check=True, 
                    #shell=True,
                    encoding='utf-8',  # Specify the encoding
                    errors='ignore'    # Ignore characters that can't be decoded
                )
                
                # --- IMPROVED SUCCESS MESSAGE ---
                # Display stdout if it exists, otherwise just show a generic success message.
                output_log = process.stdout if process.stdout.strip() else "Command executed successfully."
                self.log_to_monitor(f"STDOUT: {output_log}\n")
                messagebox.showinfo("Success", f"Flashed successfully!\n\n{output_log}")

            except FileNotFoundError:
                 messagebox.showerror("Flashing Failed", f"Command not found: '{command_list[0]}'.\nIs it in your system's PATH or is the path incorrect in config.json?")
            except subprocess.CalledProcessError as e:
                # Also decode stderr with 'ignore' for robustness
                stderr_log = e.stderr if e.stderr else "No error output."
                self.log_to_monitor(f"STDERR: {stderr_log}\n")
                messagebox.showerror("Flashing Failed", f"Error during flashing:\n\n{stderr_log}")
            except Exception as e:
                self.log_to_monitor(f"ERROR: {str(e)}\n")
                messagebox.showerror("Error", f"An unexpected error occurred: {e}")
            finally:
                self.flash_button.config(state="normal", text=f"Flash '{self.assembly_name}'")

        threading.Thread(target=run_flash_in_thread, daemon=True).start()

    def _send_command(self, cmd_string):
        if self.serial_port and self.serial_port.is_open:
            print(f"DEBUG: Port is open. Sending: {cmd_string}\\n") 
            self.serial_port.write((cmd_string + '\n').encode('utf-8'))
        else:
            print("DEBUG: Port is NOT open or connected.")
            messagebox.showwarning("Not Connected", "Cannot send command. Serial port is not connected.")
            self.log_to_monitor("Failed to send command: Not connected.\n")

    def update_serial_ports(self):
        self.port_combo['values'] = [port.device for port in serial.tools.list_ports.comports()]

    def toggle_serial_connection(self):
        if self.serial_port and self.serial_port.is_open:
            self.stop_serial()
        else:
            self.start_serial()

    def start_serial(self):
        port_name = self.port_var.get()
        if not port_name:
            messagebox.showerror("Error", "Please select a serial port.")
            return
        try:
            self.serial_port = serial.Serial(port_name, 115200, timeout=1)
            self.is_running = True
            self.serial_thread = threading.Thread(target=self.read_serial_data, daemon=True)
            self.serial_thread.start()
            self.connect_button.config(text="Disconnect")
            self.log_to_monitor(f"Connected to {port_name}\n")
        except serial.SerialException as e:
            messagebox.showerror("Connection Error", f"Failed to open port {port_name}:\n{e}")

    def stop_serial(self):
        if self.serial_port and self.serial_port.is_open:
            self.is_running = False
            if self.serial_thread:
                self.serial_thread.join(timeout=2)
            self.serial_port.close()
            self.serial_port = None
            self.connect_button.config(text="Connect")
            self.log_to_monitor("Disconnected\n")

    def read_serial_data(self):
        """Runs in a background thread to read from the serial port."""
        while self.is_running and self.serial_port.is_open:
            try:
                # --- THE FIX ---
                # We add errors='ignore' to the decode call.
                # This tells Python to simply discard any bytes that are not valid UTF-8.
                line = self.serial_port.readline().decode('utf-8', errors='ignore').strip()
                
                if line:
                    self.gui_queue.put(line)
            except serial.SerialException:
                self.is_running = False
                self.gui_queue.put("SERIAL_ERROR")
                break
            except Exception as e:
                # Catch any other unexpected errors in the thread and report them
                # This prevents the whole application from crashing silently.
                self.gui_queue.put(f"THREAD_ERROR: {e}")
                self.is_running = False
                break

    def process_queue(self):
        """Processes messages from the queue to update the GUI safely."""
        try:
            while not self.gui_queue.empty():
                line = self.gui_queue.get_nowait()
                
                if line.startswith("THREAD_ERROR:"):
                    self.stop_serial()
                    messagebox.showerror("Thread Error", f"A critical error occurred in the serial reading thread:\n{line}")
                    break

                # The rest of the logic is the same
                self.log_to_monitor(line + '\n') # Log everything raw
                
                if line.startswith("DATA,"):
                    values = line.strip().split(',')[1:]
                    self.update_table(values)
                elif line == "SERIAL_ERROR":
                    self.stop_serial()
                    messagebox.showerror("Serial Error", "Device disconnected or port error.")
                    break
        finally:
            self.root.after(100, self.process_queue)

    def log_to_monitor(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        

    def on_closing(self):
        """
        Handles closing the dashboard. This is called by the window's 'X' button
        and our new 'Change Assembly' button.
        """
        # 1. Stop any running background threads (like the serial reader)
        self.stop_serial()
        
        # 2. Re-show the main selection window that was hidden.
        #    self.root is the Toplevel dashboard window.
        #    self.root.master is the main root window where SelectionWindow lives.
        self.root.master.deiconify()
        
        # 3. Destroy the current dashboard window.
        self.root.destroy()