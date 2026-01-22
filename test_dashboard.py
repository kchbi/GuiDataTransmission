import os
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import serial
import serial.tools.list_ports
import subprocess
import threading
import queue
import shlex
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font

from gui_flasher import resource_path


class TestDashboardApp:
    def __init__(self, root, config, assembly_name):
        self.root = root
        self.config = config
        self.assembly_name = assembly_name

        # Matcher state
        self.matcher1_var = tk.IntVar(value=0)
        self.matcher2_var = tk.IntVar(value=0)
        self._m1_after = None
        self._m2_after = None

        if self.assembly_name not in self.config.get('assemblies', {}):
            messagebox.showerror(
                "Config Error",
                f"Assembly '{self.assembly_name}' not found in config.json."
            )
            self.root.destroy()
            return

        self.assembly_config = self.config['assemblies'][self.assembly_name]

        self.root.title(f"Match Network Analyzer - {self.assembly_name}")
        self.root.geometry("1020x720")

        # Serial state
        self.serial_port = None
        self.is_running = False
        self.gui_queue = queue.Queue()

        self.item_map = {}

        self.create_widgets()
        self.populate_table()
        self.update_serial_ports()

        self.root.after(100, self.process_queue)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    # --------------------------------------------------
    # UI SETUP
    # --------------------------------------------------

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)

        main_frame.columnconfigure(0, weight=1, minsize=300)
        main_frame.columnconfigure(1, weight=2)
        main_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=0)

        control_pane = self._create_control_panel(main_frame)
        control_pane.grid(row=0, column=0, sticky="nswe", padx=(0, 10))

        data_pane = self._create_data_table(main_frame)
        data_pane.grid(row=0, column=1, sticky="nswe")

        matcher_pane = self._create_matcher_panel(main_frame)
        matcher_pane.grid(row=1, column=1, sticky="we", pady=(10, 0))

        style = ttk.Style()
        style.map("Treeview", background=[('selected', '#347083')])
        self.data_table.tag_configure('fail', background='#FFC0CB')
        self.data_table.tag_configure('pass', background='#90EE90')

    # --------------------------------------------------

    def _create_control_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Controls", padding=10)
        frame.pack_propagate(False)

        # Connection
        conn_frame = ttk.Frame(frame)
        conn_frame.pack(fill="x", pady=5)

        ttk.Label(conn_frame, text="Port:").pack(side="left")
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(
            conn_frame, textvariable=self.port_var,
            state="readonly", width=15
        )
        self.port_combo.pack(side="left", fill="x", expand=True, padx=5)

        self.connect_button = ttk.Button(
            conn_frame, text="Connect",
            command=self.toggle_serial_connection
        )
        self.connect_button.pack(side="left")

        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)

        ttk.Button(frame, text="Change Matching Network",
                   command=self.on_closing).pack(fill="x", pady=5)

        ttk.Button(frame, text="Perform Test",
                   command=self.perform_test).pack(fill="x", pady=5)

        ttk.Button(frame, text="Stop Test",
                   command=self.stop_test).pack(fill="x", pady=5)

        # Corner buttons
        motor_frame = ttk.LabelFrame(frame, text="Corner Positions", padding=10)
        motor_frame.pack(fill="x", pady=10)

        motor_frame.columnconfigure(0, weight=1)
        motor_frame.columnconfigure(1, weight=1)

        ttk.Button(motor_frame, text="Min-Min",
                   command=self.Go_To_Corner1).grid(row=0, column=0, sticky="ew", padx=5, pady=2)
        ttk.Button(motor_frame, text="Max-Min",
                   command=self.Go_To_Corner2).grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        ttk.Button(motor_frame, text="Min-Max",
                   command=self.Go_To_Corner3).grid(row=1, column=0, sticky="ew", padx=5, pady=2)
        ttk.Button(motor_frame, text="Max-Max",
                   command=self.Go_To_Corner4).grid(row=1, column=1, sticky="ew", padx=5, pady=2)

        self.flash_button = ttk.Button(
            frame, text=f"Flash '{self.assembly_name}'",
            command=self.flash_firmware
        )
        self.flash_button.pack(fill="x", pady=10)

        ttk.Button(
            frame,
            text="Export Test Report",
            command=self.export_report
        ).pack(fill="x", pady=5)

        ttk.Label(frame, text="Observation Window:").pack(anchor="w")
        self.log_text = scrolledtext.ScrolledText(
            frame, height=10, state='disabled', wrap=tk.WORD
        )
        self.log_text.pack(fill="both", expand=True)

        return frame

    # --------------------------------------------------

    def _create_matcher_panel(self, parent):
        frame = ttk.LabelFrame(parent, text="Matcher Position", padding=10)

        ttk.Label(frame, text="Matcher 1").grid(row=0, column=0, sticky="w")
        ttk.Scale(
            frame, from_=0, to=1000,
            variable=self.matcher1_var,
            command=self.on_matcher1_change
        ).grid(row=0, column=1, sticky="ew", padx=10)
        ttk.Label(frame, textvariable=self.matcher1_var).grid(row=0, column=2, sticky="e")

        ttk.Label(frame, text="Matcher 2").grid(row=1, column=0, sticky="w")
        ttk.Scale(
            frame, from_=0, to=1000,
            variable=self.matcher2_var,
            command=self.on_matcher2_change
        ).grid(row=1, column=1, sticky="ew", padx=10)
        ttk.Label(frame, textvariable=self.matcher2_var).grid(row=1, column=2, sticky="e")

        frame.columnconfigure(1, weight=1)
        return frame

    # --------------------------------------------------
    # MATCHER HANDLERS
    # --------------------------------------------------

    def on_matcher1_change(self, value):
        if self._m1_after:
            self.root.after_cancel(self._m1_after)
        self._m1_after = self.root.after(
            50,
            lambda: self._send_command(f"CMD:SET_M1 {int(float(value))}")
        )

    def on_matcher2_change(self, value):
        if self._m2_after:
            self.root.after_cancel(self._m2_after)
        self._m2_after = self.root.after(
            50,
            lambda: self._send_command(f"CMD:SET_M2 {int(float(value))}")
        )

    # --------------------------------------------------
    # TABLE / TEST
    # --------------------------------------------------

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
        for param in self.assembly_config.get('parameters', []):
            item = self.data_table.insert(
                "", "end",
                values=(param['name'], param['min'], "---", param['max'])
            )
            self.item_map[param['name']] = item

    def perform_test(self):
        self._send_command("CMD:PERFORM_TEST")

    def stop_test(self):
        self._send_command("CMD:STOP_TEST")

    def Go_To_Corner1(self): self._send_command("CMD:CORNER1")
    def Go_To_Corner2(self): self._send_command("CMD:CORNER2")
    def Go_To_Corner3(self): self._send_command("CMD:CORNER3")
    def Go_To_Corner4(self): self._send_command("CMD:CORNER4")

    # --------------------------------------------------
    # EXPORT REPORT (EXCEL)
    # --------------------------------------------------

    def _get_table_data(self):
        data = []
        for item in self.data_table.get_children():
            data.append(self.data_table.item(item, "values"))
        return data

    def export_report(self):
        table_data = self._get_table_data()
        if not table_data:
            messagebox.showwarning("No Data", "No test data to export.")
            return

        default_name = f"TestReport_{self.assembly_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not file_path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Test Report"

        headers = ["Test Parameter", "Min", "Measured", "Max", "Result"]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)

        for param, min_v, meas, max_v in table_data:
            try:
                meas_f = float(meas)
                min_f = float(min_v)
                max_f = float(max_v)
                result = "PASS" if min_f <= meas_f <= max_f else "FAIL"
            except:
                result = "N/A"
            ws.append([param, min_v, meas, max_v, result])

        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 22

        wb.save(file_path)
        messagebox.showinfo("Export Successful", f"Report saved:\n{file_path}")

    # --------------------------------------------------
    # FLASHING
    # --------------------------------------------------

    def flash_firmware(self):
        firmware = resource_path(self.assembly_config['firmwareFile'])
        cmd = self.assembly_config['flashCommand'].replace(
            "{firmware_path}", f'"{firmware}"'
        )

        def run():
            subprocess.run(shlex.split(cmd), check=True)
            messagebox.showinfo("Success", "Firmware flashed successfully.")

        threading.Thread(target=run, daemon=True).start()

    # --------------------------------------------------
    # SERIAL
    # --------------------------------------------------

    def _send_command(self, cmd):
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.write((cmd + '\n').encode())

    def update_serial_ports(self):
        self.port_combo['values'] = [
            p.device for p in serial.tools.list_ports.comports()
        ]

    def toggle_serial_connection(self):
        self.stop_serial() if self.serial_port else self.start_serial()

    def start_serial(self):
        self.serial_port = serial.Serial(self.port_var.get(), 115200, timeout=1)
        self.is_running = True
        threading.Thread(target=self.read_serial_data, daemon=True).start()
        self.connect_button.config(text="Disconnect")

    def stop_serial(self):
        self.is_running = False
        if self.serial_port:
            self.serial_port.close()
            self.serial_port = None
        self.connect_button.config(text="Connect")

    def read_serial_data(self):
        while self.is_running and self.serial_port:
            line = self.serial_port.readline().decode(errors='ignore').strip()
            if line:
                self.gui_queue.put(line)

    def process_queue(self):
        while not self.gui_queue.empty():
            self.log_to_monitor(self.gui_queue.get() + "\n")
        self.root.after(100, self.process_queue)

    # --------------------------------------------------
    # UTIL
    # --------------------------------------------------

    def log_to_monitor(self, msg):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, msg)
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')

    def on_closing(self):
        self.stop_serial()
        self.root.master.deiconify()
        self.root.destroy()
