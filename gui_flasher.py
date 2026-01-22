import tkinter as tk
from tkinter import ttk, scrolledtext
import json
import subprocess
import os
import threading
import sys

# ---------------- PYINSTALLER SAFE PATH ----------------
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ---------------- CONFIG ----------------
CONFIG_FILE = "config.json"

def load_config():
    try:
        with open(resource_path(CONFIG_FILE), "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return [{'assemblyName': f"ERROR: '{CONFIG_FILE}' not found."}]
    except json.JSONDecodeError:
        return [{'assemblyName': f"ERROR: Invalid JSON in '{CONFIG_FILE}'."}]

def flash_firmware(assembly_config, log_widget):
    def log(message):
        log_widget.insert(tk.END, message + '\n')
        log_widget.see(tk.END)

    assembly_name = assembly_config['assemblyName']
    firmware_relative_path = assembly_config['firmwareFile']
    flash_command_template = assembly_config['flashCommand']

    firmware_abs_path = resource_path(firmware_relative_path)

    if not os.path.exists(firmware_abs_path):
        log("\n--- FAILURE ---")
        log(f"Firmware file not found at: {firmware_abs_path}")
        return

    command_to_run = flash_command_template.format(
        firmware_path=f'"{firmware_abs_path}"'
    )

    log("---------------------------------------------------")
    log(f"Selected Assembly: {assembly_name}")
    log(f"Firmware: {firmware_relative_path}")
    log(f"Executing command: {command_to_run}")
    log("---------------------------------------------------")
    log("Flashing in progress... please wait.")

    result = subprocess.run(
        command_to_run,
        shell=True,
        capture_output=True,
        text=True,
        encoding="latin-1"
    )

    if result.stdout:
        log("\n--- Programmer Output ---")
        log(result.stdout)
    if result.stderr:
        log("\n--- Programmer Errors ---")
        log(result.stderr)

    if result.returncode == 0:
        log("\n--- SUCCESS ---")
        log("Firmware flashed successfully.")
    else:
        log("\n--- FAILURE ---")
        log(f"An error occurred (return code: {result.returncode}).")

# ---------------- GUI ----------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Firmware Flashing Utility")
        self.geometry("700x500")

        self.assemblies = load_config()
        self.assembly_names = [asm['assemblyName'] for asm in self.assemblies]

        self.create_widgets()

    def create_widgets(self):
        control_frame = ttk.Frame(self, padding="10")
        control_frame.pack(fill=tk.X)

        ttk.Label(control_frame, text="Select Assembly:").pack(side=tk.LEFT, padx=(0, 10))
        self.combo_box = ttk.Combobox(
            control_frame,
            values=self.assembly_names,
            state="readonly",
            width=40
        )
        if self.assembly_names:
            self.combo_box.current(0)
        self.combo_box.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.flash_button = ttk.Button(
            control_frame,
            text="Flash Firmware",
            command=self.start_flash_thread
        )
        self.flash_button.pack(side=tk.LEFT, padx=(10, 0))

        log_frame = ttk.Frame(self, padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_widget = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            state="normal"
        )
        self.log_widget.pack(fill=tk.BOTH, expand=True)

    def start_flash_thread(self):
        selected_name = self.combo_box.get()
        if not selected_name or "ERROR" in selected_name:
            self.log_widget.insert(tk.END, "Please select a valid assembly.\n")
            return

        selected_assembly_config = next(
            item for item in self.assemblies
            if item["assemblyName"] == selected_name
        )

        self.log_widget.delete("1.0", tk.END)
        self.flash_button.config(state="disabled")

        flash_thread = threading.Thread(
            target=flash_firmware,
            args=(selected_assembly_config, self.log_widget)
        )
        flash_thread.start()

        self.after(100, self.check_flash_thread, flash_thread)

    def check_flash_thread(self, thread):
        if thread.is_alive():
            self.after(100, self.check_flash_thread, thread)
        else:
            self.flash_button.config(state="normal")

if __name__ == "__main__":
    app = App()
    app.mainloop()
