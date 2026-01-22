import json
import subprocess
import os
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
        print(f"ERROR: Configuration file '{CONFIG_FILE}' not found.")
        return None
    except json.JSONDecodeError:
        print(f"ERROR: Could not decode JSON from '{CONFIG_FILE}'. Please check its format.")
        return None

def present_choices(assemblies):
    print("Please select the assembly to test:")
    for i, assembly in enumerate(assemblies):
        print(f"  {i + 1}: {assembly['assemblyName']}")

    while True:
        try:
            choice = int(input("Enter choice number: "))
            if 1 <= choice <= len(assemblies):
                return choice - 1
            else:
                print("Invalid choice. Please try again.")
        except ValueError:
            print("Invalid input. Please enter a number.")

def flash_firmware(assembly_config):
    assembly_name = assembly_config['assemblyName']
    firmware_relative_path = assembly_config['firmwareFile']
    flash_command_template = assembly_config['flashCommand']

    firmware_abs_path = resource_path(firmware_relative_path)

    if not os.path.exists(firmware_abs_path):
        print("\n--- FAILURE ---")
        print(f"Firmware file not found at: {firmware_abs_path}")
        return False

    command_to_run = flash_command_template.format(
        firmware_path=f'"{firmware_abs_path}"'
    )

    print("\n---------------------------------------------------")
    print(f"Selected Assembly: {assembly_name}")
    print(f"Firmware: {firmware_relative_path}")
    print(f"Executing command: {command_to_run}")
    print("---------------------------------------------------")
    print("Flashing in progress... please wait.")

    result = subprocess.run(
        command_to_run,
        shell=True,
        capture_output=True,
        text=True,
        encoding="latin-1"
    )

    if result.returncode == 0:
        print("\n--- SUCCESS ---")
        print("Firmware flashed successfully.")
        print("Programmer Output:\n", result.stdout)
        return True
    else:
        print("\n--- FAILURE ---")
        print(f"An error occurred (return code: {result.returncode}).")
        print("Standard Output:\n", result.stdout)
        print("Error Output:\n", result.stderr)
        return False

# ---------------- MAIN ----------------
if __name__ == "__main__":
    assemblies = load_config()

    if assemblies:
        selected_index = present_choices(assemblies)
        selected_assembly = assemblies[selected_index]
        flash_firmware(selected_assembly)

    input("\nPress Enter to exit.")
