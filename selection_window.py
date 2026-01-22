import tkinter as tk
from tkinter import ttk, messagebox, Toplevel
import json

from test_dashboard import TestDashboardApp
from gui_flasher import resource_path


class SelectionWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Select an Option")
        self.root.geometry("300x400")

        try:
            with open(resource_path("config.json"), "r") as f:
                self.config = json.load(f)
        except FileNotFoundError:
            messagebox.showerror("Error", "config.json not found!")
            self.root.quit()
            return
        except json.JSONDecodeError:
            messagebox.showerror("Error", "Invalid config.json format!")
            self.root.quit()
            return

        self.create_widgets()

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill="both", expand=True)

        self.listbox = tk.Listbox(
            main_frame,
            font=("TkDefaultFont", 12)
        )
        self.listbox.pack(pady=5, fill="both", expand=True)

        for assembly_name in self.config.get("assemblies", {}).keys():
            self.listbox.insert(tk.END, assembly_name)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=5)

        select_button = ttk.Button(
            button_frame,
            text="Select",
            command=self.open_dashboard
        )
        select_button.pack(side="right", padx=5)

        cancel_button = ttk.Button(
            button_frame,
            text="Cancel",
            command=self.root.quit
        )
        cancel_button.pack(side="right")

    def open_dashboard(self):
        selection_indices = self.listbox.curselection()
        if not selection_indices:
            messagebox.showwarning(
                "No Selection",
                "Please select a match style from the list."
            )
            return

        selected_assembly_name = self.listbox.get(selection_indices[0])

        # Hide selection window
        self.root.withdraw()

        # Open dashboard as Toplevel
        dashboard_window = Toplevel(self.root)
        TestDashboardApp(
            dashboard_window,
            self.config,
            selected_assembly_name
        )
