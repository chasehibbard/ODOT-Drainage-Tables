"""
gui.py — Simple GUI for the ODOT Drainage Tables Generator.

Provides a window with:
  - File picker buttons for each input file
  - A "Generate" button to produce the formatted output
  - Status messages to show progress
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import Workbook

from parser import parse_da_summary, parse_inlets
from formatter_da import create_da_summary_sheet
from formatter_inlet import create_inlet_sheet


class DrainageApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ODOT Drainage Tables Generator")
        self.root.geometry("600x350")
        self.root.resizable(False, False)

        # File path variables
        self.da_file = tk.StringVar(value="")
        self.inlet_file = tk.StringVar(value="")

        self._build_ui()

    def _build_ui(self):
        # --- Title ---
        title = tk.Label(
            self.root,
            text="ODOT Drainage Tables Generator",
            font=("Arial", 16, "bold"),
        )
        title.pack(pady=(20, 5))

        subtitle = tk.Label(
            self.root,
            text="Select your OpenRoads Designer export files below",
            font=("Arial", 10),
        )
        subtitle.pack(pady=(0, 20))

        # --- File selection frame ---
        file_frame = tk.Frame(self.root)
        file_frame.pack(padx=20, fill="x")

        # DA Summary file picker
        tk.Label(file_frame, text="DA Summary Table:", font=("Arial", 10)).grid(
            row=0, column=0, sticky="w", pady=5
        )
        da_entry = tk.Entry(file_frame, textvariable=self.da_file, width=45)
        da_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(
            file_frame, text="Browse...", command=self._browse_da
        ).grid(row=0, column=2, pady=5)

        # Inlet file picker
        tk.Label(file_frame, text="Inlets Table:", font=("Arial", 10)).grid(
            row=1, column=0, sticky="w", pady=5
        )
        inlet_entry = tk.Entry(file_frame, textvariable=self.inlet_file, width=45)
        inlet_entry.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(
            file_frame, text="Browse...", command=self._browse_inlet
        ).grid(row=1, column=2, pady=5)

        # --- Generate button ---
        self.generate_btn = tk.Button(
            self.root,
            text="Generate ODOT Tables",
            font=("Arial", 12, "bold"),
            bg="#4CAF50",
            fg="white",
            padx=20,
            pady=10,
            command=self._generate,
        )
        self.generate_btn.pack(pady=25)

        # --- Status label ---
        self.status = tk.Label(
            self.root, text="", font=("Arial", 10), fg="gray"
        )
        self.status.pack()

    def _browse_da(self):
        path = filedialog.askopenfilename(
            title="Select DA Summary Table export",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.da_file.set(path)

    def _browse_inlet(self):
        path = filedialog.askopenfilename(
            title="Select Inlets Table export",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.inlet_file.set(path)

    def _generate(self):
        da_path = self.da_file.get().strip()
        inlet_path = self.inlet_file.get().strip()

        # Validate at least one file is selected
        if not da_path and not inlet_path:
            messagebox.showwarning(
                "No files selected",
                "Please select at least one input file.",
            )
            return

        # Validate files exist
        if da_path and not os.path.isfile(da_path):
            messagebox.showerror("File not found", f"Cannot find:\n{da_path}")
            return
        if inlet_path and not os.path.isfile(inlet_path):
            messagebox.showerror("File not found", f"Cannot find:\n{inlet_path}")
            return

        # Ask where to save
        save_path = filedialog.asksaveasfilename(
            title="Save output as...",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="ODOT_Drainage_Tables.xlsx",
        )
        if not save_path:
            return

        self.status.config(text="Processing...", fg="blue")
        self.root.update()

        try:
            # Parse input files
            da_data = []
            inlet_data = []

            if da_path:
                da_data = parse_da_summary(da_path)
                self.status.config(text=f"Read {len(da_data)} drainage areas...")
                self.root.update()

            if inlet_path:
                inlet_data = parse_inlets(inlet_path)
                self.status.config(text=f"Read {len(inlet_data)} inlets...")
                self.root.update()

            # Create output workbook
            wb = Workbook()
            wb.remove(wb.active)  # Remove the default blank sheet

            if da_data:
                create_da_summary_sheet(wb, da_data)

            if inlet_data:
                create_inlet_sheet(wb, inlet_data, da_data if da_data else None)

            wb.save(save_path)

            self.status.config(text=f"Saved to: {save_path}", fg="green")
            messagebox.showinfo(
                "Success",
                f"ODOT drainage tables generated!\n\n"
                f"DA records: {len(da_data)}\n"
                f"Inlet records: {len(inlet_data)}\n\n"
                f"Saved to:\n{save_path}",
            )

        except Exception as e:
            self.status.config(text="Error — see details", fg="red")
            messagebox.showerror("Error", f"Something went wrong:\n\n{e}")
