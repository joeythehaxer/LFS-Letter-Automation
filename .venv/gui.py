import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import json
import os
from letter_generation import LetterGenerator
from template_management import TemplateManager
from custom_logging import Logger
from printing import Printer
from data_collection import DataCollector
import config

class LetterAutomationGUI:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.logger = Logger()
        self.data_collector = DataCollector(self.logger, self.config)
        self.printer = Printer(self.config['PRINT_SERVER_DIR'], self.logger)
        self.template_manager = TemplateManager(self.config['TEMPLATES_DIR'], self.logger)
        self.letter_generator = LetterGenerator(self.config, self.logger, self.printer, self.template_manager)
        self.df = None
        self.file_path = None

        self.create_widgets()
        self.load_defaults()

    def create_widgets(self):
        self.file_label = tk.Label(self.root, text="Select Excel File:")
        self.file_label.pack()

        self.file_button = tk.Button(self.root, text="Browse", command=self.browse_file)
        self.file_button.pack()

        self.sheet_label = tk.Label(self.root, text="Select Sheet:")
        self.sheet_label.pack()

        self.sheet_var = tk.StringVar()
        self.sheet_menu = tk.OptionMenu(self.root, self.sheet_var, '')
        self.sheet_menu.pack()
        self.sheet_var.trace('w', self.on_sheet_change)

        self.header_row_label = tk.Label(self.root, text="Select Header Row (1-based index):")
        self.header_row_label.pack()

        self.header_row_var = tk.IntVar(value=self.config['HEADER_ROW'])
        self.header_row_spinbox = tk.Spinbox(self.root, from_=1, to=100, textvariable=self.header_row_var)
        self.header_row_spinbox.pack()
        self.header_row_var.trace('w', self.on_header_row_change)

        self.generate_button = tk.Button(self.root, text="Generate and Print Letters", command=self.generate_and_print_letters)
        self.generate_button.pack(pady=20)

        self.set_default_button = tk.Button(self.root, text="Set as Default", command=self.set_as_default)
        self.set_default_button.pack(pady=10)

        self.reset_button = tk.Button(self.root, text="Reset to Defaults", command=self.reset_to_defaults)
        self.reset_button.pack(pady=10)

    def load_defaults(self):
        self.config = config.load_defaults()

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_path = file_path
            self.load_excel_file(file_path)

    def load_excel_file(self, file_path):
        try:
            self.sheet_names = pd.ExcelFile(file_path).sheet_names
            self.sheet_menu['menu'].delete(0, 'end')
            for sheet in self.sheet_names:
                self.sheet_menu['menu'].add_command(label=sheet, command=tk._setit(self.sheet_var, sheet))
            if self.sheet_names:
                self.sheet_var.set(self.sheet_names[0])
                self.load_sheet_columns(self.sheet_names[0])
        except Exception as e:
            messagebox.showerror("Error", f"Error loading Excel file: {e}")

    def on_sheet_change(self, *args):
        self.load_sheet_columns(self.sheet_var.get())

    def on_header_row_change(self, *args):
        self.load_sheet_columns(self.sheet_var.get())

    def load_sheet_columns(self, sheet_name):
        try:
            header_row = self.header_row_var.get() - 1
            self.df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=header_row)
        except Exception as e:
            messagebox.showerror("Error", f"Error loading sheet columns: {e}")

    def generate_and_print_letters(self):
        if not self.file_path:
            messagebox.showerror("Error", "No Excel file selected.")
            return

        try:
            df = pd.read_excel(self.file_path, sheet_name=self.sheet_var.get(), header=self.header_row_var.get() - 1)
            filtered_data = self.data_collector.filter_data(df)
            self.letter_generator.generate_and_print_letters(filtered_data.to_dict(orient='records'))
            messagebox.showinfo("Success", "Letters have been generated and printed.")
        except Exception as e:
            messagebox.showerror("Error", f"Error generating and printing letters: {e}")

    def reset_to_defaults(self):
        self.load_defaults()
        self.load_excel_file(self.config['LOCAL_EXCEL_FILE'])
        self.sheet_var.set(self.config['EXCEL_SHEET_NAME'])
        self.header_row_var.set(self.config['HEADER_ROW'])

    def set_as_default(self):
        if not self.file_path:
            messagebox.showerror("Error", "No Excel file selected.")
            return

        self.config['LOCAL_EXCEL_FILE'] = self.file_path
        self.config['EXCEL_SHEET_NAME'] = self.sheet_var.get()
        self.config['HEADER_ROW'] = self.header_row_var.get()
        with open(config.DEFAULT_CONFIG_PATH, 'w') as f:
            json.dump(self.config, f)
        messagebox.showinfo("Success", "Defaults have been set.")

def run_gui(config):
    root = tk.Tk()
    app = LetterAutomationGUI(root, config)
    root.mainloop()

if __name__ == "__main__":
    config = config.load_defaults()
    run_gui(config)
