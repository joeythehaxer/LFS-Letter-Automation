import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import pandas as pd
from config.settings import Settings, load_defaults
from custom_logging.logger import Logger
from template_management.template_manager import TemplateManager
from letter_generation.letter_generator import LetterGenerator
from data_collection.data_collector import DataCollector
from printing.printer import Printer
from watcher.teams_excel_watcher import TeamsExcelWatcher

class LetterAutomationGUI:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.logger = Logger(config)
        self.data_collector = DataCollector(self.logger, self.config)
        self.printer = Printer(self.config.PRINT_SERVER_DIR, self.logger)
        self.template_manager = TemplateManager(self.config, self.logger)
        self.letter_generator = LetterGenerator(self.config, self.logger, self.printer, self.template_manager)

        self.root.title("Letter Automation Configurator")
        self.create_widgets()
        self.load_defaults()

    def create_widgets(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True)

        self.general_frame = ttk.Frame(notebook)
        self.templates_frame = ttk.Frame(notebook)
        self.teams_excel_frame = ttk.Frame(notebook)
        self.filters_frame = ttk.Frame(notebook)

        notebook.add(self.general_frame, text="General")
        notebook.add(self.templates_frame, text="Templates")
        notebook.add(self.teams_excel_frame, text="Teams Excel")
        notebook.add(self.filters_frame, text="Filters")

        self.create_general_tab()
        self.create_templates_tab()
        self.create_teams_excel_tab()
        self.create_filters_tab()

        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill='x', pady=10)
        ttk.Button(button_frame, text="Set as Default", command=self.set_as_default).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Reset to Defaults", command=self.reset_to_defaults).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Generate Letters", command=self.generate_letters).pack(side=tk.RIGHT, padx=5)

    def create_general_tab(self):
        frame = ttk.Frame(self.general_frame)
        frame.pack(fill='both', expand=True, padx=10, pady=10)

        ttk.Label(frame, text="Local Excel File:").grid(row=0, column=0, sticky=tk.W)
        self.excel_file_entry = ttk.Entry(frame)
        self.excel_file_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Browse", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(frame, text="Sheet Name:").grid(row=1, column=0, sticky=tk.W)
        self.sheet_var = tk.StringVar()
        self.sheet_menu = ttk.OptionMenu(frame, self.sheet_var, '')
        self.sheet_menu.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Header Row:").grid(row=2, column=0, sticky=tk.W)
        self.header_row_var = tk.IntVar(value=self.config.HEADER_ROW)
        self.header_row_spinbox = ttk.Spinbox(frame, from_=1, to=100, textvariable=self.header_row_var)
        self.header_row_spinbox.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Use Teams Excel:").grid(row=3, column=0, sticky=tk.W)
        self.use_teams_excel_var = tk.BooleanVar(value=self.config.USE_TEAMS_EXCEL)
        ttk.Checkbutton(frame, variable=self.use_teams_excel_var, command=self.toggle_teams_excel).grid(row=3, column=1, padx=5, pady=5)

    def create_templates_tab(self):
        frame = ttk.Frame(self.templates_frame)
        frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.create_template_row(frame, "Template 1:", 0)
        self.create_template_row(frame, "Template 2:", 1)
        self.create_template_row(frame, "Template 3:", 2)
        self.create_template_row(frame, "Template 1A:", 3)
        self.create_template_row(frame, "Template 2A:", 4)
        self.create_template_row(frame, "Template 3A:", 5)

    def create_template_row(self, frame, label, row):
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky=tk.W)
        entry = ttk.Entry(frame)
        entry.grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Browse", command=lambda: self.browse_template(entry)).grid(row=row, column=2, padx=5, pady=5)
        setattr(self, f"template{row+1}_entry", entry)

    def create_teams_excel_tab(self):
        frame = ttk.Frame(self.teams_excel_frame)
        frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.create_teams_excel_row(frame, "Tenant ID:", 0)
        self.create_teams_excel_row(frame, "Client ID:", 1)
        self.create_teams_excel_row(frame, "Client Secret:", 2)
        self.create_teams_excel_row(frame, "Excel File ID:", 3)
        self.create_teams_excel_row(frame, "Excel File Drive:", 4)

    def create_teams_excel_row(self, frame, label, row):
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky=tk.W)
        entry = ttk.Entry(frame)
        entry.grid(row=row, column=1, padx=5, pady=5)
        setattr(self, f"teams_{label.split()[0].lower()}_entry", entry)

    def create_filters_tab(self):
        frame = ttk.Frame(self.filters_frame)
        frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.filters_container = ttk.Frame(frame)
        self.filters_container.pack(fill='both', expand=True, padx=10, pady=10)

        ttk.Button(frame, text="Add Filter", command=self.add_filter).pack(pady=5)
        self.load_filters()

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.excel_file_entry.delete(0, tk.END)
            self.excel_file_entry.insert(0, file_path)
            self.load_sheet_names(file_path)

    def browse_template(self, entry):
        file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)

    def load_sheet_names(self, file_path):
        try:
            self.sheet_names = pd.ExcelFile(file_path).sheet_names
            menu = self.sheet_menu['menu']
            menu.delete(0, 'end')
            for sheet in self.sheet_names:
                menu.add_command(label=sheet, command=tk._setit(self.sheet_var, sheet))
            if self.sheet_names:
                self.sheet_var.set(self.sheet_names[0])
        except Exception as e:
            messagebox.showerror("Error", f"Error loading sheet names: {e}")

    def add_filter(self, column='', value=''):
        filter_frame = ttk.Frame(self.filters_container)
        filter_frame.pack(fill='x', pady=5)

        column_var = tk.StringVar(value=column)
        column_entry = ttk.Entry(filter_frame, textvariable=column_var)
        column_entry.grid(row=0, column=0, padx=5, pady=5)

        value_var = tk.StringVar(value=value)
        value_entry = ttk.Entry(filter_frame, textvariable=value_var)
        value_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Button(filter_frame, text="Remove", command=filter_frame.destroy).grid(row=0, column=2, padx=5, pady=5)

    def load_filters(self):
        for filter_cond in self.config.FILTERS:
            self.add_filter(filter_cond['column'], filter_cond['value'])

    def set_as_default(self):
        self.update_config_from_gui()
        with open('default_config.json', 'w') as f:
            json.dump(self.config.__dict__, f, indent=4)
        messagebox.showinfo("Success", "Defaults have been set.")

    def reset_to_defaults(self):
        self.load_defaults()
        self.update_gui_from_config()

    def load_defaults(self):
        self.config = load_defaults()
        self.update_gui_from_config()

    def update_gui_from_config(self):
        self.excel_file_entry.delete(0, tk.END)
        self.excel_file_entry.insert(0, self.config.LOCAL_EXCEL_FILE)
        self.sheet_var.set(self.config.EXCEL_SHEET_NAME)
        self.header_row_var.set(self.config.HEADER_ROW)
        self.use_teams_excel_var.set(self.config.USE_TEAMS_EXCEL)
        self.template1_entry.delete(0, tk.END)
        self.template1_entry.insert(0, self.config.TEMPLATE_GROUP1['LETTER_1_TEMPLATE'])
        self.template2_entry.delete(0, tk.END)
        self.template2_entry.insert(0, self.config.TEMPLATE_GROUP1['LETTER_2_TEMPLATE'])
        self.template3_entry.delete(0, tk.END)
        self.template3_entry.insert(0, self.config.TEMPLATE_GROUP1['LETTER_3_TEMPLATE'])
        self.template4_entry.delete(0, tk.END)
        self.template4_entry.insert(0, self.config.TEMPLATE_GROUP2['LETTER_1_TEMPLATE'])
        self.template5_entry.delete(0, tk.END)
        self.template5_entry.insert(0, self.config.TEMPLATE_GROUP2['LETTER_2_TEMPLATE'])
        self.template6_entry.delete(0, tk.END)
        self.template6_entry.insert(0, self.config.TEMPLATE_GROUP2['LETTER_3_TEMPLATE'])
        self.teams_tenant_entry.delete(0, tk.END)
        self.teams_tenant_entry.insert(0, self.config.TENANT_ID)
        self.teams_client_entry.delete(0, tk.END)
        self.teams_client_entry.insert(0, self.config.CLIENT_ID)
        self.teams_client_secret_entry.delete(0, tk.END)
        self.teams_client_secret_entry.insert(0, self.config.CLIENT_SECRET)
        self.teams_excel_file_id_entry.delete(0, tk.END)
        self.teams_excel_file_id_entry.insert(0, self.config.EXCEL_FILE_ID)
        self.teams_excel_file_drive_entry.delete(0, tk.END)
        self.teams_excel_file_drive_entry.insert(0, self.config.EXCEL_FILE_DRIVE)

    def update_config_from_gui(self):
        self.config.LOCAL_EXCEL_FILE = self.excel_file_entry.get()
        self.config.EXCEL_SHEET_NAME = self.sheet_var.get()
        self.config.HEADER_ROW = self.header_row_var.get()
        self.config.USE_TEAMS_EXCEL = self.use_teams_excel_var.get()
        self.config.TEMPLATE_GROUP1['LETTER_1_TEMPLATE'] = self.template1_entry.get()
        self.config.TEMPLATE_GROUP1['LETTER_2_TEMPLATE'] = self.template2_entry.get()
        self.config.TEMPLATE_GROUP1['LETTER_3_TEMPLATE'] = self.template3_entry.get()
        self.config.TEMPLATE_GROUP2['LETTER_1_TEMPLATE'] = self.template4_entry.get()
        self.config.TEMPLATE_GROUP2['LETTER_2_TEMPLATE'] = self.template5_entry.get()
        self.config.TEMPLATE_GROUP2['LETTER_3_TEMPLATE'] = self.template6_entry.get()
        self.config.TENANT_ID = self.teams_tenant_entry.get()
        self.config.CLIENT_ID = self.teams_client_entry.get()
        self.config.CLIENT_SECRET = self.teams_client_secret_entry.get()
        self.config.EXCEL_FILE_ID = self.teams_excel_file_id_entry.get()
        self.config.EXCEL_FILE_DRIVE = self.teams_excel_file_drive_entry.get()

        self.config.FILTERS = []
        for filter_frame in self.filters_container.winfo_children():
            column = filter_frame.children['!entry'].get()
            value = filter_frame.children['!entry2'].get()
            if column and value:
                self.config.FILTERS.append({'column': column, 'value': value})

    def toggle_teams_excel(self):
        if self.use_teams_excel_var.get():
            self.teams_excel_frame.pack(fill='both', expand=True, padx=10, pady=10)
        else:
            self.teams_excel_frame.pack_forget()

    def generate_letters(self):
        self.update_config_from_gui()
        try:
            if self.config.USE_TEAMS_EXCEL:
                watcher = TeamsExcelWatcher(self.data_collector, self.logger, self.config.TENANT_ID, self.config.CLIENT_ID, self.config.CLIENT_SECRET, self.config.EXCEL_FILE_ID, self.config.EXCEL_FILE_DRIVE)
                excel_data = watcher.get_excel_data()
                df = self.data_collector.parse_excel_data(excel_data)
            else:
                df = self.data_collector.collect_data()

            filtered_data = self.data_collector.filter_data(df)
            self.letter_generator.generate_and_print_letters(filtered_data.to_dict(orient='records'))
            messagebox.showinfo("Success", "Letters have been generated and printed.")
        except Exception as e:
            self.logger.log('error', f"Error generating letters: {e}")
            messagebox.showerror("Error", f"Error generating letters: {e}")

def run_gui(config):
    root = tk.Tk()
    app = LetterAutomationGUI(root, config)
    root.mainloop()

if __name__ == "__main__":
    config = load_defaults()
    run_gui(config)
