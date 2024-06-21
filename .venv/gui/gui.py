import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from data_collection.data_collector import DataCollector
from custom_logging.logger import Logger
from printing.printer import Printer
from template_management.template_manager import TemplateManager
from letter_generation.letter_generator import LetterGenerator
from config.settings import Settings, load_defaults
import json
class LetterAutomationGUI:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.logger = Logger(config)
        self.data_collector = DataCollector(self.logger, self.config)
        self.printer = Printer(self.config.PRINT_SERVER_DIR, self.logger)
        self.template_manager = TemplateManager(self.config, self.logger)
        self.letter_generator = LetterGenerator(self.config, self.logger, self.printer, self.template_manager)
        self.df = None
        self.file_path = None

        self.root.title("Letter Automation System")
        self.root.geometry("800x600")
        self.style = ttk.Style()
        self.style.theme_use('clam')

        self.create_widgets()
        self.load_defaults()

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both')

        self.general_tab = ttk.Frame(self.notebook)
        self.templates_tab = ttk.Frame(self.notebook)
        self.filters_tab = ttk.Frame(self.notebook)
        self.teams_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.general_tab, text='General')
        self.notebook.add(self.templates_tab, text='Templates')
        self.notebook.add(self.filters_tab, text='Filters')
        self.notebook.add(self.teams_tab, text='Teams Excel')

        self.create_general_tab()
        self.create_templates_tab()
        self.create_filters_tab()
        self.create_teams_tab()

        self.generate_button = ttk.Button(self.root, text="Generate Letters", command=self.generate_letters)
        self.generate_button.pack(pady=10)

    def create_general_tab(self):
        ttk.Label(self.general_tab, text="Select Excel File:").pack(pady=5)
        self.excel_file_entry = ttk.Entry(self.general_tab, width=50)
        self.excel_file_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(self.general_tab, text="Browse", command=self.browse_file).pack(side=tk.LEFT, padx=5)

        ttk.Label(self.general_tab, text="Select Sheet:").pack(pady=5)
        self.sheet_var = tk.StringVar()
        self.sheet_menu = ttk.Combobox(self.general_tab, textvariable=self.sheet_var)
        self.sheet_menu.pack(fill='x', padx=5)
        self.sheet_var.trace('w', self.on_sheet_change)

        ttk.Label(self.general_tab, text="Select Header Row (1-based index):").pack(pady=5)
        self.header_row_var = tk.IntVar(value=self.config.HEADER_ROW)
        self.header_row_spinbox = ttk.Spinbox(self.general_tab, from_=1, to=100, textvariable=self.header_row_var)
        self.header_row_spinbox.pack(fill='x', padx=5)
        self.header_row_var.trace('w', self.on_header_row_change)

        ttk.Button(self.general_tab, text="Set as Default", command=self.set_as_default).pack(pady=10)
        ttk.Button(self.general_tab, text="Reset to Defaults", command=self.reset_to_defaults).pack(pady=10)

    def create_templates_tab(self):
        ttk.Label(self.templates_tab, text="Template Group 1").pack(pady=5)
        self.template1_entry = self.create_template_entry(self.templates_tab, "Letter 1 Template:")
        self.template2_entry = self.create_template_entry(self.templates_tab, "Letter 2 Template:")
        self.template3_entry = self.create_template_entry(self.templates_tab, "Letter 3 Template:")

        ttk.Label(self.templates_tab, text="Template Group 2").pack(pady=5)
        self.template1a_entry = self.create_template_entry(self.templates_tab, "Letter 1A Template:")
        self.template2a_entry = self.create_template_entry(self.templates_tab, "Letter 2A Template:")
        self.template3a_entry = self.create_template_entry(self.templates_tab, "Letter 3A Template:")

    def create_template_entry(self, parent, label_text):
        frame = ttk.Frame(parent)
        frame.pack(fill='x', pady=5)
        ttk.Label(frame, text=label_text).pack(side=tk.LEFT, padx=5)
        entry = ttk.Entry(frame, width=50)
        entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(frame, text="Browse", command=lambda e=entry: self.browse_template(e)).pack(side=tk.LEFT, padx=5)
        return entry

    def create_filters_tab(self):
        self.filters_frame = ttk.Frame(self.filters_tab)
        self.filters_frame.pack(fill='both', expand=True, padx=5, pady=5)

        self.add_filter_button = ttk.Button(self.filters_tab, text="Add Filter", command=self.add_filter)
        self.add_filter_button.pack(pady=5)
        self.load_filters()

    def create_teams_tab(self):
        ttk.Label(self.teams_tab, text="Tenant ID:").pack(pady=5)
        self.tenant_id_entry = self.create_entry(self.teams_tab, self.config.TENANT_ID)

        ttk.Label(self.teams_tab, text="Client ID:").pack(pady=5)
        self.client_id_entry = self.create_entry(self.teams_tab, self.config.CLIENT_ID)

        ttk.Label(self.teams_tab, text="Client Secret:").pack(pady=5)
        self.client_secret_entry = self.create_entry(self.teams_tab, self.config.CLIENT_SECRET)

        ttk.Label(self.teams_tab, text="Excel File ID:").pack(pady=5)
        self.excel_file_id_entry = self.create_entry(self.teams_tab, self.config.EXCEL_FILE_ID)

        ttk.Label(self.teams_tab, text="Excel File Drive:").pack(pady=5)
        self.excel_file_drive_entry = self.create_entry(self.teams_tab, self.config.EXCEL_FILE_DRIVE)

    def create_entry(self, parent, value):
        entry = ttk.Entry(parent, width=50)
        entry.insert(0, value)
        entry.pack(padx=5, pady=5)
        return entry

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_path = file_path
            self.excel_file_entry.delete(0, tk.END)
            self.excel_file_entry.insert(0, file_path)
            self.load_excel_file(file_path)

    def browse_template(self, entry):
        file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)

    def load_excel_file(self, file_path):
        try:
            self.df = self.data_collector.collect_data()  # Adjusted for simplicity, assuming file_path set in config
            self.sheet_names = pd.ExcelFile(file_path).sheet_names  # Assuming file_path still needed for sheet names
            self.sheet_menu['values'] = self.sheet_names
            if self.sheet_names:
                self.sheet_var.set(self.sheet_names[0])
        except Exception as e:
            messagebox.showerror("Error", f"Error loading Excel file: {e}")
            self.logger.log('error', f"Error loading Excel file: {e}")

    def on_sheet_change(self, *args):
        pass  # Placeholder for any additional actions when sheet changes

    def on_header_row_change(self, *args):
        pass  # Placeholder for actions when header row changes

    def add_filter(self):
        frame = ttk.Frame(self.filters_frame)
        frame.pack(fill='x', pady=5)
        ttk.Label(frame, text="Column:").pack(side=tk.LEFT, padx=5)
        column_entry = ttk.Entry(frame)
        column_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(frame, text="Value:").pack(side=tk.LEFT, padx=5)
        value_entry = ttk.Entry(frame)
        value_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(frame, text="Remove", command=frame.destroy).pack(side=tk.LEFT, padx=5)

    def load_filters(self):
        for filter_cond in self.config.FILTERS:
            frame = ttk.Frame(self.filters_frame)
            frame.pack(fill='x', pady=5)
            ttk.Label(frame, text="Column:").pack(side=tk.LEFT, padx=5)
            column_entry = ttk.Entry(frame)
            column_entry.insert(0, filter_cond['column'])
            column_entry.pack(side=tk.LEFT, padx=5)
            ttk.Label(frame, text="Value:").pack(side=tk.LEFT, padx=5)
            value_entry = ttk.Entry(frame)
            value_entry.insert(0, filter_cond['value'])
            value_entry.pack(side=tk.LEFT, padx=5)
            ttk.Button(frame, text="Remove", command=frame.destroy).pack(side=tk.LEFT, padx=5)

    def reset_to_defaults(self):
        self.load_defaults()
        self.excel_file_entry.delete(0, tk.END)
        self.excel_file_entry.insert(0, self.config.LOCAL_EXCEL_FILE)
        self.sheet_var.set(self.config.EXCEL_SHEET_NAME)
        self.header_row_var.set(self.config.HEADER_ROW)

        self.template1_entry.delete(0, tk.END)
        self.template1_entry.insert(0, self.config.TEMPLATE_GROUP1['LETTER_1_TEMPLATE'])
        self.template2_entry.delete(0, tk.END)
        self.template2_entry.insert(0, self.config.TEMPLATE_GROUP1['LETTER_2_TEMPLATE'])
        self.template3_entry.delete(0, tk.END)
        self.template3_entry.insert(0, self.config.TEMPLATE_GROUP1['LETTER_3_TEMPLATE'])

        self.template1a_entry.delete(0, tk.END)
        self.template1a_entry.insert(0, self.config.TEMPLATE_GROUP2['LETTER_1_TEMPLATE'])
        self.template2a_entry.delete(0, tk.END)
        self.template2a_entry.insert(0, self.config.TEMPLATE_GROUP2['LETTER_2_TEMPLATE'])
        self.template3a_entry.delete(0, tk.END)
        self.template3a_entry.insert(0, self.config.TEMPLATE_GROUP2['LETTER_3_TEMPLATE'])

        self.tenant_id_entry.delete(0, tk.END)
        self.tenant_id_entry.insert(0, self.config.TENANT_ID)
        self.client_id_entry.delete(0, tk.END)
        self.client_id_entry.insert(0, self.config.CLIENT_ID)
        self.client_secret_entry.delete(0, tk.END)
        self.client_secret_entry.insert(0, self.config.CLIENT_SECRET)
        self.excel_file_id_entry.delete(0, tk.END)
        self.excel_file_id_entry.insert(0, self.config.EXCEL_FILE_ID)
        self.excel_file_drive_entry.delete(0, tk.END)
        self.excel_file_drive_entry.insert(0, self.config.EXCEL_FILE_DRIVE)

    def set_as_default(self):
        # Update config and possibly write to a file or a persistent store
        self.config.LOCAL_EXCEL_FILE = self.excel_file_entry.get()
        self.config.EXCEL_SHEET_NAME = self.sheet_var.get()
        self.config.HEADER_ROW = self.header_row_var.get()

        self.config.TEMPLATE_GROUP1 = {
            'LETTER_1_TEMPLATE': self.template1_entry.get(),
            'LETTER_2_TEMPLATE': self.template2_entry.get(),
            'LETTER_3_TEMPLATE': self.template3_entry.get()
        }
        self.config.TEMPLATE_GROUP2 = {
            'LETTER_1_TEMPLATE': self.template1a_entry.get(),
            'LETTER_2_TEMPLATE': self.template2a_entry.get(),
            'LETTER_3_TEMPLATE': self.template3a_entry.get()
        }

        self.config.TENANT_ID = self.tenant_id_entry.get()
        self.config.CLIENT_ID = self.client_id_entry.get()
        self.config.CLIENT_SECRET = self.client_secret_entry.get()
        self.config.EXCEL_FILE_ID = self.excel_file_id_entry.get()
        self.config.EXCEL_FILE_DRIVE = self.excel_file_drive_entry.get()

        filters = []
        for child in self.filters_frame.winfo_children():
            column_entry = child.winfo_children()[1]
            value_entry = child.winfo_children()[3]
            filters.append({'column': column_entry.get(), 'value': value_entry.get()})
        self.config.FILTERS = filters

        with open('default_config.json', 'w') as f:
            json.dump(self.config.__dict__, f)
        messagebox.showinfo("Success", "Defaults have been set.")

    def load_defaults(self):
        self.config = load_defaults()

    def generate_letters(self):
        if not self.file_path or self.df is None:
            messagebox.showerror("Error", "No Excel file is loaded or selected.")
            return
        try:
            filtered_data = self.data_collector.filter_data(self.df)  # Apply filters using DataCollector method
            self.letter_generator.generate_and_print_letters(filtered_data.to_dict(orient='records'))
            messagebox.showinfo("Success", "Letters have been generated and printed.")
        except Exception as e:
            self.logger.log('error', f"Error generating letters: {e}")
            messagebox.showerror("Error", f"Error generating letters: {e}")

def run_gui(config):
    root = tk.Tk()
    app = LetterAutomationGUI(root, config)
    try:
        root.mainloop()
    except Exception as e:
        app.logger.log('error', f"Error running GUI: {e}")
        messagebox.showerror("Error", f"Error running GUI: {e}")

if __name__ == "__main__":
    config = load_defaults()
    run_gui(config)
