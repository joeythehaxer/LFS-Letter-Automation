import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from data_collection.data_collector import DataCollector
from custom_logging.logger import Logger
from printing.printer import Printer
from template_management.template_manager import TemplateManager
from letter_generation.letter_generator import LetterGenerator
from config.settings import load_defaults

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

    def create_templates_tab(self):
        # Implementation similar to general_tab for handling templates
        pass

    def create_filters_tab(self):
        # Implementation for creating filters dynamically
        pass

    def create_teams_tab(self):
        # Implementation for Teams Excel settings
        pass

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_path = file_path
            self.excel_file_entry.delete(0, tk.END)
            self.excel_file_entry.insert(0, file_path)
            self.load_excel_file(file_path)

    def load_excel_file(self, file_path):
        try:
            self.df = self.data_collector.collect_data()  # Includes handling for Excel filters
            self.sheet_names = pd.ExcelFile(file_path).sheet_names
            self.sheet_menu['values'] = self.sheet_names
            if self.sheet_names:
                self.sheet_var.set(self.sheet_names[0])
        except Exception as e:
            messagebox.showerror("Error", f"Error loading Excel file: {e}")
            self.logger.log('error', f"Error loading Excel file: {e}")

    def on_sheet_change(self, *args):
        # Additional logic can be implemented here
        pass

    def on_header_row_change(self, *args):
        # Additional logic can be implemented here
        pass

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
