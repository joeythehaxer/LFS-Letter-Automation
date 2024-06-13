import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import config
from data_collection import DataCollector
from template_management import TemplateManager
from letter_generation import LetterGenerator
from printing import Printer
from custom_logging import Logger


class LetterAutomationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Letter Automation System")

        self.logger = Logger()
        self.df = None  # Initialize the DataFrame attribute

        # Initialize components
        self.data_collector = DataCollector(self.logger)
        self.template_manager = TemplateManager(config.TEMPLATES_DIR, self.logger)
        self.printer = Printer(config.PRINT_SERVER_DIR, self.logger)
        self.letter_generator = LetterGenerator(self.template_manager, self.logger, self.printer)

        self.filter_conditions = config.FILTERS.copy()  # Start with filter conditions from config

        self.create_widgets()

    def create_widgets(self):
        # File selection
        self.file_label = tk.Label(self.root, text="Select Excel File:")
        self.file_label.pack()

        self.file_button = tk.Button(self.root, text="Browse", command=self.browse_file)
        self.file_button.pack()

        # Sheet selection
        self.sheet_label = tk.Label(self.root, text="Select Sheet:")
        self.sheet_label.pack()

        self.sheet_var = tk.StringVar()
        self.sheet_menu = tk.OptionMenu(self.root, self.sheet_var, '')
        self.sheet_menu.pack()
        self.sheet_var.trace('w', self.on_sheet_change)

        # Header row selection
        self.header_row_label = tk.Label(self.root, text="Select Header Row (1-based index):")
        self.header_row_label.pack()

        self.header_row_var = tk.IntVar(value=1)
        self.header_row_spinbox = tk.Spinbox(self.root, from_=1, to=100, textvariable=self.header_row_var)
        self.header_row_spinbox.pack()
        self.header_row_var.trace('w', self.on_header_row_change)

        # Column Selection
        self.column_frame = tk.Frame(self.root)
        self.column_frame.pack(pady=10)

        self.address_label = tk.Label(self.column_frame, text="Address Column:")
        self.address_label.grid(row=0, column=0, padx=5, pady=5)
        self.address_var = tk.StringVar()
        self.address_menu = tk.OptionMenu(self.column_frame, self.address_var, '')
        self.address_menu.grid(row=0, column=1, padx=5, pady=5)

        self.name_label = tk.Label(self.column_frame, text="Name Column:")
        self.name_label.grid(row=1, column=0, padx=5, pady=5)
        self.name_var = tk.StringVar()
        self.name_menu = tk.OptionMenu(self.column_frame, self.name_var, '')
        self.name_menu.grid(row=1, column=1, padx=5, pady=5)

        self.work_order_label = tk.Label(self.column_frame, text="Work Order Column:")
        self.work_order_label.grid(row=2, column=0, padx=5, pady=5)
        self.work_order_var = tk.StringVar()
        self.work_order_menu = tk.OptionMenu(self.column_frame, self.work_order_var, '')
        self.work_order_menu.grid(row=2, column=1, padx=5, pady=5)

        self.letter1_label = tk.Label(self.column_frame, text="1st Letter Column:")
        self.letter1_label.grid(row=3, column=0, padx=5, pady=5)
        self.letter1_var = tk.StringVar()
        self.letter1_menu = tk.OptionMenu(self.column_frame, self.letter1_var, '')
        self.letter1_menu.grid(row=3, column=1, padx=5, pady=5)

        self.letter2_label = tk.Label(self.column_frame, text="2nd Letter Column:")
        self.letter2_label.grid(row=4, column=0, padx=5, pady=5)
        self.letter2_var = tk.StringVar()
        self.letter2_menu = tk.OptionMenu(self.column_frame, self.letter2_var, '')
        self.letter2_menu.grid(row=4, column=1, padx=5, pady=5)

        self.letter3_label = tk.Label(self.column_frame, text="3rd Letter Column:")
        self.letter3_label.grid(row=5, column=0, padx=5, pady=5)
        self.letter3_var = tk.StringVar()
        self.letter3_menu = tk.OptionMenu(self.column_frame, self.letter3_var, '')
        self.letter3_menu.grid(row=5, column=1, padx=5, pady=5)

        # Filter Conditions
        self.filter_frame = tk.Frame(self.root)
        self.filter_frame.pack(pady=10)

        self.add_filter_button = tk.Button(self.filter_frame, text="Add Filter Condition",
                                           command=self.add_filter_condition)
        self.add_filter_button.grid(row=0, column=0, padx=5, pady=5)

        self.filters = []

        self.update_filter_conditions()

        # Generate and Print Button
        self.generate_button = tk.Button(self.root, text="Generate and Print Letters",
                                         command=self.generate_and_print_letters)
        self.generate_button.pack(pady=20)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.load_excel_file(file_path)

    def load_excel_file(self, file_path):
        self.file_path = file_path
        self.sheet_names = pd.ExcelFile(file_path).sheet_names
        self.sheet_menu['menu'].delete(0, 'end')
        for sheet in self.sheet_names:
            self.sheet_menu['menu'].add_command(label=sheet, command=tk._setit(self.sheet_var, sheet))
        self.sheet_var.set(self.sheet_names[0])
        self.load_sheet_columns(self.sheet_names[0])

    def on_sheet_change(self, *args):
        self.load_sheet_columns(self.sheet_var.get())

    def on_header_row_change(self, *args):
        self.load_sheet_columns(self.sheet_var.get())

    def load_sheet_columns(self, sheet_name):
        header_row = self.header_row_var.get() - 1  # Convert to 0-based index
        df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=header_row)
        self.df = df  # Store DataFrame for filter value selection
        columns = df.columns.tolist()
        self.update_column_menu(self.address_menu, self.address_var, columns)
        self.update_column_menu(self.name_menu, self.name_var, columns)
        self.update_column_menu(self.work_order_menu, self.work_order_var, columns)
        self.update_column_menu(self.letter1_menu, self.letter1_var, columns)
        self.update_column_menu(self.letter2_menu, self.letter2_var, columns)
        self.update_column_menu(self.letter3_menu, self.letter3_var, columns)

    def update_column_menu(self, menu, variable, columns):
        menu['menu'].delete(0, 'end')
        for col in columns:
            menu['menu'].add_command(label=col, command=tk._setit(variable, col))
        variable.set(columns[0] if columns else '')

    def add_filter_condition(self):
        row = len(self.filters) + 1

        column_var = tk.StringVar()
        value_var = tk.StringVar()

        column_label = tk.Label(self.filter_frame, text=f"Filter Column {row}:")
        column_label.grid(row=row, column=0, padx=5, pady=5)
        column_menu = tk.OptionMenu(self.filter_frame, column_var, *self.df.columns)
        column_menu.grid(row=row, column=1, padx=5, pady=5)
        column_var.trace('w', lambda *args: self.update_filter_values(column_var, value_var))

        value_label = tk.Label(self.filter_frame, text="Value:")
        value_label.grid(row=row, column=2, padx=5, pady=5)
        value_menu = tk.OptionMenu(self.filter_frame, value_var, '')
        value_menu.grid(row=row, column=3, padx=5, pady=5)

        self.filters.append((column_var, value_var, value_menu))

    def update_filter_conditions(self):
        if self.df is None:
            return

        for filter_cond in self.filter_conditions:
            row = len(self.filters) + 1

            column_var = tk.StringVar(value=filter_cond['column'])
            value_var = tk.StringVar(value=filter_cond['value'])

            column_label = tk.Label(self.filter_frame, text=f"Filter Column {row}:")
            column_label.grid(row=row, column=0, padx=5, pady=5)
            column_menu = tk.OptionMenu(self.filter_frame, column_var, *self.df.columns)
            column_menu.grid(row=row, column=1, padx=5, pady=5)
            column_var.trace('w', lambda *args: self.update_filter_values(column_var, value_var))

            value_label = tk.Label(self.filter_frame, text="Value:")
            value_label.grid(row=row, column=2, padx=5, pady=5)
            value_menu = tk.OptionMenu(self.filter_frame, value_var, *self.df[filter_cond['column']].dropna().unique())
            value_menu.grid(row=row, column=3, padx=5, pady=5)

            self.filters.append((column_var, value_var, value_menu))

    def update_filter_values(self, column_var, value_var):
        selected_column = column_var.get()
        if selected_column:
            unique_values = self.df[selected_column].dropna().unique()
            # Find the corresponding value menu for this column_var
            for filter_tuple in self.filters:
                if filter_tuple[0] == column_var:
                    value_menu = filter_tuple[2]
                    value_menu['menu'].delete(0, 'end')
                    for value in unique_values:
                        value_menu['menu'].add_command(label=value, command=tk._setit(value_var, value))
                    if unique_values.size > 0:
                        value_var.set(unique_values[0])
                    else:
                        value_var.set('')

    def generate_and_print_letters(self):
        # Override config values with selected values from GUI
        config.ADDRESS_COLUMN = self.address_var.get()
        config.NAME_COLUMN = self.name_var.get()
        config.WORK_ORDER_COLUMN = self.work_order_var.get()
        config.LETTER_1_COLUMN = self.letter1_var.get()
        config.LETTER_2_COLUMN = self.letter2_var.get()
        config.LETTER_3_COLUMN = self.letter3_var.get()
        config.EXCEL_SHEET_NAME = self.sheet_var.get()
        config.LOCAL_EXCEL_FILE = self.file_path

        # Collect filter conditions from the GUI
        self.filter_conditions = [{'column': column_var.get(), 'value': value_var.get()} for
                                  column_var, value_var, _ in self.filters]
        config.FILTERS = self.filter_conditions

        df = self.data_collector.collect_data()
        filtered_data = self.data_collector.filter_data(df)
        data = filtered_data.to_dict(orient='records')
        self.letter_generator.generate_and_print_letters(data)
        messagebox.showinfo("Success", "Letters have been generated and printed.")


if __name__ == "__main__":
    root = tk.Tk()
    app = LetterAutomationGUI(root)
    root.mainloop()
