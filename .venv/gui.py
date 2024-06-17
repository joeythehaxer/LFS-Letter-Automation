import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import json
import os
from docx import Document  # Import the Document class from python-docx
import re
from email_validator import validate_email, EmailNotValidError
from openai import OpenAI

DEFAULT_CONFIG_PATH = 'default_config.json'

# Load your OpenAI API key from an environment variable or a secure location


client = OpenAI(
    api_key=os.environ['sk-proj-0jmj1NH7ohBRKNRw7mP0T3BlbkFJY0ELu4NE0cEQbnqFbuLU'],
    # this is also the default, it can be omitted
)


class LetterAutomationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Letter Automation System")

        self.df = None  # Initialize the DataFrame attribute

        # Load default values
        self.load_defaults()

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

        self.header_row_var = tk.IntVar(value=self.default_config['HEADER_ROW'])
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

        # Set Defaults Button
        self.set_default_button = tk.Button(self.root, text="Set as Default", command=self.set_as_default)
        self.set_default_button.pack(pady=10)

        # Reset to Defaults Button
        self.reset_button = tk.Button(self.root, text="Reset to Defaults", command=self.reset_to_defaults)
        self.reset_button.pack(pady=10)

        # Load initial state
        self.reset_to_defaults()

    def load_defaults(self):
        if os.path.exists(DEFAULT_CONFIG_PATH):
            with open(DEFAULT_CONFIG_PATH, 'r') as f:
                self.default_config = json.load(f)
        else:
            self.default_config = {
                'ADDRESS_COLUMN': 'Address',
                'NAME_COLUMN': 'Name',
                'WORK_ORDER_COLUMN': 'Work Order Number',
                'LETTER_1_COLUMN': '1ST ACCESS LETTER DATE/CALL',
                'LETTER_2_COLUMN': '2ND ACCESS LETTER DATE/CALL',
                'LETTER_3_COLUMN': '3RD ACCESS LETTER DATE/CALL',
                'EXCEL_SHEET_NAME': 'Sheet1',
                'LOCAL_EXCEL_FILE': 'residents.xlsx',
                'HEADER_ROW': 1,
                'FILTERS': [{'column': 'New Filter Column', 'value': 'Default Value'}],
                'TEMPLATES_DIR': 'templates',
                'PRINT_SERVER_DIR': 'print_server'
            }
            self.save_defaults()

    def save_defaults(self):
        with open(DEFAULT_CONFIG_PATH, 'w') as f:
            json.dump(self.default_config, f)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.load_excel_file(file_path)

    def load_excel_file(self, file_path):
        try:
            self.file_path = file_path
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
        except Exception as e:
            messagebox.showerror("Error", f"Error loading sheet columns: {e}")

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

    def extract_name(self, text):
        titles = ["Mr", "Mrs", "Ms", "Dr", "Miss", "Mister"]
        # Split the text by line breaks
        parts = text.split("\n")
        for part in parts:
            part = part.strip()

            # Remove numbers
            part = re.sub(r'\d+', '', part)

            # Remove emails
            try:
                validate_email(part)
                continue
            except EmailNotValidError:
                pass

                # Use OpenAI API to identify names and titles
                try:
                    client.completions.create(
                        engine="text-davinci-003",
                        prompt=f"Extract the name and title from the following text: '{part}'",
                        max_tokens=50
                    )
                    name = response.choices[0].text.strip()

                    return name
                except Exception as e:
                    print(f"Error during API processing: {e}")

            return None

    def clean_name(self, text):
        if not text:
            return "Resident"
        name = self.extract_name(text)
        if not name:
            return "Resident"
        return name

    def sanitize_filename(self, filename):
        return "".join(c for c in filename if c.isalnum() or c in (" ", ".", "_")).rstrip()

    def replace_placeholders(self, document, data):
        for placeholder, column in self.default_config['PLACEHOLDERS'].items():
            if column in data:
                value = str(data[column])  # Ensure the value is a string
                if placeholder == 'NAME_PLACEHOLDER':
                    value = self.clean_name(value)
                for paragraph in document.paragraphs:
                    if f'{{{{{placeholder}}}}}' in paragraph.text:
                        print(f'Replacing {placeholder} with {value} in paragraph: {paragraph.text}')
                        paragraph.text = paragraph.text.replace(f'{{{{{placeholder}}}}}', value)
                for table in document.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            if f'{{{{{placeholder}}}}}' in cell.text:
                                print(f'Replacing {placeholder} with {value} in cell: {cell.text}')
                            cell.text = cell.text.replace(f'{{{{{placeholder}}}}}', value)
        return document

    def generate_and_print_letters(self):
        try:
            # Override default config values with selected values from GUI
            self.default_config['ADDRESS_COLUMN'] = self.address_var.get()
            self.default_config['NAME_COLUMN'] = self.name_var.get()
            self.default_config['WORK_ORDER_COLUMN'] = self.work_order_var.get()
            self.default_config['LETTER_1_COLUMN'] = self.letter1_var.get()
            self.default_config['LETTER_2_COLUMN'] = self.letter2_var.get()
            self.default_config['LETTER_3_COLUMN'] = self.letter3_var.get()
            self.default_config['EXCEL_SHEET_NAME'] = self.sheet_var.get()
            self.default_config['LOCAL_EXCEL_FILE'] = self.file_path
            self.default_config['HEADER_ROW'] = self.header_row_var.get()

            # Collect filter conditions from the GUI
            self.filter_conditions = [{'column': column_var.get(), 'value': value_var.get()} for
                                      column_var, value_var, _ in self.filters]
            self.default_config['FILTERS'] = self.filter_conditions

            # Collect data
            df = pd.read_excel(self.file_path, sheet_name=self.default_config['EXCEL_SHEET_NAME'],
                               header=self.default_config['HEADER_ROW'] - 1)
            filtered_data = self.filter_data(df)
            data = filtered_data.to_dict(orient='records')

            # Generate letters
            self.generate_letters(data)
            messagebox.showinfo("Success", "Letters have been generated and printed.")
        except Exception as e:
            messagebox.showerror("Error", f"Error generating and printing letters: {e}")

    def reset_to_defaults(self):
        self.load_excel_file(self.default_config['LOCAL_EXCEL_FILE'])
        self.sheet_var.set(self.default_config['EXCEL_SHEET_NAME'])
        self.header_row_var.set(self.default_config['HEADER_ROW'])
        self.address_var.set(self.default_config['ADDRESS_COLUMN'])
        self.name_var.set(self.default_config['NAME_COLUMN'])
        self.work_order_var.set(self.default_config['WORK_ORDER_COLUMN'])
        self.letter1_var.set(self.default_config['LETTER_1_COLUMN'])
        self.letter2_var.set(self.default_config['LETTER_2_COLUMN'])
        self.letter3_var.set(self.default_config['LETTER_3_COLUMN'])
        self.filter_conditions = self.default_config['FILTERS'].copy()
        self.update_filter_conditions()

    def set_as_default(self):
        self.default_config['LOCAL_EXCEL_FILE'] = self.file_path
        self.default_config['EXCEL_SHEET_NAME'] = self.sheet_var.get()
        self.default_config['HEADER_ROW'] = self.header_row_var.get()
        self.default_config['ADDRESS_COLUMN'] = self.address_var.get()
        self.default_config['NAME_COLUMN'] = self.name_var.get()
        self.default_config['WORK_ORDER_COLUMN'] = self.work_order_var.get()
        self.default_config['LETTER_1_COLUMN'] = self.letter1_var.get()
        self.default_config['LETTER_2_COLUMN'] = self.letter2_var.get()
        self.default_config['LETTER_3_COLUMN'] = self.letter3_var.get()
        self.default_config['FILTERS'] = [{'column': column_var.get(), 'value': value_var.get()} for
                                          column_var, value_var, _ in self.filters]
        self.save_defaults()
        messagebox.showinfo("Success", "Defaults have been set.")

    def filter_data(self, df):
        """
        Filters the DataFrame to include only rows where any of the letter columns are empty
        and applies additional filter conditions.
        """
        letter_columns = [self.default_config['LETTER_1_COLUMN'], self.default_config['LETTER_2_COLUMN'],
                          self.default_config['LETTER_3_COLUMN']]

        for col in letter_columns:
            if col not in df.columns:
                raise KeyError(f"Column '{col}' not found in DataFrame columns")

        filter_condition = df[letter_columns].isnull().any(axis=1) | (df[letter_columns] == '').any(axis=1)

        # Apply additional filter conditions from default config
        for filter_cond in self.default_config['FILTERS']:
            column = filter_cond['column']
            value = filter_cond['value']
            if column not in df.columns:
                raise KeyError(f"Column '{column}' not found in DataFrame columns")
            filter_condition = filter_condition & (df[column] == value)

        return df[filter_condition]

    def generate_letters(self, data_list):
        for data in data_list:
            # Load and replace placeholders in the template
            template_name = self.default_config['TEMPLATE_GROUP1']['LETTER_1_TEMPLATE']
            document = self.load_template(template_name)
            personalized_document = self.replace_placeholders(document, data)
            sanitized_name = self.sanitize_filename(f"{data[self.default_config['NAME_COLUMN']]}.docx")
            file_path = os.path.join(self.default_config['PRINT_SERVER_DIR'], sanitized_name)
            file_path = os.path.normpath(file_path)  # Normalize the path to make it Windows-friendly
            personalized_document.save(file_path)
            # Placeholder for actual print logic
            # self.print_letter(file_path)

    def load_template(self, template_name):
        template_path = os.path.join(self.default_config['TEMPLATES_DIR'], f"{template_name}.docx")
        return Document(template_path)

    def replace_placeholders(self, document, data):
        for placeholder, column in self.default_config['PLACEHOLDERS'].items():
            if column in data:
                value = str(data[column])  # Ensure the value is a string
                if placeholder == 'NAME_PLACEHOLDER':
                    value = self.clean_name(value)
                for paragraph in document.paragraphs:
                    if f'{{{{{placeholder}}}}}' in paragraph.text:
                        print(f'Replacing {placeholder} with {value} in paragraph: {paragraph.text}')
                        paragraph.text = paragraph.text.replace(f'{{{{{placeholder}}}}}', value)
                for table in document.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            if f'{{{{{placeholder}}}}}' in cell.text:
                                print(f'Replacing {placeholder} with {value} in cell: {cell.text}')
                            cell.text = cell.text.replace(f'{{{{{placeholder}}}}}', value)
        return document


if __name__ == "__main__":
    root = tk.Tk()
    app = LetterAutomationGUI(root)
    root.mainloop()
