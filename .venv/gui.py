import tkinter as tk
from tkinter import messagebox
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

        # Initialize components
        self.data_collector = DataCollector(self.logger)
        self.template_manager = TemplateManager(config.TEMPLATES_DIR, self.logger)
        self.printer = Printer(config.PRINT_SERVER_DIR, self.logger)
        self.letter_generator = LetterGenerator(self.template_manager, self.logger, self.printer)

        self.create_widgets()

    def create_widgets(self):
        # Data Source Selection
        self.source_label = tk.Label(self.root, text="Select Data Source:")
        self.source_label.pack()

        self.source_var = tk.StringVar(value="local")
        self.local_radio = tk.Radiobutton(self.root, text="Local Excel File", variable=self.source_var, value="local")
        self.teams_radio = tk.Radiobutton(self.root, text="Teams Excel Sheet", variable=self.source_var, value="teams")
        self.local_radio.pack()
        self.teams_radio.pack()

        # Column Selection
        self.column_frame = tk.Frame(self.root)
        self.column_frame.pack(pady=10)

        self.address_label = tk.Label(self.column_frame, text="Address Column:")
        self.address_label.grid(row=0, column=0, padx=5, pady=5)
        self.address_var = tk.StringVar(value=config.ADDRESS_COLUMN)
        self.address_entry = tk.Entry(self.column_frame, textvariable=self.address_var)
        self.address_entry.grid(row=0, column=1, padx=5, pady=5)

        self.name_label = tk.Label(self.column_frame, text="Name Column:")
        self.name_label.grid(row=1, column=0, padx=5, pady=5)
        self.name_var = tk.StringVar(value=config.NAME_COLUMN)
        self.name_entry = tk.Entry(self.column_frame, textvariable=self.name_var)
        self.name_entry.grid(row=1, column=1, padx=5, pady=5)

        self.work_order_label = tk.Label(self.column_frame, text="Work Order Column:")
        self.work_order_label.grid(row=2, column=0, padx=5, pady=5)
        self.work_order_var = tk.StringVar(value=config.WORK_ORDER_COLUMN)
        self.work_order_entry = tk.Entry(self.column_frame, textvariable=self.work_order_var)
        self.work_order_entry.grid(row=2, column=1, padx=5, pady=5)

        self.letter1_label = tk.Label(self.column_frame, text="1st Letter Column:")
        self.letter1_label.grid(row=3, column=0, padx=5, pady=5)
        self.letter1_var = tk.StringVar(value=config.LETTER_1_COLUMN)
        self.letter1_entry = tk.Entry(self.column_frame, textvariable=self.letter1_var)
        self.letter1_entry.grid(row=3, column=1, padx=5, pady=5)

        self.letter2_label = tk.Label(self.column_frame, text="2nd Letter Column:")
        self.letter2_label.grid(row=4, column=0, padx=5, pady=5)
        self.letter2_var = tk.StringVar(value=config.LETTER_2_COLUMN)
        self.letter2_entry = tk.Entry(self.column_frame, textvariable=self.letter2_var)
        self.letter2_entry.grid(row=4, column=1, padx=5, pady=5)

        self.letter3_label = tk.Label(self.column_frame, text="3rd Letter Column:")
        self.letter3_label.grid(row=5, column=0, padx=5, pady=5)
        self.letter3_var = tk.StringVar(value=config.LETTER_3_COLUMN)
        self.letter3_entry = tk.Entry(self.column_frame, textvariable=self.letter3_var)
        self.letter3_entry.grid(row=5, column=1, padx=5, pady=5)

        self.new_filter_label = tk.Label(self.column_frame, text="Filter Column:")
        self.new_filter_label.grid(row=6, column=0, padx=5, pady=5)
        self.new_filter_var = tk.StringVar(value=config.NEW_FILTER_COLUMN)
        self.new_filter_entry = tk.Entry(self.column_frame, textvariable=self.new_filter_var)
        self.new_filter_entry.grid(row=6, column=1, padx=5, pady=5)

        self.new_filter_value_label = tk.Label(self.column_frame, text="Filter Value:")
        self.new_filter_value_label.grid(row=7, column=0, padx=5, pady=5)
        self.new_filter_value_var = tk.StringVar()
        self.new_filter_value_entry = tk.Entry(self.column_frame, textvariable=self.new_filter_value_var)
        self.new_filter_value_entry.grid(row=7, column=1, padx=5, pady=5)

        # Generate and Print Button
        self.generate_button = tk.Button(self.root, text="Generate and Print Letters",
                                         command=self.generate_and_print_letters)
        self.generate_button.pack(pady=20)

    def generate_and_print_letters(self):
        source = self.source_var.get()

        # Override config values with selected values from GUI
        config.ADDRESS_COLUMN = self.address_var.get()
        config.NAME_COLUMN = self.name_var.get()
        config.WORK_ORDER_COLUMN = self.work_order_var.get()
        config.LETTER_1_COLUMN = self.letter1_var.get()
        config.LETTER_2_COLUMN = self.letter2_var.get()
        config.LETTER_3_COLUMN = self.letter3_var.get()
        config.NEW_FILTER_COLUMN = self.new_filter_var.get()
        filter_value = self.new_filter_value_var.get()

        if source == "teams":
            messagebox.showinfo("Info", "Fetching data from Teams is not yet implemented.")
        else:
            df = self.data_collector.collect_data()
            filtered_data = self.data_collector.filter_data(df, filter_value)
            data = filtered_data.to_dict(orient='records')
            self.letter_generator.generate_and_print_letters(data)
            messagebox.showinfo("Success", "Letters have been generated and printed.")


if __name__ == "__main__":
    root = tk.Tk()
    app = LetterAutomationGUI(root)
    root.mainloop()
