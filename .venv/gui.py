import tkinter as tk
from tkinter import filedialog, messagebox
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
        self.source_label = tk.Label(self.root, text="Select Data Source:")
        self.source_label.pack()

        self.source_var = tk.StringVar(value="local")
        self.local_radio = tk.Radiobutton(self.root, text="Local Excel File", variable=self.source_var, value="local")
        self.teams_radio = tk.Radiobutton(self.root, text="Teams Excel Sheet", variable=self.source_var, value="teams")
        self.local_radio.pack()
        self.teams_radio.pack()

        self.filter_label = tk.Label(self.root, text="Apply Filters:")
        self.filter_label.pack()

        self.filter_frame = tk.Frame(self.root)
        self.filter_frame.pack()

        self.letter1_var = tk.StringVar(value="")
        self.letter2_var = tk.StringVar(value="")
        self.letter3_var = tk.StringVar(value="")

        self.letter1_entry = tk.Entry(self.filter_frame, textvariable=self.letter1_var, width=30)
        self.letter1_label = tk.Label(self.filter_frame, text="Filter for 1st Letter:")
        self.letter2_entry = tk.Entry(self.filter_frame, textvariable=self.letter2_var, width=30)
        self.letter2_label = tk.Label(self.filter_frame, text="Filter for 2nd Letter:")
        self.letter3_entry = tk.Entry(self.filter_frame, textvariable=self.letter3_var, width=30)
        self.letter3_label = tk.Label(self.filter_frame, text="Filter for 3rd Letter:")

        self.letter1_label.grid(row=0, column=0, padx=5, pady=5)
        self.letter1_entry.grid(row=0, column=1, padx=5, pady=5)
        self.letter2_label.grid(row=1, column=0, padx=5, pady=5)
        self.letter2_entry.grid(row=1, column=1, padx=5, pady=5)
        self.letter3_label.grid(row=2, column=0, padx=5, pady=5)
        self.letter3_entry.grid(row=2, column=1, padx=5, pady=5)

        self.generate_button = tk.Button(self.root, text="Generate and Print Letters",
                                         command=self.generate_and_print_letters)
        self.generate_button.pack(pady=20)

    def generate_and_print_letters(self):
        source = self.source_var.get()
        filters = {
            config.LETTER_1_COLUMN: self.letter1_var.get(),
            config.LETTER_2_COLUMN: self.letter2_var.get(),
            config.LETTER_3_COLUMN: self.letter3_var.get()
        }

        if source == "teams":
            # Assuming TeamsExcelWatcher will be implemented to get data from Teams
            messagebox.showinfo("Info", "Fetching data from Teams is not yet implemented.")
        else:
            df = self.data_collector.collect_data()
            filtered_data = self.data_collector.filter_data(df, filters)
            data = filtered_data.to_dict(orient='records')
            self.letter_generator.generate_and_print_letters(data)
            messagebox.showinfo("Success", "Letters have been generated and printed.")


if __name__ == "__main__":
    root = tk.Tk()
    app = LetterAutomationGUI(root)
    root.mainloop()
