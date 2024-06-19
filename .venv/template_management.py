import os
from docx import Document
import json
import pandas as pd
from datetime import datetime

DEFAULT_CONFIG_PATH = 'default_config.json'
EXCEL_FILE_PATH = 'residents.xlsx'  # Update this path accordingly

class TemplateManager:
    def __init__(self, templates_dir, logger):
        self.templates_dir = templates_dir
        self.logger = logger
        self.load_defaults()

    def load_defaults(self):
        if os.path.exists(DEFAULT_CONFIG_PATH):
            with open(DEFAULT_CONFIG_PATH, 'r') as f:
                self.default_config = json.load(f)
                self.logger.log('info', f"Loaded default config: {self.default_config}")
        else:
            raise FileNotFoundError(f"{DEFAULT_CONFIG_PATH} not found. Please create it with the necessary configurations.")

    def load_template(self, template_name):
        template_path = os.path.join(self.templates_dir, f"{template_name}.docx")
        document = Document(template_path)
        return document

    def determine_next_letter(self, data):
        try:
            self.logger.log('info', f"Checking which letter to send for: {data[self.default_config['NAME_COLUMN']]}")
            if pd.isna(data[self.default_config['LETTER_1_COLUMN']]) or data[self.default_config['LETTER_1_COLUMN']] == "":
                self.logger.log('info', f"First letter needs to be sent to: {data[self.default_config['NAME_COLUMN']]}")
                return self.default_config['TEMPLATE_GROUP1']['LETTER_1_TEMPLATE']
            elif pd.isna(data[self.default_config['LETTER_2_COLUMN']]) or data[self.default_config['LETTER_2_COLUMN']] == "":
                self.logger.log('info', f"Second letter needs to be sent to: {data[self.default_config['NAME_COLUMN']]}")
                return self.default_config['TEMPLATE_GROUP1']['LETTER_2_TEMPLATE']
            elif pd.isna(data[self.default_config['LETTER_3_COLUMN']]) or data[self.default_config['LETTER_3_COLUMN']] == "":
                self.logger.log('info', f"Third letter needs to be sent to: {data[self.default_config['NAME_COLUMN']]}")
                return self.default_config['TEMPLATE_GROUP1']['LETTER_3_TEMPLATE']
            else:
                self.logger.log('info', f"All letters have been sent to: {data[self.default_config['NAME_COLUMN']]}")
                return None  # If all letters have been sent
        except KeyError as e:
            self.logger.log('error', f"Missing key in configuration or data: {e}")
            raise

    def update_excel(self, data, letter_type):
        df = pd.read_excel(EXCEL_FILE_PATH)
        row_index = df[df[self.default_config['NAME_COLUMN']] == data[self.default_config['NAME_COLUMN']]].index[0]
        if letter_type == self.default_config['TEMPLATE_GROUP1']['LETTER_1_TEMPLATE']:
            df.at[row_index, self.default_config['LETTER_1_COLUMN']] = datetime.now().strftime("%d %B %Y")
        elif letter_type == self.default_config['TEMPLATE_GROUP1']['LETTER_2_TEMPLATE']:
            df.at[row_index, self.default_config['LETTER_2_COLUMN']] = datetime.now().strftime("%d %B %Y")
        elif letter_type == self.default_config['TEMPLATE_GROUP1']['LETTER_3_TEMPLATE']:
            df.at[row_index, self.default_config['LETTER_3_COLUMN']] = datetime.now().strftime("%d %B %Y")
        df.to_excel(EXCEL_FILE_PATH, index=False)
        self.logger.log('info', f"Updated Excel file for: {data[self.default_config['NAME_COLUMN']]}")
