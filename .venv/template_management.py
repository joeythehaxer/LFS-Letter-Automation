import os
import json
import pandas as pd
from datetime import datetime
from docx import Document
from custom_logging import Logger
import config

class TemplateManager:
    def __init__(self, templates_dir, logger):
        self.templates_dir = templates_dir
        self.logger = logger
        self.config = config.config

    def load_template(self, template_name):
        template_path = os.path.join(self.templates_dir, f"{template_name}.docx")
        document = Document(template_path)
        return document

    def determine_next_letter(self, data):
        try:
            self.logger.log('info', f"Checking which letter to send for: {data[self.config['NAME_COLUMN']]}")
            if pd.isna(data[self.config['LETTER_1_COLUMN']]) or data[self.config['LETTER_1_COLUMN']] == "":
                self.logger.log('info', f"First letter needs to be sent to: {data[self.config['NAME_COLUMN']]}")
                return self.config['TEMPLATE_GROUP1']['LETTER_1_TEMPLATE']
            elif pd.isna(data[self.config['LETTER_2_COLUMN']]) or data[self.config['LETTER_2_COLUMN']] == "":
                self.logger.log('info', f"Second letter needs to be sent to: {data[self.config['NAME_COLUMN']]}")
                return self.config['TEMPLATE_GROUP1']['LETTER_2_TEMPLATE']
            elif pd.isna(data[self.config['LETTER_3_COLUMN']]) or data[self.config['LETTER_3_COLUMN']] == "":
                self.logger.log('info', f"Third letter needs to be sent to: {data[self.config['NAME_COLUMN']]}")
                return self.config['TEMPLATE_GROUP1']['LETTER_3_TEMPLATE']
            else:
                self.logger.log('info', f"All letters have been sent to: {data[self.config['NAME_COLUMN']]}")
                return None  # If all letters have been sent
        except KeyError as e:
            self.logger.log('error', f"Missing key in configuration or data: {e}")
            raise

    def update_excel(self, data, letter_type):
        df = pd.read_excel(self.config['LOCAL_EXCEL_FILE'])
        row_index = df[df[self.config['NAME_COLUMN']] == data[self.config['NAME_COLUMN']]].index[0]
        if letter_type == self.config['TEMPLATE_GROUP1']['LETTER_1_TEMPLATE']:
            df.at[row_index, self.config['LETTER_1_COLUMN']] = datetime.now().strftime("%d %B %Y")
        elif letter_type == self.config['TEMPLATE_GROUP1']['LETTER_2_TEMPLATE']:
            df.at[row_index, self.config['LETTER_2_COLUMN']] = datetime.now().strftime("%d %B %Y")
        elif letter_type == self.config['TEMPLATE_GROUP1']['LETTER_3_TEMPLATE']:
            df.at[row_index, self.config['LETTER_3_COLUMN']] = datetime.now().strftime("%d %B %Y")
        df.to_excel(self.config['LOCAL_EXCEL_FILE'], index=False)
        self.logger.log('info', f"Updated Excel file for: {data[self.config['NAME_COLUMN']]}")
