import os
import json
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from docx import Document
from custom_logging.logger import Logger

class TemplateManager:
    def __init__(self, config, logger):
        self.config = config
        self.logger = logger
        self.load_defaults()

    def load_defaults(self):
        try:
            self.logger.log('info', 'Loading default configuration')
            if os.path.exists('default_config.json'):
                with open('default_config.json', 'r') as f:
                    self.default_config = json.load(f)
                    self.logger.log('info', f"Loaded default config: {self.default_config}")
            else:
                raise FileNotFoundError(f"default_config.json not found. Please create it with the necessary configurations.")
        except Exception as e:
            self.logger.log('error', f"Error loading defaults: {e}")
            raise

    def load_template(self, template_name):
        template_path = os.path.join(self.config.TEMPLATES_DIR, f"{template_name}.docx")
        document = Document(template_path)
        return document

    def determine_next_letter(self, data):
        try:
            self.logger.log('info', f"Checking which letter to send for: {data[self.config.NAME_COLUMN]}")
            if pd.isna(data[self.config.LETTER_1_COLUMN]) or data[self.config.LETTER_1_COLUMN] == "":
                self.logger.log('info', f"First letter needs to be sent to: {data[self.config.NAME_COLUMN]}")
                return self.config.TEMPLATE_GROUP1['LETTER_1_TEMPLATE']
            elif pd.isna(data[self.config.LETTER_2_COLUMN]) or data[self.config.LETTER_2_COLUMN] == "":
                self.logger.log('info', f"Second letter needs to be sent to: {data[self.config.NAME_COLUMN]}")
                return self.config.TEMPLATE_GROUP1['LETTER_2_TEMPLATE']
            elif pd.isna(data[self.config.LETTER_3_COLUMN]) or data[self.config.LETTER_3_COLUMN] == "":
                self.logger.log('info', f"Third letter needs to be sent to: {data[self.config.NAME_COLUMN]}")
                return self.config.TEMPLATE_GROUP1['LETTER_3_TEMPLATE']
            else:
                self.logger.log('info', f"All letters have been sent to: {data[self.config.NAME_COLUMN]}")
                return None
        except KeyError as e:
            self.logger.log('error', f"Missing key in configuration or data: {e}")
            raise

    def update_excel(self, data, letter_type):
        try:
            wb = load_workbook(self.config.LOCAL_EXCEL_FILE)
            sheet = wb[self.config.EXCEL_SHEET_NAME]
            col_idx = {v: k for k, v in enumerate(next(sheet.iter_rows(values_only=True)), 1)}  # Map column names to indices

            row_idx = None
            for row in sheet.iter_rows(values_only=True):
                if row[col_idx[self.config.ADDRESS_COLUMN] - 1] == data[self.config.ADDRESS_COLUMN]:
                    row_idx = row[0]
                    break

            if row_idx is None:
                self.logger.log('error', f"No matching row found for address: {data[self.config.ADDRESS_COLUMN]}")
                raise ValueError(f"No matching row found for address: {data[self.config.ADDRESS_COLUMN]}")

            col_name = None
            if letter_type == self.config.TEMPLATE_GROUP1['LETTER_1_TEMPLATE']:
                col_name = self.config.LETTER_1_COLUMN
            elif letter_type == self.config.TEMPLATE_GROUP1['LETTER_2_TEMPLATE']:
                col_name = self.config.LETTER_2_COLUMN
            elif letter_type == self.config.TEMPLATE_GROUP1['LETTER_3_TEMPLATE']:
                col_name = self.config.LETTER_3_COLUMN

            if col_name:
                sheet[f"{get_column_letter(col_idx[col_name])}{row_idx}"] = f"sent letter {datetime.now().strftime('%d %B %Y')}"
                wb.save(self.config.LOCAL_EXCEL_FILE)
                self.logger.log('info', f"Updated {col_name} with 'sent letter {datetime.now().strftime('%d %B %Y')}' for {data[self.config.NAME_COLUMN]}")
            else:
                self.logger.log('error', "No valid letter column to update.")
        except Exception as e:
            self.logger.log('error', f"Error updating Excel file: {e}")
            raise
