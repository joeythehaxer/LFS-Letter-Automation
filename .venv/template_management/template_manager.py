import os
import json
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from docx import Document
import pandas as pd
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
                raise FileNotFoundError(
                    f"default_config.json not found. Please create it with the necessary configurations.")
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
            wb = self.get_workbook()
            sheet = wb[self.config.EXCEL_SHEET_NAME]
            row_idx = self.find_row_index(sheet, data[self.config.ADDRESS_COLUMN])
            if row_idx is None:
                raise ValueError(f"No matching row found for address: {data[self.config.ADDRESS_COLUMN]}")

            col_name = self.get_column_name_for_letter_type(letter_type)
            if col_name:
                self.update_cell(sheet, row_idx, col_name, f"sent letter {datetime.now().strftime('%d %B %Y')}")
                wb.save(self.config.LOCAL_EXCEL_FILE)
        except Exception as e:
            self.logger.log('error', f"Error updating Excel file: {e}")
            raise

    def get_workbook(self):
        return load_workbook(self.config.LOCAL_EXCEL_FILE)

    def find_row_index(self, sheet, address):
        for idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
            if row[self.get_column_index(sheet, self.config.ADDRESS_COLUMN) - 1] == address:
                return idx
        return None

    def get_column_index(self, sheet, column_name):
        return {v: k for k, v in enumerate(next(sheet.iter_rows(values_only=True)), 1)}[column_name]

    def get_column_name_for_letter_type(self, letter_type):
        for key, value in self.config.TEMPLATE_GROUP1.items():
            if value == letter_type:
                return getattr(self.config, key.replace('TEMPLATE', 'COLUMN'))
        return None

    def update_cell(self, sheet, row_index, column_name, value):
        column_letter = get_column_letter(self.get_column_index(sheet, column_name))
        cell = f"{column_letter}{row_index}"
        sheet[cell] = value
        self.logger.log('info', f"Updated {column_name} with '{value}' for row {row_index}")
