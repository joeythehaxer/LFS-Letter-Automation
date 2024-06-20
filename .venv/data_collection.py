import pandas as pd
import json
import os
from custom_logging import Logger


class DataCollector:
    def __init__(self, logger, config):
        self.logger = logger
        self.config = config

    def parse_excel_data(self, excel_data):
        self.logger.log('info', 'Parsing Excel data')
        if excel_data is None:
            raise ValueError("Excel data is None")
        try:
            values = excel_data.get('values', None)
            if not values:
                raise ValueError("Excel data does not contain 'values'")

            headers = values[0]
            data = values[1:]
            df = pd.DataFrame(data, columns=headers)
            return df
        except Exception as e:
            self.logger.log('error', f"Error parsing Excel data: {e}")
            self.logger.log('error', f"Excel data structure: {json.dumps(excel_data, indent=2)}")
            raise

    def collect_data(self):
        self.logger.log('info', 'Collecting data from local Excel file')
        try:
            df = pd.read_excel(self.config['LOCAL_EXCEL_FILE'], sheet_name=self.config['EXCEL_SHEET_NAME'],
                               header=self.config['HEADER_ROW'] - 1)
            return df
        except Exception as e:
            self.logger.log('error', f"Error collecting data from local Excel file: {e}")
            raise

    def filter_data(self, df):
        letter_columns = [self.config['LETTER_1_COLUMN'], self.config['LETTER_2_COLUMN'],
                          self.config['LETTER_3_COLUMN']]
        filter_condition = df[letter_columns].isnull().any(axis=1) | (df[letter_columns] == '').any(axis=1)

        for filter_cond in self.config['FILTERS']:
            column = filter_cond['column']
            value = filter_cond['value']
            filter_condition &= (df[column] == value)

        return df[filter_condition]
