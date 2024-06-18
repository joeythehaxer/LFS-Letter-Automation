import pandas as pd
import json
import os
from custom_logging import Logger

DEFAULT_CONFIG_PATH = 'default_config.json'

class DataCollector:
    def __init__(self, logger, config):
        self.logger = logger
        self.config = config

    def parse_excel_data(self, excel_data):
        self.logger.log('info', 'Parsing Excel data')
        values = excel_data['values']
        headers = values[0]
        data = values[1:]
        df = pd.DataFrame(data, columns=headers)
        return df

    def collect_data(self):
        self.logger.log('info', 'Collecting data from local Excel file')
        df = pd.read_excel(self.config['LOCAL_EXCEL_FILE'], sheet_name=self.config['EXCEL_SHEET_NAME'], header=self.config['HEADER_ROW'] - 1)
        return df

    def filter_data(self, df):
        letter_columns = [self.config['LETTER_1_COLUMN'], self.config['LETTER_2_COLUMN'], self.config['LETTER_3_COLUMN']]
        filter_condition = df[letter_columns].isnull().any(axis=1) | (df[letter_columns] == '').any(axis=1)

        for filter_cond in self.config['FILTERS']:
            column = filter_cond['column']
            value = filter_cond['value']
            filter_condition &= (df[column] == value)

        return df[filter_condition]
