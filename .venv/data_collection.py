import pandas as pd
import json
import os
from custom_logging import Logger

DEFAULT_CONFIG_PATH = 'default_config.json'


class DataCollector:
    def __init__(self, logger):
        self.logger = logger
        self.load_defaults()

    def load_defaults(self):
        if os.path.exists(DEFAULT_CONFIG_PATH):
            with open(DEFAULT_CONFIG_PATH, 'r') as f:
                self.default_config = json.load(f)
        else:
            raise FileNotFoundError(f"{DEFAULT_CONFIG_PATH} not found. Please create it with the necessary configurations.")

    def parse_excel_data(self, excel_data):
        self.logger.log('info', 'Parsing Excel data')
        values = excel_data['values']
        headers = values[0]
        data = values[1:]
        df = pd.DataFrame(data, columns=headers)
        return df

    def get_resident_data(self, excel_data):
        df = self.parse_excel_data(excel_data)
        return df.to_dict(orient='records')

    def collect_data(self):
        self.logger.log('info', 'Collecting data from local Excel file')
        df = pd.read_excel(self.default_config['LOCAL_EXCEL_FILE'], sheet_name=self.default_config['EXCEL_SHEET_NAME'], header=self.default_config['HEADER_ROW'] - 1)
        print("DataFrame Columns:", df.columns.tolist())  # Debug statement
        return df

    def filter_data(self, df):
        """
        Filters the DataFrame to include only rows where any of the letter columns are empty
        and applies additional filter conditions.
        """
        letter_columns = [self.default_config['LETTER_1_COLUMN'], self.default_config['LETTER_2_COLUMN'], self.default_config['LETTER_3_COLUMN']]
        print("Expected Letter Columns:", letter_columns)  # Debug statement

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
