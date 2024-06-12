import pandas as pd
import config
from custom_logging import Logger


class DataCollector:
    def __init__(self, logger):
        self.logger = logger

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
        if config.USE_TEAMS_EXCEL:
            self.logger.log('info', 'Collecting data from Teams-stored Excel sheet')
            return None  # Placeholder, actual data will come from TeamsExcelWatcher
        else:
            self.logger.log('info', 'Collecting data from local Excel file')
            df = pd.read_excel(config.LOCAL_EXCEL_FILE, sheet_name=config.EXCEL_SHEET_NAME)
            return df

    def filter_data(self, df, filter_value):
        """
        Filters the DataFrame to include only rows where any of the letter columns are empty
        and applies additional filter conditions.
        """
        letter_columns = [config.LETTER_1_COLUMN, config.LETTER_2_COLUMN, config.LETTER_3_COLUMN]
        print(letter_columns)
        filter_condition = df[letter_columns].isnull().any(axis=1) | (df[letter_columns] == '').any(axis=1)
        print(filter_condition)
        # Apply additional filter condition
        additional_filter_condition = df[config.NEW_FILTER_COLUMN] == filter_value
        filter_condition = filter_condition & additional_filter_condition

        return df[filter_condition]
