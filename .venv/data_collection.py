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

    def get_last_letter_date(self, resident):
        return resident.get(config.LETTER_1_COLUMN, None)

    def collect_data(self):
        if config.USE_TEAMS_EXCEL:
            self.logger.log('info', 'Collecting data from Teams-stored Excel sheet')
            return None  # Placeholder, actual data will come from TeamsExcelWatcher
        else:
            self.logger.log('info', 'Collecting data from local Excel file')
            df = pd.read_excel(config.LOCAL_EXCEL_FILE, sheet_name=config.EXCEL_SHEET_NAME)
            return df

    def filter_data(self, df, filters):
        for column, value in filters.items():
            df = df[df[column] == value]
        return df

    def collect_and_filter_data(self, filters):
        df = self.collect_data()
        if df is not None:
            df = self.filter_data(df, filters)
            return df.to_dict(orient='records')
        return []
