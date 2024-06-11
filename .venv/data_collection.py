import pandas as pd
from datetime import datetime
from custom_logging import Logger
import config

class DataCollector:
    def __init__(self, logger):
        self.logger = logger

    def parse_excel_data(self, excel_data):
        self.logger.log('info', 'Parsing Excel data')
        # Assuming excel_data is in the required format
        # Convert the data to a pandas DataFrame
        values = excel_data['values']
        print(values)
        headers = values[0]
        data = values[1:]
        df = pd.DataFrame(data, columns=headers)
        return df.to_dict(orient='records')

    def get_resident_data(self, excel_data):
        return self.parse_excel_data(excel_data)

    def get_last_letter_date(self, resident):
        # Logic to get last letter date from resident data
        return resident.get('last_letter_date', None)

    def collect_data(self):
        if config.USE_TEAMS_EXCEL:
            self.logger.log('info', 'Collecting data from Teams-stored Excel sheet')
            return None  # Placeholder, actual data will come from TeamsExcelWatcher
        else:
            self.logger.log('info', 'Collecting data from local Excel file')
            df = pd.read_excel(config.LOCAL_EXCEL_FILE)
            data = df.to_dict(orient='records')
            return data
