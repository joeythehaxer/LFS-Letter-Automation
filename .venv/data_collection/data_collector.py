import pandas as pd
import openpyxl
from custom_logging.logger import Logger

class DataCollector:
    def __init__(self, logger, config):
        self.logger = logger
        self.config = config

    def get_active_filters(self, file_path, sheet_name):
        """Retrieve active filters from an Excel sheet using openpyxl."""
        workbook = openpyxl.load_workbook(filename=file_path, data_only=True)
        sheet = workbook[sheet_name]

        filters = {}
        if sheet.auto_filter.ref is not None:
            for column, filter_column in sheet.auto_filter.columns.items():
                filters[column.min_col - 1] = [crit.val for crit in filter_column.filterCriteria]

        return filters

    def apply_filters_to_dataframe(self, df, filters):
        """Apply Excel column filters to a pandas DataFrame."""
        for column_index, criteria in filters.items():
            column_name = df.columns[column_index]  # Get the column name by index
            df = df[df[column_name].isin(criteria)]
        return df

    def collect_data(self):
        """Collect data considering Excel file filters."""
        self.logger.log('info', 'Collecting data from local Excel file')
        try:
            file_path = self.config.LOCAL_EXCEL_FILE
            sheet_name = self.config.EXCEL_SHEET_NAME
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=self.config.HEADER_ROW - 1)

            # Retrieve active filters and apply them to the DataFrame
            filters = self.get_active_filters(file_path, sheet_name)
            if filters:
                df = self.apply_filters_to_dataframe(df, filters)

            return df
        except Exception as e:
            self.logger.log('error', f"Error collecting data from local Excel file: {e}")
            raise
