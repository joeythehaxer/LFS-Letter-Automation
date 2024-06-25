import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from custom_logging.logger import Logger
from zipfile import ZipFile
from lxml import etree
import re

class DataCollector:
    def __init__(self, logger, config):
        self.logger = logger
        self.config = config
        self.active_filters = []

    def get_active_filters(self, file_path, sheet_name):
        """Check for active filter regions in an Excel sheet and log details of each filtered column."""
        try:
            workbook = openpyxl.load_workbook(filename=file_path, data_only=True)
            sheet = workbook[sheet_name]
            if sheet.auto_filter.ref:
                self.logger.log('info', f"Filter range detected: {sheet.auto_filter.ref}")
                filter_range = sheet.auto_filter.ref

                # Extract the filter information directly from the sheet's XML
                with ZipFile(file_path) as z:
                    sheet_name_xml = f'xl/worksheets/sheet{workbook.sheetnames.index(sheet_name) + 1}.xml'
                    with z.open(sheet_name_xml) as f:
                        sheet_xml = f.read()
                        root = etree.fromstring(sheet_xml)

                        namespaces = {'a': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                        auto_filter = root.find('.//a:autoFilter', namespaces)

                        if auto_filter is not None:
                            self.active_filters = []  # Clear any existing filters
                            for filter_col in auto_filter.findall('.//a:filterColumn', namespaces):
                                col_id = filter_col.get('colId')
                                col_letter = get_column_letter(int(col_id) + 1)
                                if filter_col.find('.//a:filters', namespaces) is not None:
                                    filter_values = [f.get('val') for f in filter_col.findall('.//a:filter', namespaces)]
                                    self.logger.log('info', f"Column {col_letter} has filters: {filter_values}")
                                    self.active_filters.append((col_letter, filter_values))

                                if filter_col.find('.//a:customFilters', namespaces) is not None:
                                    custom_filters = [(cf.get('operator'), cf.get('val')) for cf in filter_col.findall('.//a:customFilter', namespaces)]
                                    self.logger.log('info', f"Column {col_letter} has custom filters: {custom_filters}")
                                    self.active_filters.append((col_letter, custom_filters))

                                if filter_col.find('.//a:dynamicFilter', namespaces) is not None:
                                    dynamic_filter = filter_col.find('.//a:dynamicFilter', namespaces).get('type')
                                    self.logger.log('info', f"Column {col_letter} has dynamic filter: {dynamic_filter}")
                                    self.active_filters.append((col_letter, dynamic_filter))

                                if filter_col.find('.//a:top10', namespaces) is not None:
                                    top10_filter = filter_col.find('.//a:top10', namespaces)
                                    top10_details = {
                                        "Top": top10_filter.get('top'),
                                        "Percent": top10_filter.get('percent'),
                                        "Value": top10_filter.get('val')
                                    }
                                    self.logger.log('info', f"Column {col_letter} has top10 filter: {top10_details}")
                                    self.active_filters.append((col_letter, top10_details))

                                if filter_col.find('.//a:colorFilter', namespaces) is not None:
                                    self.logger.log('info', f"Column {col_letter} has color filter")
                                    self.active_filters.append((col_letter, 'colorFilter'))
                        else:
                            self.logger.log('info', 'No filters detected in autoFilter XML')
            else:
                self.logger.log('info', 'No filter range detected')
        except Exception as e:
            self.logger.log('error', f"Error checking filters in Excel file: {e}")

    def escape_special_chars(self, pattern):
        """Escape special characters in the pattern for regex."""
        return re.escape(pattern)

    def apply_filters(self, df):
        """Apply active filters to the DataFrame."""
        self.logger.log('info', f"Applying filters to the DataFrame: {self.active_filters}")
        for col_letter, filter_values in self.active_filters:
            col_name = df.columns[column_index_from_string(col_letter) - 1]
            if not filter_values:  # Handle empty filter value to filter rows with empty cells
                df = df[df[col_name].isna() | (df[col_name] == '')]
            elif isinstance(filter_values, list) and all(isinstance(fv, str) for fv in filter_values):
                df = df[df[col_name].isin(filter_values)]
            elif isinstance(filter_values, list) and all(isinstance(fv, tuple) for fv in filter_values):
                for operator, value in filter_values:
                    if operator == 'notEqual':
                        if "*-*" in value:
                            self.logger.log('info', f"Applying removal for values containing: {value}")
                            escaped_value = self.escape_special_chars(value.replace("*-*", "-"))
                            df = df[~df[col_name].str.contains(escaped_value, na=False)]
                        else:
                            escaped_value = self.escape_special_chars(value)
                            df = df[~df[col_name].str.contains(escaped_value, na=False)]
                    # Implement other custom filter operators as necessary
            elif isinstance(filter_values, str):
                if filter_values == 'colorFilter':
                    self.logger.log('info', f"Skipping color filter on column {col_name}")
                    # Implement color filter if necessary
        return df

    def collect_data(self):
        """Collect data from the configured Excel file."""
        file_path = self.config.LOCAL_EXCEL_FILE
        sheet_name = self.config.EXCEL_SHEET_NAME
        try:
            self.get_active_filters(file_path, sheet_name)  # Log if filters exist
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=self.config.HEADER_ROW - 1)
            df = self.apply_filters(df)  # Apply filters to the DataFrame
            return df
        except Exception as e:
            self.logger.log('error', f"Error collecting data from local Excel file: {e}")
            raise
