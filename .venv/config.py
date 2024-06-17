# Configuration file to manage constants and configuration settings

import os

USE_GUI = True  # Switch to toggle between GUI and command-line interface

USE_TEAMS_EXCEL = False  # Switch to toggle between local and Teams Excel

LOCAL_EXCEL_FILE = 'residents.xlsx'
EXCEL_SHEET_NAME = 'ABRI ONLY TRACKER'  # Sheet name to use

# Column names in the Excel sheet
ADDRESS_COLUMN = 'Address'
NAME_COLUMN = 'Name'
WORK_ORDER_COLUMN = 'Work Order Number'
LETTER_1_COLUMN = '1ST ACCESS LETTER DATE/CALL'
LETTER_2_COLUMN = '2ND ACCESS LETTER DATE/CALL'
LETTER_3_COLUMN = '3RD ACCESS LETTER DATE/CALL'
REVIEW_COLUMN = 'Review 1'  # Column used to determine the template group

# Filter configuration: list of dictionaries with column names and values to filter by
FILTERS = [
    {'column': 'New Filter Column', 'value': 'Default Value'}
]

# Template group values
REVIEW_POSITIVE_VALUE = 'CertainValue'  # Replace with the actual value
TEMPLATE_GROUP1 = {
    'LETTER_1_TEMPLATE': 'template1',
    'LETTER_2_TEMPLATE': 'template2',
    'LETTER_3_TEMPLATE': 'template3'
}
TEMPLATE_GROUP2 = {
    'LETTER_1_TEMPLATE': 'template1',
    'LETTER_2_TEMPLATE': 'template2',
    'LETTER_3_TEMPLATE': 'template3'
}

# Placeholder configuration
PLACEHOLDERS = {
    'NAME_PLACEHOLDER': 'Name',
    'ADDRESS_PLACEHOLDER': 'Address',
    'EMAIL_PLACEHOLDER': 'Email',
    'WORK_ORDER_NUMBER_PLACEHOLDER': 'Work Order Number'
}

TEMPLATES_DIR = 'templates'
PRINT_SERVER_DIR = 'print_server'
WATCHER_INTERVAL = 604800  # 7 days in seconds

# Microsoft Graph API configuration
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
EXCEL_FILE_ID = os.getenv('EXCEL_FILE_ID')
EXCEL_FILE_DRIVE = os.getenv('EXCEL_FILE_DRIVE')  # Usually "drive-id" or "groups/group-id"
