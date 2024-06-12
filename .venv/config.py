# Configuration file to manage constants and configuration settings

USE_GUI = False# Switch to toggle between GUI and command-line interface

USE_TEAMS_EXCEL = False  # Switch to toggle between local and Teams Excel

LOCAL_EXCEL_FILE = 'residents.xlsx'
EXCEL_SHEET_NAME = 'ABRI ONLY TRACKER  '  # Sheet name to use

# Column names in the Excel sheet
ADDRESS_COLUMN = 'ITEM LOCATION / ADDRESS'
NAME_COLUMN = 'Supplied Contact'
WORK_ORDER_COLUMN = 'PO number / Action Number'
LETTER_1_COLUMN = '1ST ACCESS LETTER DATE/CALL '
LETTER_2_COLUMN = '2ND ACCESS LETTER DATE/CALL'
LETTER_3_COLUMN = '3RD ACCESS LETTER DATE/CALL'
REVIEW_COLUMN = 'Review 1'  # Column used to determine the template group
NEW_FILTER_COLUMN = 'ORDER STATUS'  # New column for the additional filter

# Template group values
REVIEW_POSITIVE_VALUE = 'A NEW DOOR/S REQUIRED'  # Replace with the actual value
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

TEMPLATES_DIR = 'templates'
PRINT_SERVER_DIR = 'print_server'
WATCHER_INTERVAL = 604800  # 7 days in seconds

# Microsoft Graph API configuration
TENANT_ID = 'your-tenant-id'
CLIENT_ID = 'your-client-id'
CLIENT_SECRET = 'your-client-secret'
EXCEL_FILE_ID = 'your-excel-file-id'
EXCEL_FILE_DRIVE = 'your-excel-file-drive'  # Usually "drive-id" or "groups/group-id"
