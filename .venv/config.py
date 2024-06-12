# Configuration file to manage constants and configuration settings

USE_TEAMS_EXCEL = True  # Switch to toggle between local and Teams Excel

LOCAL_EXCEL_FILE = 'residents.xlsx'
EXCEL_SHEET_NAME = 'Sheet1'  # Sheet name to use

# Column names in the Excel sheet
ADDRESS_COLUMN = 'address'
NAME_COLUMN = 'name'
WORK_ORDER_COLUMN = 'work_order_number'
LETTER_1_COLUMN = '1st_letter'
LETTER_2_COLUMN = '2nd_letter'
LETTER_3_COLUMN = '3rd_letter'
REVIEW_COLUMN = 'review 1'  # Column used to determine the template group

# Template group values
REVIEW_POSITIVE_VALUE = 'CertainValue'  # Replace with the actual value
TEMPLATE_GROUP1 = {
    'LETTER_1_TEMPLATE': 'template1_group1',
    'LETTER_2_TEMPLATE': 'template2_group1',
    'LETTER_3_TEMPLATE': 'template3_group1'
}
TEMPLATE_GROUP2 = {
    'LETTER_1_TEMPLATE': 'template1_group2',
    'LETTER_2_TEMPLATE': 'template2_group2',
    'LETTER_3_TEMPLATE': 'template3_group2'
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
