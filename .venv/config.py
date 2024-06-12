# Configuration file to manage constants and configuration settings

USE_TEAMS_EXCEL = True  # Switch to toggle between local and Teams Excel

LOCAL_EXCEL_FILE = 'residents.xlsx'
EXCEL_SHEET_NAME = 'Sheet1'  # Sheet name to use
ADDRESS_COLUMN = 'address'
NAME_COLUMN = 'name'
WORK_ORDER_COLUMN = 'work_order_number'
LETTER_1_COLUMN = '1st_letter'
LETTER_2_COLUMN = '2nd_letter'
LETTER_3_COLUMN = '3rd_letter'

TEMPLATES_DIR = 'templates'
PRINT_SERVER_DIR = 'print_server'
WATCHER_INTERVAL = 604800  # 7 days in seconds

# Microsoft Graph API configuration
TENANT_ID = 'your-tenant-id'
CLIENT_ID = 'your-client-id'
CLIENT_SECRET = 'your-client-secret'
EXCEL_FILE_ID = 'your-excel-file-id'
EXCEL_FILE_DRIVE = 'your-excel-file-drive'  # Usually "drive-id" or "groups/group-id"
