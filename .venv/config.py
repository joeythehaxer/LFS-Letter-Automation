# Configuration file to manage constants and configuration settings

USE_TEAMS_EXCEL = True  # Switch to toggle between local and Teams Excel

LOCAL_EXCEL_FILE = 'residents.xlsx'
EXCEL_SHEET_NAME = 'Sheet1'  # Sheet name to use
ADDRESS_COLUMN = 'ITEM LOCATION / ADDRESS'
NAME_COLUMN = 'name'
WORK_ORDER_COLUMN = 'PO number / Action Number'
LETTER_1_COLUMN = '1ST ACCESS LETTER DATE/CALL '
LETTER_2_COLUMN = '2ND ACCESS LETTER DATE/CALL'
LETTER_3_COLUMN = '3RD ACCESS LETTER DATE/CALL'
ORDER_STATUS_COLUMN = "TYPE OF WORKS"
TEMPLATES_DIR = 'templates'
PRINT_SERVER_DIR = 'print_server'
WATCHER_INTERVAL = 604800  # 7 days in seconds

# Microsoft Graph API configuration
TENANT_ID = 'your-tenant-id'
CLIENT_ID = 'your-client-id'
CLIENT_SECRET = 'your-client-secret'
EXCEL_FILE_ID = 'your-excel-file-id'
EXCEL_FILE_DRIVE = 'your-excel-file-drive'  # Usually "drive-id" or "groups/group-id"
