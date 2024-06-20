import time
import requests
from msal import ConfidentialClientApplication
import config
from data_collection import DataCollector
from custom_logging import Logger

class TeamsExcelWatcher:
    def __init__(self, data_collector, logger, config):
        self.data_collector = data_collector
        self.interval = config['WATCHER_INTERVAL']
        self.excel_file_id = config['EXCEL_FILE_ID']
        self.excel_file_drive = config['EXCEL_FILE_DRIVE']
        self.logger = logger

        self.app = ConfidentialClientApplication(
            config['CLIENT_ID'],
            authority=f"https://login.microsoftonline.com/{config['TENANT_ID']}",
            client_credential=config['CLIENT_SECRET']
        )

        self.token = self.acquire_token()

    def acquire_token(self):
        self.logger.log('info', 'Acquiring token...')
        result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" in result:
            self.logger.log('info', 'Successfully acquired token')
            return result['access_token']
        else:
            error_message = result.get('error_description', 'Unknown error occurred while acquiring token')
            self.logger.log('error', f"Could not acquire token: {error_message}")
            raise Exception(f"Could not acquire token: {error_message}")

    def get_excel_data(self):
        self.logger.log('info', 'Getting Excel data from Teams')
        url = f"https://graph.microsoft.com/v1.0/drives/{self.excel_file_drive}/items/{self.excel_file_id}/workbook/worksheets"
        headers = {
            "Authorization": f"Bearer {self.token}"
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            worksheets = response.json().get('value', [])
            self.logger.log('info', f"Worksheets: {worksheets}")
            worksheet_id = next((ws['id'] for ws in worksheets if ws['name'] == config.config['EXCEL_SHEET_NAME']), None)
            if worksheet_id:
                sheet_url = f"https://graph.microsoft.com/v1.0/drives/{self.excel_file_drive}/items/{self.excel_file_id}/workbook/worksheets/{worksheet_id}/usedRange"
                sheet_response = requests.get(sheet_url, headers=headers)
                if sheet_response.status_code == 200:
                    data = sheet_response.json()
                    self.logger.log('info', f"Sheet data: {json.dumps(data, indent=2)}")
                    return data
                else:
                    self.logger.log('error', f"Error getting used range: {sheet_response.status_code} {sheet_response.text}")
            else:
                self.logger.log('error', f"Worksheet with name {config.config['EXCEL_SHEET_NAME']} not found.")
        else:
            self.logger.log('error', f"Error getting worksheets: {response.status_code} {response.text}")
        return None

    def watch_for_changes(self):
        while True:
            if config.config['USE_TEAMS_EXCEL']:
                self.logger.log('info', 'Watching for changes in Teams-stored Excel sheet')
                excel_data = self.get_excel_data()
                if excel_data:
                    data = self.data_collector.parse_excel_data(excel_data)
            else:
                self.logger.log('info', 'Watching for changes in local Excel file')
                data = self.data_collector.collect_data()

            if data:
                # Call your main workflow with updated data here
                pass

            time.sleep(self.interval)  # Wait for the next check
