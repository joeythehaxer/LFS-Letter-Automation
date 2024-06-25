import time
import requests
from msal import ConfidentialClientApplication
import config.settings as settings
from data_collection.data_collector import DataCollector
from custom_logging.logger import Logger

class TeamsExcelWatcher:
    def __init__(self, data_collector, logger, tenant_id, client_id, client_secret, excel_file_id, excel_file_drive):
        self.data_collector = data_collector
        self.logger = logger

        self.app = ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret
        )

        self.token = self.acquire_token()

    def acquire_token(self):
        retries = 3
        while retries > 0:
            result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            if "access_token" in result:
                self.logger.log('info', 'Successfully acquired token')
                return result['access_token']
            else:
                self.logger.log('warning', 'Failed to acquire token, retrying...')
                retries -= 1
                time.sleep(10)  # Wait before retrying
        self.logger.log('error', 'Could not acquire token after retries')
        raise Exception("Could not acquire token after retries.")

    def get_excel_data(self):
        self.logger.log('info', 'Getting Excel data from Teams')
        url = f"https://graph.microsoft.com/v1.0/drives/{self.config.EXCEL_FILE_DRIVE}/items/{self.config.EXCEL_FILE_ID}/workbook/worksheets"
        headers = {"Authorization": f"Bearer {self.token}"}
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                worksheets = response.json().get('value', [])
                worksheet_id = next(ws['id'] for ws in worksheets if ws['name'] == self.config.EXCEL_SHEET_NAME)
                sheet_url = f"https://graph.microsoft.com/v1.0/drives/{self.config.EXCEL_FILE_DRIVE}/items/{self.config.EXCEL_FILE_ID}/workbook/worksheets/{worksheet_id}/usedRange"
                sheet_response = requests.get(sheet_url, headers=headers)
                if sheet_response.status_code == 200:
                    return sheet_response.json()
                else:
                    self.logger.log('error', 'Failed to fetch worksheet data')
            else:
                self.logger.log('error', 'Failed to fetch worksheets list')
        except Exception as e:
            self.logger.log('error', f'Exception during Excel data fetch: {str(e)}')
        return None

    def watch_for_changes(self):
        while True:
            if self.config.USE_TEAMS_EXCEL:
                self.logger.log('info', 'Watching for changes in Teams-stored Excel sheet')
                excel_data = self.get_excel_data()
                if excel_data:
                    data = self.data_collector.get_resident_data(excel_data)  # Assuming implementation
            else:
                self.logger.log('info', 'Watching for changes in local Excel file')
                data = self.data_collector.collect_data()

            if data:
                pass  # Implement processing logic or calling workflow functions

            time.sleep(self.config.WATCHER_INTERVAL)  # Configurable interval
