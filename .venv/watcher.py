import time
import requests
from msal import ConfidentialClientApplication
import config
from data_collection import DataCollector
from custom_logging import Logger

class TeamsExcelWatcher:
    def __init__(self, data_collector, logger, tenant_id, client_id, client_secret, excel_file_id, excel_file_drive, interval=config.WATCHER_INTERVAL):
        self.data_collector = data_collector
        self.interval = interval
        self.excel_file_id = excel_file_id
        self.excel_file_drive = excel_file_drive
        self.logger = logger

        self.app = ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret
        )

        self.token = self.acquire_token()

    def acquire_token(self):
        result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" in result:
            self.logger.log('info', 'Successfully acquired token')
            return result['access_token']
        else:
            self.logger.log('error', 'Could not acquire token')
            raise Exception("Could not acquire token.")

    def get_excel_data(self):
        self.logger.log('info', 'Getting Excel data from Teams')
        url = f"https://graph.microsoft.com/v1.0/drives/{self.excel_file_drive}/items/{self.excel_file_id}/workbook/worksheets"
        headers = {
            "Authorization": f"Bearer {self.token}"
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            worksheets = response.json().get('value', [])
            worksheet_id = next(ws['id'] for ws in worksheets if ws['name'] == config.EXCEL_SHEET_NAME)
            sheet_url = f"https://graph.microsoft.com/v1.0/drives/{self.excel_file_drive}/items/{self.excel_file_id}/workbook/worksheets/{worksheet_id}/usedRange"
            sheet_response = requests.get(sheet_url, headers=headers)
            if sheet_response.status_code == 200:
                return sheet_response.json()
        return None

    def watch_for_changes(self):
        while True:
            if config.USE_TEAMS_EXCEL:
                self.logger.log('info', 'Watching for changes in Teams-stored Excel sheet')
                excel_data = self.get_excel_data()
                if excel_data:
                    data = self.data_collector.get_resident_data(excel_data)
            else:
                self.logger.log('info', 'Watching for changes in local Excel file')
                data = self.data_collector.collect_data()

            if data:
                # Call your main workflow with updated data here
                pass

            time.sleep(self.interval)  # Wait for the next check
