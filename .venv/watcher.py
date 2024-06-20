import requests
from msal import ConfidentialClientApplication
from data_collection import DataCollector
from custom_logging import Logger

class TeamsExcelWatcher:
    def __init__(self, data_collector, logger, config):
        self.data_collector = data_collector
        self.logger = logger
        self.config = config
        self.app = ConfidentialClientApplication(
            config['CLIENT_ID'],
            authority=f"https://login.microsoftonline.com/{config['TENANT_ID']}",
            client_credential=config['CLIENT_SECRET']
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
        url = f"https://graph.microsoft.com/v1.0/drives/{self.config['EXCEL_FILE_DRIVE']}/items/{self.config['EXCEL_FILE_ID']}/workbook/worksheets"
        headers = {
            "Authorization": f"Bearer {self.token}"
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            worksheets = response.json().get('value', [])
            worksheet_id = next(ws['id'] for ws in worksheets if ws['name'] == self.config['EXCEL_SHEET_NAME'])
            sheet_url = f"https://graph.microsoft.com/v1.0/drives/{self.config['EXCEL_FILE_DRIVE']}/items/{self.config['EXCEL_FILE_ID']}/workbook/worksheets/{worksheet_id}/usedRange"
            sheet_response = requests.get(sheet_url, headers=headers)
            if sheet_response.status_code == 200:
                return sheet_response.json()
        else:
            self.logger.log('error', f"Error getting worksheets: {response.status_code} {response.text}")
        return None

def load_defaults():
    with open('default_config.json', 'r') as f:
        return json.load(f)

if __name__ == "__main__":
    config = load_defaults()
    logger = Logger()
    data_collector = DataCollector(logger, config)
    watcher = TeamsExcelWatcher(data_collector, logger, config)
    excel_data = watcher.get_excel_data()
    if excel_data:
        df = data_collector.parse_excel_data(excel_data)
        # Continue with processing logic
