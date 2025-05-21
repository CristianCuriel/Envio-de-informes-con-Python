#Librerias
from google.oauth2.credentials import credentials
from google.oauth2 import service_account
from googleapiclient.discovery import build
from config.setting import SCOPES, KEY, SPREADSHEET_ID, RANGE_NAME

class GoogleSheet:
    def __init__(self):
        self.key = KEY
        self.spreadsheet_id = SPREADSHEET_ID
        self.range_name = RANGE_NAME
        self.scopes = SCOPES    
        self.service = self._authenticate()
        
    def _authenticate(self):
        creds = service_account.Credentials.from_service_account_file(
            self.key,
            scopes=self.scopes
        )
        return build('sheets', 'v4', credentials=creds).spreadsheets()

    def get_values(self):
        result = self.service.values().get(
            spreadsheetId=self.spreadsheet_id,
            range=self.range_name
        ).execute()

        return result.get('values', [])