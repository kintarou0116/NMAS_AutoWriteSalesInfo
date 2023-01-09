from __future__ import print_function
import pickle
import os.path

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient import errors

from ErrorInfo import GoogleApiProcError

class GoogleApiManage(object):
    # 定数宣言
    API_SERVICE_GMAIL_NAME = 'gmail'
    API_GMAIL_VERSION = 'v1'
    API_SERVICE_SPREADSHEETS_NAME = 'sheets'
    API_SPREADSHEETS_VERSION = 'v4'
    GMAIL_SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
    SPREADSHEETS_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    GMAIL_PIKCLE_FILE_NAME = '../Tokens/token_gmail.pickle'
    SHEETS_PIKCLE_FILE_NAME = '../Tokens/token_sheets.pickle'
    CLIENT_SECRET_FILE = '../JSON/client_id.json'
    UPDATE_SHEETS_ID = '1VYetNdD_sYur82X_VIeDJ-wkEDfV9QsIFJ24wqep2Fc'
    #UPDATE_SHEETS_ID = '1HRHzm7tumzk5C4PP7CtmnSJ-Q-OPBJ8YUdMyG2G64wQ'
    #UPDATE_SHEETS_ID = '1AI9k5_bV1H91X20pN5tn6F_fcEdkPKupKOsKNwAaihA'
    #UPDATE_SHEETS_ID = '13pM1Jemk-deUVC38wiN8SCMc2WgUUW3H7graunQH_DM'
    
    # ローカルAPIサービスインスタンス
    gmail = None
    sheets = None
    
    def __init__(self):
        try:
            # GMailのAPIサービスインスタンスを取得（認証済み）
            self.gmail = self.get_authenticated_service(self.API_SERVICE_GMAIL_NAME, self.API_GMAIL_VERSION, self.GMAIL_PIKCLE_FILE_NAME, self.GMAIL_SCOPES)
            
            # SpreadSheetsのAPIサービスインスタンスを取得（認証済み）
            self.sheets = self.get_authenticated_service(self.API_SERVICE_SPREADSHEETS_NAME, self.API_SPREADSHEETS_VERSION, self.SHEETS_PIKCLE_FILE_NAME, self.SPREADSHEETS_SCOPES)
        except errors.HttpError as err:
            print(err.error_details)
            raise GoogleApiProcError
    
    # サービス接続（認証含む）
    # 引数
    #   apiServiceName  :接続先APIサービス名
    #   apiVersion      :APIバージョン数
    # 戻り値
    #   googleapiclient.discovery
    def get_authenticated_service(self, apiServiceName, apiVersion, tokenFileNeme, serviceScope):
        creds = None

        if os.path.exists(tokenFileNeme):
            with open(tokenFileNeme, 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except:
                    flow = InstalledAppFlow.from_client_secrets_file(self.CLIENT_SECRET_FILE, serviceScope)
                    creds = flow.run_local_server(port=0)
            else:
                flow = InstalledAppFlow.from_client_secrets_file(self.CLIENT_SECRET_FILE, serviceScope)
                creds = flow.run_local_server(port=0)
            with open(tokenFileNeme, 'wb') as token:
                pickle.dump(creds, token)
        
        return build(apiServiceName, apiVersion, credentials=creds)
    
import sys
sys.modules[__name__] = GoogleApiManage()