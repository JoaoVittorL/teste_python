# PARTE RESPONSÁVEL PELA INTERAÇÃO COM O GOOGLE SHEETS E PYTHON!

#-------------------BIBLIOTECAS DO SHEETS-------------------↓
from __future__ import print_function                      #|
import os.path                                             #|
from google.auth.transport.requests import Request         #|
from google.oauth2.credentials import Credentials          #|
from google_auth_oauthlib.flow import InstalledAppFlow     #|
from googleapiclient.discovery import build                #|
from googleapiclient.errors import HttpError               #|
#-----------------------------------------------------------↑
#--------------IMPORT'S AVULSOS--------------↓
import time                                 #|
from datetime import datetime               #|
import os                                   #|
import pandas as pd                         #|
#--------------------------------------------↑

class Sheets_Python():
    def __init__(self):
        # VARIÁVEIS GLOBAIS
        self.token_path = os.path.join(os.getcwd(),'assets\\token.json')
        self.credential_path = os.path.join(os.getcwd(),'assets\\credentials.json')
        self.SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        self.creds = None
        
    def coleta_dados_sheets(self,id_sheets,range_sheets): # O RETORNO DA FUNÇÃO É UM DATAFRAME DOS VALORES CONTIDOS NO ID E RANGE INSERIDOS
        while True:
            try:
                if os.path.exists(self.token_path):
                    self.creds = Credentials.from_authorized_user_file(self.token_path, self.SCOPES)
                # Se as credenciais não forem válidas, aí será necessário relogar.
                if not self.creds or not self.creds.valid:
                    if self.creds and self.creds.expired and self.creds.refresh_token:
                        self.creds.refresh(Request())
                    else:
                        
                        flow = InstalledAppFlow.from_client_secrets_file(
                            self.credential_path, self.SCOPES)

                        self.creds = flow.run_local_server(port=0)
                        

                    with open(self.token_path, 'w') as token:
                        token.write(self.creds.to_json())
                break
            except:
                self.creds = None
                if os.path.exists(self.token_path):
                    os.remove(self.token_path)
            
        try:
            service = build('sheets', 'v4', credentials=self.creds)

            # Call the Sheets API
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=id_sheets,
                                        range=range_sheets).execute()
            
            values = result.get('values', [])
            dados_sheets = pd.DataFrame(values)
            # print(dados_sheets)
            if not values:
                return None
            
            return dados_sheets
            

        except HttpError as err:
            raise f'Erro ao acessar a api do sheets, segundo python o erro foi:\n>>{err}<<\n\n'
        
        
        
    def cola_dados_sheets(self, data_frame,id_sheets,range_sheets, clear=False,range_clear=None, append=False, append_col_ref=None):
        
        data_frame = data_frame.fillna('')
        cola = data_frame.values.tolist()
        
        if not cola:
            return 'Dataframe fornecido está vazio.\n'
        
        if os.path.exists(self.token_path):
            self.creds = Credentials.from_authorized_user_file(self.token_path, self.SCOPES)
        # Se as credenciais não forem válidas, aí será necessário relogar.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credential_path, self.SCOPES)

                self.creds = flow.run_local_server(port=0)
                

            with open(self.token_path, 'w') as token:
                token.write(self.creds.to_json())

        
        try:
            if (append==True and append_col_ref!=None):
                range_aux = range_sheets.split('!')
                range_ref = range_aux[0]
                range_ref = f"{range_ref}!{append_col_ref}1:{append_col_ref}"
                data_from_ref_col = Sheets_Python.coleta_dados_sheets(self,id_sheets=id_sheets, range_sheets=range_ref)
                last_row = len(data_from_ref_col) + 1
                range_aux2 = range_aux[1].split(':')
                new_sheets_range = f"{range_aux[0]}!{range_aux2[0]}{last_row}:{range_aux2[1]}"
                range_sheets = new_sheets_range
            
            service = build('sheets', 'v4', credentials=self.creds)

            # Call the Sheets API
            sheet = service.spreadsheets()
            
            if (clear==True and range_clear != None):
                clear_result = sheet.values().clear(spreadsheetId=id_sheets,
                                            range=range_clear).execute()
            
            result = sheet.values().update(spreadsheetId = id_sheets,
                                range= range_sheets, valueInputOption = 'USER_ENTERED',
                                body = {"values": cola}).execute()
            
            return True
        except HttpError as err:
            raise f'Erro ao acessar a api do sheets, segundo python o erro foi:\n>>{err}<<\n\n'
            

#----------------------------TESTE DE BLOCOS----------------------------↓
# if __name__ == '__main__':                                            
#     data = {'A': [111, 211],
#         'B': [411, 511],
#         'C': [711, 811]}

#     df = pd.DataFrame(data)     
    
#     start = Sheets_Python()                                           
#     resp = start.coleta_dados_sheets(id_sheets='1-U6qveKwm9xqwylMq8ibumTo45SOoGBs8YUh82H6oqA',range_sheets='Página2!A:C')                                
#     print("TESTE 1")
#     print (resp)
#     resp = start.cola_dados_sheets(id_sheets='1-U6qveKwm9xqwylMq8ibumTo45SOoGBs8YUh82H6oqA',range_sheets='Página2!D:F',data_frame=df,clear=True,range_clear='Página2!D:F')                            
#     print("TESTE 2")
#     print (resp)
#     resp = start.cola_dados_sheets(id_sheets='1-U6qveKwm9xqwylMq8ibumTo45SOoGBs8YUh82H6oqA',range_sheets='Página2!D:F',data_frame=df,append=True,append_col_ref="D")                            
#     print("TESTE 3")
#     print (resp)
#-----------------------------------------------------------------------↑