import sys
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Defina o escopo de acesso
SCOPES = ['https://www.googleapis.com/auth/drive']

# Autenticação com o arquivo de credenciais da conta de serviço
credentials = service_account.Credentials.from_service_account_file(
    'avec.json', scopes=SCOPES)

# Construir o serviço do Google Drive
service = build('drive', 'v3', credentials=credentials)

def get_sheet_ids():
    """Obtém os IDs das planilhas do arquivo settings.json."""
    with open('settings.json', 'r', encoding='utf-8') as file:
        settings = json.load(file)
        return [
            value.replace('https://docs.google.com/spreadsheets/d/', '').split('/')[0]
            for value in settings['excel_files'].values()
        ]

def grant_access(file_id, emails):
    """Concede acesso de escritor aos e-mails fornecidos para um arquivo do Google Drive."""
    for email in emails:
        try:
            permission = {
                'type': 'user',
                'role': 'writer',
                'emailAddress': email
            }
            service.permissions().create(
                fileId=file_id,
                body=permission,
                sendNotificationEmail=False
            ).execute()
            print(f'Acesso concedido a {email} para o arquivo {file_id}')
        except HttpError as error:
            print(f'Erro ao conceder acesso a {email}: {error}')

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python script.py email1@example.com email2@example.com ...")
        sys.exit(1)
    
    emails_to_share = sys.argv[1:]
    sheet_ids = get_sheet_ids()
    
    for sheet_id in sheet_ids:
        grant_access(sheet_id, emails_to_share)
