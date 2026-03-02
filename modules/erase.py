# -*- coding: utf-8 -*- 

import os
import gspread
from gspread.worksheet import Worksheet
from gspread.spreadsheet import Spreadsheet
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
import json
from time import sleep
SCOPES = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

def login():
    credentials = Credentials.from_service_account_file('avec.json', scopes=SCOPES)
    gc = gspread.auth.authorize(credentials)
    drive_service = build('drive', 'v3', credentials=credentials)
    return credentials, gc, drive_service

def get_sheet_file(gc, id):
    try:
        spreadsheet = gc.open_by_key(id)
        return spreadsheet
    except Exception as e:
        print(f'Error opening spreadsheet: {str(e)}')

def delete_sheet(drive_service, id):
    """
    Deleta um arquivo do Google Drive usando a ID.

    Args:
        id: ID do arquivo no Google Drive.
    """
    try:
        # Exclui o arquivo
        drive_service.files().delete(fileId=id).execute()
        print(f'Arquivo com ID "{id}" excluído com sucesso.')
    except Exception as e:
        print(f'Erro ao excluir o arquivo: {str(e)}')

def delete_all_sheets():
    credentials, gc, drive_service = login()
    sheet_files = os.listdir('planilhas')
    sheet_files = [file for file in sheet_files if file.endswith('.xlsx')]
    # already_seen = ['Financeiro.xlsx','Auditoria.xlsx','Produtos.xlsx']
    # already_seen = []
    with open('settings.json', 'r') as file:
        settings = json.load(file)
        for key, value in settings['excel_files'].items():
            ID = value.replace('https://docs.google.com/spreadsheets/d/', '')
            spreadsheet = get_sheet_file(gc, ID)
            delete_sheet(drive_service, ID)

def delete_one_tab(spreadsheet, worksheet):
    try:
        spreadsheet.del_worksheet(worksheet)
        print(f'Aba "{worksheet.title}" do arquivo {spreadsheet.title} deletada com sucesso.')
    except Exception as e:
        print(f'Erro ao deletar aba: {str(e)}')

def delete_all_tabs(spreadsheet):
    """
    Deleta todas as abas (worksheets) de um arquivo do Google Sheets remoto.
    
    Args:
        spreadsheet: Objeto Spreadsheet do gspread, que representa o arquivo remoto.
    """
    try:
        # Lista todas as abas no arquivo
        worksheets = spreadsheet.worksheets()
        lenght = len(worksheets)
        for i, worksheet in enumerate(worksheets):
            print(i)
            if i+1 == lenght:
                spreadsheet.add_worksheet('Sheet', rows=100, cols=20)
                delete_one_tab(spreadsheet, worksheet) 
            delete_one_tab(spreadsheet, worksheet)
            sleep(3)  # Evita o rate limit de requests do Google Sheets, que é 400 requests por minuto.
    except Exception as e:
        print(f'Erro ao deletar abas: {str(e)}')


if __name__ == '__main__':
    credentials, gc, drive_service = login()
    # sheet_files = os.listdir('planilhas')
    # sheet_files = [file for file in sheet_files if file.endswith('.xlsx')]
    # already_seen = ['Financeiro.xlsx','Auditoria.xlsx','Produtos.xlsx']
    already_seen = []
    '''
    Não alterar instruções seguintes porque o app.py executa diretamente o if name == __main__
    '''
    with open('settings.json', 'r') as file:
        settings = json.load(file)
        for key, value in settings['excel_files'].items():
            ID = value.replace('https://docs.google.com/spreadsheets/d/', '')
            spreadsheet = get_sheet_file(gc, ID)
            delete_all_tabs(spreadsheet)