# -*- coding: utf-8 -*- 

import os
import gspread
from gspread.worksheet import Worksheet
from gspread.spreadsheet import Spreadsheet
from google.oauth2.service_account import Credentials
import pandas as pd
import json
from time import sleep

SCOPES = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

def __get_id_by_name(gc, name:str):
    files = gc.list_spreadsheet_files()
    sheet_file_id = None
    for file in files:
        if file['name'] == name:
            sheet_file_id = file['id']
            break
    if sheet_file_id:
        return sheet_file_id
    else:
        return None

def get_sheet_file(gc, file_path):
    file = open('settings.json', 'r+', encoding='utf-8')
    settings = json.load(file)
    filename = f'{os.path.basename(file_path).replace(".xlsx","")}'
    try:
        # id_sheet = __get_id_by_name(gc, settings[filename])
        # spreadsheet = gc.open_by_key(id_sheet)
        id_sheet = settings['excel_files'][filename].replace('https://docs.google.com/spreadsheets/d/','')
        spreadsheet = gc.open_by_key(id_sheet)
    except KeyError:
        spreadsheet = gc.create(filename)
        settings[filename] = spreadsheet.id
        file.truncate(0)
        file.seek(0)
        json.dump(settings, file, indent=4, ensure_ascii=False)
    return spreadsheet
    
def login():
    credentials = Credentials.from_service_account_file('avec.json', scopes=SCOPES)
    gc = gspread.auth.authorize(credentials)
    return gc

def send_sheets():
    
    def convert_excel_to_csv(file_path):
        def create_folder_csv():
            folder_path = f'planilhas/{os.path.basename(file_path).replace(".xlsx", "")}/'
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            return folder_path

        data_excel = pd.ExcelFile(file_path)

        for sheet_name in data_excel.sheet_names:
            data_sheet = pd.read_excel(data_excel, sheet_name=sheet_name)
            folder_path = create_folder_csv()
            csv_file_path = os.path.join(folder_path, f'{sheet_name}.csv')
            data_sheet.to_csv(csv_file_path, index=False)
        return f'planilhas/{os.path.basename(file_path).replace(".xlsx", "")}/'
    
    def push_worksheets(spreadsheet:Spreadsheet, folder_path):
        print('atualizando planilhas')
        def get_or_create_worksheet(spreadsheet, sheet_title, num_rows, num_cols):
            try:
                worksheet = spreadsheet.worksheet(sheet_title)
            except gspread.exceptions.WorksheetNotFound:
                print(sheet_title)
                worksheet = spreadsheet.add_worksheet(title=sheet_title[:100], rows=num_rows, cols=num_cols)
            return worksheet
        
        def update_worksheets():
            def check_duplicates(ws:Worksheet):
                novos_dados = ws.get_all_values()
                df_novos_dados = pd.DataFrame(novos_dados[1:], columns=novos_dados[0])
                duplicatas = df_novos_dados[df_novos_dados.duplicated(keep=False)]
                if not duplicatas.empty:
                    print(f"Duplicatas encontradas na planilha: {ws.title}")
                    print(duplicatas)
                else:
                    print(f"Não há duplicatas na planilha: {ws.title}")

            def remove_duplicates(ws:Worksheet):
                data = ws.get_all_values()
                unique_lines = []
                lines_already_seen = set()
                for linha in data:
                    chave = tuple(linha)
                    if chave not in lines_already_seen:
                        unique_lines.append(linha)
                        lines_already_seen.add(chave)
                ws.clear()
                sleep(3)
                if len(unique_lines) == 0:
                    pass
                else:
                    ws.update(unique_lines)

            files = os.listdir(folder_path)
            for file in files:
                with open(f'{folder_path}/{file}', 'r') as file:
                    csv_contents = file.read().splitlines()
                    have_text = False
                    for line in csv_contents:
                        if not line.isspace() and have_text is False:
                            have_text = True
                            continue
                        else:
                            break
                    if have_text is False:
                        continue

                    folder = folder_path +'.xlsx'
                    folder = folder.strip()
                    if folder[-1] == '/' in folder:
                        folder = folder[:-1]
                    if '//' in folder:
                        folder = folder.replace('//','/')
                    with pd.ExcelFile(folder_path[:-1]+'.xlsx') as xls:
                        try:
                            data_sheet = pd.read_excel(xls, os.path.basename(file.name).replace('.csv',''))
                        except:
                            print(f'{file.name} não tem uma planilha?')
                            continue
                        num_rows, num_cols = data_sheet.shape
                ws = get_or_create_worksheet(spreadsheet, os.path.basename(file.name).replace('.csv',''), num_rows, num_cols)

                # update de valores
                try:
                    ws.update([line.split(',') for line in csv_contents])
                except Exception as e:
                    print(f'{e}: \n\n{csv_contents}\n\n')
                    continue
                # removendo qualquer row duplicata
                remove_duplicates(ws)
                # final check
                check_duplicates(ws)
            try:
                sheet_origin = spreadsheet.get_worksheet(0)
                if sheet_origin.title == 'Sheet' or sheet_origin.title == 'Página':
                    spreadsheet.del_worksheet(sheet_origin)
            except Exception as e:
                print(f'remove sheet problem: {e}')

        def share_with_users():
            emails_to_share = [
                'someemail@gmail.com',
            ]
            users_with_access = spreadsheet.list_permissions()
            users_to_share = []
            for email in emails_to_share:
                has_access = False
                for user in users_with_access:
                    if email == user['emailAddress']:
                        has_access = True
                        break
                if not has_access:
                    users_to_share.append(email)

            for user in users_to_share:
                spreadsheet.share(user, perm_type='user', role='writer')
        update_worksheets()
        # ERROR
        # KEY ERROR EmailAdress
        # share_with_users()
        print(f"Arquivo {spreadsheet.title} carregado com sucesso para o Google Planilhas!")

    gc = login()
    sheet_files = os.listdir('planilhas')
    sheet_files = [file for file in sheet_files if file.endswith('.xlsx')]
    # already_seen = ['Financeiro.xlsx','Auditoria.xlsx','Produtos.xlsx']
    already_seen = []
    for sheet_file in sheet_files:
        if sheet_file not in already_seen:
            file_path = f'planilhas/{sheet_file}'
            spreadsheet = get_sheet_file(gc, file_path)
            folder_path = convert_excel_to_csv(file_path)
            push_worksheets(spreadsheet, folder_path)

def change_values():
    def login():
        credentials = Credentials.from_service_account_file('avec.json', scopes=SCOPES)
        gc = gspread.auth.authorize(credentials)
        return gc

    def get_sheet_files():
        with open('settings.json', 'r', encoding='utf-8') as file:
            settings = json.load(file)
            sheet_files = []

            for key, value in settings['excel_files'].items():
                print(f'{key}\n')
                value = value.replace('https://docs.google.com/spreadsheets/d/', '')
                sheet_files.append(value)
            return sheet_files

    def alter_values(sheet_files):
        def is_digit(value: str) -> bool:
            value = value.replace(',', '.')
            try:
                float(value)
                return True
            except ValueError:
                return False

        def convert_value(value: str) -> float:
            if ',' in value:
                value = value.replace('.', '').replace(',', '.')
            else:
                value = value.replace('.', '')
            return float(value)

        for sheet_file in sheet_files:
            gc = login()
            spreadsheet = gc.open_by_key(sheet_file)
            worksheets = spreadsheet.worksheets()

            for worksheet in worksheets:
                if worksheet.title.lower() in ['sheet', 'página']:
                    continue
                
                # Obtém todos os valores
                all_cells = worksheet.get_all_values()
                modified_rows = []

                # Processa as células
                for row in all_cells:
                    new_row = []
                    for cell in row:
                        new_value = cell
                        if isinstance(new_value, str):
                            # Remove aspas se necessário
                            new_value = new_value.replace('"', '').replace("'", "")

                            # Converte valores numéricos
                            if is_digit(new_value):
                                new_value = convert_value(new_value)
                        new_row.append(new_value)
                    modified_rows.append(new_row)

                # Atualiza a planilha inteira com os novos valores
                try:
                    worksheet.update(modified_rows, 'A1')
                    print(f"Worksheet {worksheet.title} updated successfully.")
                    sleep(2)
                except Exception as e:
                    print(f"Failed to update worksheet {worksheet.title}: {e}")
                    sleep(10)
                    try:
                        worksheet.update(modified_rows, 'A1')
                        print(f"Retrying: Worksheet {worksheet.title} updated successfully.")
                        sleep(2)
                    except Exception as e:
                        print(f"Failed to update worksheet {worksheet.title} on retry: {e}")
                        sleep(10)
                    
    # Execução
    sheet_files = get_sheet_files()
    alter_values(sheet_files)

if __name__ == '__main__':
    with open('settings.json', 'r') as file:
        settings = json.load(file)
    send_sheets()
    change_values()