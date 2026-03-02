from fileinput import filename
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from googleapiclient.http import MediaIoBaseDownload
from io import BytesIO
import json

# Defina o escopo de acesso
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# Autenticação com o arquivo de credenciais da conta de serviço
credentials = service_account.Credentials.from_service_account_file(
    'conta-de-servico-google.json', scopes=SCOPES)

# Construir o serviço do Google Drive
service = build('drive', 'v3', credentials=credentials)

def update_json(cloud_json, local_json):
# primeiro json vai ser colocado atrás dos dados do segundo
    def recursive_update(source, target):
        for key, value in source.items():
            if isinstance(value, dict):
                # Se o valor for um dicionário, faça a atualização recursiva
                if key not in target:
                    target[key] = {}
                recursive_update(value, target[key])
            elif isinstance(value, list):
                # Se o valor for uma lista, faça a fusão preservando todos os itens
                if key not in target:
                    target[key] = []
                
                # Criar uma lista temporária para armazenar os itens já existentes e novos
                updated_list = {item['name']: item for item in target[key] if isinstance(item, dict)}
                
                # Atualizar ou adicionar os itens da lista de origem, preservando a ordem
                for item in value:
                    if isinstance(item, dict) and 'name' in item:
                        # Adiciona ou atualiza o item, mas no final
                        if item['name'] not in updated_list:
                            target[key].append(item)  # Adiciona novos itens ao final
                        else:
                            # Atualiza se o item já existir
                            updated_list[item['name']] = item
                    else:
                        # Se o item não for um dicionário, apenas adicione
                        target[key].append(item)
            else:
                # Para outros tipos de valor, substitua diretamente
                target[key] = value

    # Chame a função recursiva para atualizar o JSON local com os dados da nuvem
    recursive_update(cloud_json, local_json)
    
    return local_json

def get_json_information(file_id):
    try:
        # Solicitar o arquivo do Google Drive
        request = service.files().get_media(fileId=file_id)
        
        # Ler o conteúdo do arquivo em um buffer de bytes
        file_stream = BytesIO()
        downloader = MediaIoBaseDownload(file_stream, request)

        done = False
        while not done:
            status, done = downloader.next_chunk()

        # Mover o cursor do buffer para o início
        file_stream.seek(0)
        
        # Ler o conteúdo como string e converter para JSON
        file_content = file_stream.read().decode('utf-8')
        json_data = json.loads(file_content)
        
        return json_data

    except HttpError as error:
        print(f'An error occurred: {error}')
        return None

def upload_to_drive(file_name, file_path, folder_id=None):
    
    def share_file(file_id, user_email=None):
        if user_email:
            permission = {
                'type': 'user',
                'role': 'writer',
                'emailAddress': user_email
            }
        else:
            permission = {
                'type': 'anyone',
                'role': 'reader'
            }
            
        service.permissions().create(
            fileId=file_id,
            body=permission
        ).execute()
    
    def search_file_by_name(file_name):
        try:
            results = service.files().list(
                q=f"name='{file_name}'",
                spaces='drive',
                fields='files(id, name)').execute()

            items = results.get('files', [])

            if not items:
                print(f'No files found with name: {file_name}')
                return None
            else:
                for item in items:
                    print(f"Found file: {item['name']} with ID: {item['id']}")
                    return item['id']  # Retorna o ID do primeiro arquivo encontrado com esse nome
        
        except HttpError as error:
            print(f'An error occurred: {error}')

    def delete_file(file_id):
        try:
            service.files().delete(fileId=file_id).execute()
            print(f'File with ID {file_id} deleted successfully.')
        except HttpError as error:
            print(f'An error occurred: {error}')
            if error.resp.status == 404:
                print(f'File not found: {file_id}')

    # Buscar e apagar o arquivo antigo
    old_file_id = search_file_by_name(file_name)

    js_remote = get_json_information(old_file_id)

    # Verificando o tipo de 'js_remote' após a chamada da função
    print(f"Tipo de 'js_remote': {type(js_remote)}")

    '''
    BUG ARQUIVO ENORME DEMAIS PRA CARREGAR COMO JSON
    '''
    
    # Abrir o arquivo local para leitura
    # with open(file_path, 'r') as f:
    #     js_local = json.loads(f.read())  # Carrega o conteúdo local do arquivo JSON
        # print(f"Tipo de 'js_local': {type(js_local)}")

    # JS_LOCAL DOES NOT EXIST 

    # Realizar o merge dos dois JSONs
    obj = update_json(js_local, js_remote)

    # Abrir o arquivo local para escrita e salvar as alterações
    with open(file_path, 'w') as f:
        json.dump(obj, f, indent=4, ensure_ascii=False)

    # Reabrir o arquivo para leitura e imprimir o conteúdo final
    with open(file_path, 'r') as f:
        print(f.read())

    exit()

    if old_file_id:
        delete_file(old_file_id)

    # Fazer o upload do novo arquivo
    file_metadata = {
        'name': file_name,
    }
    
    # Se você quiser enviar o arquivo para uma pasta específica
    if folder_id:
        file_metadata['parents'] = [folder_id]

    media = MediaFileUpload(file_path, resumable=True)

    # Executando o upload
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()

    # Compartilhar o arquivo, se necessário
    share_file(file.get("id"), 'someemail@gmail.com')
    print(f'File ID: {file.get("id")}')

# Exemplo de uso:
if __name__ == '__main__':
    upload_to_drive('database-avec.json', './database-avec.json')
    # js = get_json_information('1vV2IWXjb7eR8YqXmJOAGTxU5CNg2xbmj')
    # print(js)