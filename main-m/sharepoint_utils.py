# sharepoint_utils.py
import requests
import pandas as pd
from datetime import datetime, date
from msal import ConfidentialClientApplication

# Importar configurações
import config

# --- Configurações da API do Microsoft Graph ---
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
AUTHORITY_URL = f"https://login.microsoftonline.com/{config.TENANT_ID}"

# --- Cache para IDs (melhora o desempenho) ---
SITE_ID_CACHE = None
LIST_ID_CACHE = {}

def get_graph_token():
    """Obtém um token de acesso para a API do Microsoft Graph."""
    app = ConfidentialClientApplication(
        client_id=config.CLIENT_ID,
        authority=AUTHORITY_URL,
        client_credential=config.CLIENT_SECRET
    )
    token_response = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" in token_response:
        return token_response["access_token"]
    else:
        raise Exception(f"Falha ao obter token do Graph: {token_response.get('error_description')}")

def get_site_id(token):
    """Obtém e armazena em cache o ID do site do SharePoint."""
    global SITE_ID_CACHE
    if SITE_ID_CACHE:
        return SITE_ID_CACHE
    
    endpoint = f"https://graph.microsoft.com/v1.0/sites/{config.SHAREPOINT_HOSTNAME}:{config.SITE_PATH}"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(endpoint, headers=headers)
    response.raise_for_status()
    SITE_ID_CACHE = response.json().get('id')
    print(f"ID do Site '{config.SITE_PATH}' obtido: {SITE_ID_CACHE}")
    return SITE_ID_CACHE

def get_list_id(token, site_id, list_name):
    """Obtém e armazena em cache o ID de uma lista específica."""
    if list_name in LIST_ID_CACHE:
        return LIST_ID_CACHE[list_name]
        
    endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists?$filter=displayName eq '{list_name}'"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(endpoint, headers=headers)
    response.raise_for_status()
    data = response.json()
    if data.get('value'):
        list_id = data['value'][0]['id']
        LIST_ID_CACHE[list_name] = list_id
        print(f"ID da Lista '{list_name}' obtido: {list_id}")
        return list_id
    else:
        raise Exception(f"Lista '{list_name}' não encontrada no site.")

def receive_data_from_sharepoint_graph(list_name: str):
    """Recebe itens de uma lista do SharePoint com STATUS '1' usando a API do Graph."""
    try:
        token = get_graph_token()
        site_id = get_site_id(token)
        list_id = get_list_id(token, site_id, list_name)
        
        # Filtra por itens onde a coluna 'STATUS' é igual a '1'
        filter_query = "fields/STATUS eq '1'"
        endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?expand=fields&filter={filter_query}"
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly" # Necessário para alguns filtros
        }
        
        response = requests.get(endpoint, headers=headers)
        response.raise_for_status()
        items = response.json().get('value', [])
        
        if not items:
            print("Nenhum item encontrado com STATUS = 1.")
            return pd.DataFrame(), site_id, list_id

        # Extrai os campos de cada item para o DataFrame
        data = []
        for item in items:
            fields = item.get('fields', {})
            fields['ID'] = item.get('id') # Adiciona o ID do item da lista
            data.append(fields)

        return pd.DataFrame(data), site_id, list_id
    except Exception as e:
        print(f"Erro ao buscar dados do SharePoint via Graph: {e}")
        return pd.DataFrame(), None, None

def send_data_to_sharepoint_graph(df: pd.DataFrame, list_name: str):
    """Envia um DataFrame para uma lista do SharePoint usando a API do Graph."""
    try:
        token = get_graph_token()
        site_id = get_site_id(token)
        list_id = get_list_id(token, site_id, list_name)
        
        endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items"
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

        print(f"Enviando {len(df)} itens para a lista '{list_name}'...")
        for index, row in df.iterrows():
            item_properties = {}
            for df_col, sp_col in config.DF_TO_SP_MAP.items():
                if df_col in row and pd.notna(row[df_col]):
                    value = row[df_col]
                    # Formatação especial para datas
                    if isinstance(value, (datetime, pd.Timestamp, date)):
                        item_properties[sp_col] = value.isoformat()
                    else:
                        item_properties[sp_col] = str(value)
            
            payload = {"fields": item_properties}
            response = requests.post(endpoint, headers=headers, json=payload)
            response.raise_for_status()
            
        print("Envio de dados para o SharePoint (Graph) concluído com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar dados para o SharePoint via Graph: {e}")
        if 'response' in locals():
            print(f"Resposta do Servidor: {response.text}")

def update_item_status_graph(site_id: str, list_id: str, item_id: str):
    """Atualiza o campo 'STATUS' de um item para '2' usando a API do Graph."""
    print(f"\n\t--- ATUALIZANDO STATUS (GRAPH) DO ITEM ID {item_id} ---")
    try:
        token = get_graph_token()
        endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        
        payload = {"STATUS": "2"}
        
        response = requests.patch(endpoint, headers=headers, json=payload)
        response.raise_for_status()
        print(f"\tItem com ID {item_id} atualizado para STATUS = 2 com sucesso.")
    except Exception as e:
        print(f"\n\tErro ao atualizar STATUS via Graph para o item {item_id}.\n\tMOTIVO: {e}")
        if 'response' in locals():
            print(f"\tResposta do Servidor: {response.text}")