# main.py

import os
import pandas as pd
import time

# Módulos locais
import config
import sharepoint_utils as sp_graph # Renomeado para clareza
import excel_utils as excel
import planner_logic as pl

# Módulos do projeto (mantidos)
from mb52 import updateStock
from montar_patan_logic import montar_patan

def clear_screen():
    """Limpa a tela do terminal."""
    os.system("cls" if os.name == "nt" else "clear")

def montar_patan_menu():
    """Gerencia a interface para o processo de montagem de Patan."""
    clear_screen()
    print("\n--- Montar Patan ---")
    letra_patan = input("Informe a letra patan (A até D): ").upper()
    linha = input("Selecione a linha (1, 2 ou 3): ")
    turno = input("Informe o Turno (1, 2 ou 3): ")

    clear_screen()
    print("ATUALIZANDO O ESTOQUE - AGUARDE UM MOMENTO!")
    updateStock()
    clear_screen()
    print("ATUALIZANDO O ARQUIVO PRINCIPAL")
    excel.updateExcel(config.PATAN_FILE_PATH)
    
    print(f"\nProcessando para Patan {letra_patan}, Linha {linha}, Turno {turno}...")
    
    df_result, df_diario, df_erros = montar_patan(letra_patan, linha, turno, config.PATAN_FILE_PATH)

    # Salva os resultados
    output_filename = f"{config.EXCEL_OUTPUT_PATH}/TEMPORARIO---PATAN-{letra_patan}-LINHA-{linha}-TURNO{turno}.xlsx"
    df_result.to_excel(output_filename, index=False)
    
    print(f"\nResultados salvos em {output_filename}")
    return df_result, linha

def main_loop():
    """Loop principal que escuta solicitações do SharePoint e processa."""
    while True:
        print("\n--- INICIANDO ROTINA DE RECEBIMENTO (GRAPH API) ---")
        
        # --- ALTERAÇÃO AQUI ---
        # 1. Chamar a nova função do Graph para receber dados
        df_solicitacao, site_id, list_id_receive = sp_graph.receive_data_from_sharepoint_graph(config.PLANNER_RECEIVE_LIST_NAME)
        
        if not df_solicitacao.empty:
            print(df_solicitacao.columns.tolist())
            print(df_solicitacao)
            print(df_solicitacao.shape)
            print("\nSolicitação recebida do SharePoint:")
            print(df_solicitacao.head().to_string())
            
            try:
                # Carregue o DataFrame de um ficheiro de exemplo ou execute a lógica para gerá-lo
                #df_original = pd.read_excel("output_patan.xlsx") 
                df_original, df_diario, df_erros = montar_patan(df_solicitacao.at[0,'patan'], df_solicitacao.at[0,'linha'], df_solicitacao.at[0,'turno'], config.PATAN_FILE_PATH) 

                df_planner_final = pl.create_worksheet_planner_reformulated(df_original, "1")
                print("\nPlanner final gerado:")
                print(df_planner_final.head(10).to_string())
                
                # --- ALTERAÇÃO AQUI ---
                # 2. Chamar a nova função do Graph para enviar dados
                sp_graph.send_data_to_sharepoint_graph(df_planner_final, config.PLANNER_SEND_LIST_NAME)
                
                # --- ALTERAÇÃO AQUI ---
                # 3. Chamar a nova função para atualizar o status, passando os IDs corretos
                id_solicitacao = df_solicitacao.iloc[0]['ID']
                sp_graph.update_item_status_graph(site_id, list_id_receive, id_solicitacao)
                
                df_planner_final.to_excel(config.PLANNER_FINAL_OUTPUT, index=False)
                print(f"\nArquivo '{config.PLANNER_FINAL_OUTPUT}' salvo com sucesso!")

            except FileNotFoundError:
                print("ERRO: O arquivo 'output_patan.xlsx' não foi encontrado para o processamento.")
            except Exception as e:
                print(f"Ocorreu um erro no processamento: {e}")

        else:
            print("Nenhuma nova solicitação. Aguardando...")
            time.sleep(15) # Pausa para não sobrecarregar o sistema

if __name__ == "__main__":
    main_loop()