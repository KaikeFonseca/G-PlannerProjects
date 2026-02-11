# teste_graph_api.py

import pandas as pd
import sys

# Importar os módulos necessários
try:
    import config
    import sharepoint_utils as sp_graph
except ImportError:
    print("ERRO: Certifique-se de que os ficheiros 'config.py' e 'sharepoint_utils.py' estão na mesma pasta.")
    sys.exit(1)

def create_sample_dataframe():
    """Cria um DataFrame de exemplo para ser enviado ao SharePoint."""
    print("--> Criando DataFrame de exemplo...")
    sample_data = {
        'Material': ['L307000B6A', 'I28100SB6A'],
        'posto': ['B642047', 'B642047'],
        'patan': ['A', 'A'],
        'linha': ['2', '2'],
        'turno': [1, 1],
        'qtdPecasSeremProduzidas': [480, 924],
        'qtdPorKanban': [6, 12],
        'kanbans': [80, 77],
        'tempoProd': [128.08, 328.26],
        'sequencia': [1, 2],
        'prodEmLinha': [0, 0],
        'compComb': ['ComponenteA|1|DescA|1000', 'ComponenteB|2|DescB|2000'],
        'estoqueMaterial': [674, -156],
        'estoqueKanbanMax': [696, 768],
        'diff': [-22, -924],
        'obs': [None, 'Falta comp.'],
        'STATUS': [0, 0],
        'horaProdInicial': ['06:00', '08:08'],
        'horaProdFinal': ['08:08', '14:16'],
        'descricaoRefeicao': [None, 'Pausa para refeição'],
        'tags': ['PATAN (A);;;1° TURNO;;;PONT. 47', 'PATAN (A);;;1° TURNO;;;PONT. 47'],
        'descricao': ['INICIO: 06:00...', 'INICIO: 08:08...'],
        'checklist': ['1 - OK;2 - OK', '1 - OK;2 - OK']
    }
    df = pd.DataFrame(sample_data)
    print("    DataFrame de exemplo criado com sucesso.")
    return df

def run_tests():
    """Executa uma sequência de testes para validar a integração com o SharePoint via Graph API."""
    
    print("\n==================================================")
    print("== INICIANDO TESTES DE INTEGRAÇÃO COM SHAREPOINT (GRAPH API) ==")
    print("==================================================\n")

    # --- Teste 1: Receber Dados ---
    print("\n[TESTE 1/3] Recebendo dados da lista de solicitações...")
    df_solicitacao, site_id, list_id_receive = sp_graph.receive_data_from_sharepoint_graph(
        config.PLANNER_RECEIVE_LIST_NAME
    )

    if df_solicitacao is None or site_id is None:
        print("\n[FALHA NO TESTE 1] Não foi possível obter os dados. Verifique os erros acima.")
        return

    if df_solicitacao.empty:
        print("\n[AVISO NO TESTE 1] A função de recebimento funcionou, mas não encontrou itens com STATUS = 1.")
        print("   Para testar a atualização (Teste 3), adicione um item manualmente na lista 'list_solicitacao_planner_py' com o campo 'STATUS' igual a '1'.")
    else:
        print("\n[SUCESSO NO TESTE 1] DataFrame recebido:")
        print(df_solicitacao.head().to_string())

    # --- Teste 2: Enviar Dados ---
    print("\n[TESTE 2/3] Enviando dados para a lista do planner...")
    df_para_enviar = create_sample_dataframe()
    try:
        sp_graph.send_data_to_sharepoint_graph(df_para_enviar, config.PLANNER_SEND_LIST_NAME)
        print("\n[SUCESSO NO TESTE 2] Os dados de exemplo foram enviados para o SharePoint.")
    except Exception as e:
        print(f"\n[FALHA NO TESTE 2] Ocorreu um erro ao enviar os dados: {e}")
        return

    # --- Teste 3: Atualizar Status ---
    print("\n[TESTE 3/3] Atualizando o status de um item recebido...")
    if not df_solicitacao.empty:
        # Pega o ID do primeiro item da lista recebida
        item_id_para_atualizar = df_solicitacao.iloc[0]['ID']
        try:
            sp_graph.update_item_status_graph(site_id, list_id_receive, item_id_para_atualizar)
            print(f"\n[SUCESSO NO TESTE 3] A função de atualização para o item ID '{item_id_para_atualizar}' foi executada.")
        except Exception as e:
            print(f"\n[FALHA NO TESTE 3] Ocorreu um erro ao atualizar o status: {e}")
    else:
        print("\n[TESTE 3 IGNORADO] Não há itens para atualizar (nenhum item com STATUS = 1 foi encontrado no Teste 1).")

    print("\n==================================================")
    print("== TESTES FINALIZADOS ==")
    print("==================================================")


if __name__ == "__main__":
    run_tests()