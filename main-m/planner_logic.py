# planner_logic.py

import pandas as pd
from datetime import datetime, timedelta

def _build_tags(row):
    """Função auxiliar para construir a string de tags para uma linha."""
    tags_list = [
        f"PATAN ({row['patan']})",
        f"{row['turno']}° TURNO",
        f"PONT. {str(row['posto'])[-2:]}"
    ]
    if row.get('prodEmLinha') == 1:
        tags_list.append("PROD. EM LINHA")
    return ";;;".join(tags_list)

def _build_descricao(row):
    """Função auxiliar para construir a string de descrição para uma linha."""
    # 1. Tratamento para Produção em Linha
    if row.get('prodEmLinha') == 1:
        return "PRODUÇÃO EM LINHA\n."

    tempo_prod = int(float(row.get("tempoProd", 0)))
    status_atual = row.get('STATUS')
    
    desc_parts = []

    # 2. Diferenciação por STATUS
    if status_atual == 3:
        # Layout para itens críticos de outros PATANs
        desc_parts.append(f"ITEM CRÍTICO - PATAN: {row.get('patan')}")
        desc_parts.append("ESTADO: AGUARDANDO PUXADA")
        desc_parts.append(f"TEMPO ESTIMADO: {tempo_prod} MIN.")
    else:
        # Layout normal para STATUS 2
        desc_parts.append(f"INICIO: {row.get('horaProdInicial', '')}")
        desc_parts.append(f"FIM: {row.get('horaProdFinal', '')}")
        desc_parts.append(f"TEMPO/PRODUÇÃO: {tempo_prod} MIN.")

    # 3. Informações de Quantidade (comum a ambos)
    desc_parts.append(f"QTD. - {row.get('kanbans', 0)} K = {row.get('qtdPecasSeremProduzidas', 0)} PÇS")

    # 4. Processamento de Componentes (sua lógica original mantida)
    comp_comb_str = str(row.get('compComb', ''))
    if pd.notna(row.get('compComb')) and comp_comb_str.lower() != 'nan':
        componentes = comp_comb_str.split('$$$$')
        for comp_info in componentes:
            comp_info = comp_info.strip()
            if not comp_info: 
                continue
            try:
                comp_parts = comp_info.split('|')
                comp_nome = comp_parts[0].strip()
                # Mantendo sua lógica de pegar o índice 2 para quantidade
                comp_qtd = int(float(comp_parts[2].strip()))
                
                tipo = "(ESTAMPADO)" if 'E' in comp_nome.upper() else "(VtoV)"
                desc_parts.append(f"{tipo}: {comp_nome} - {comp_qtd} PÇS")
            except (ValueError, IndexError):
                pass

    return "\n".join(desc_parts)

def create_worksheet_planner_reformulated(df_input, linha_str):
    """Recebe um DataFrame e o formata para o planner final."""
    if df_input.empty:
        return pd.DataFrame()
        
    df = df_input.copy()
    
    today = datetime.today().date()
    if df['turno'].iloc[0] == 1:
        today += timedelta(days=1)
        
    df['data'] = today
    df['checklist'] = "1 - (ABASTECIMENTO) ESTAMPADO ABASTECIDO;2 - (ABASTECIMENTO) EMBALAGEM E VTOV ABASTECIDOS;3 - (PRODUÇÃO) PRÉ-SETUP;4 - (PRODUÇÃO) FECHAMENTO;5 - LIMPEZA WIP2 OU DTR3"
    df['tags'] = df.apply(_build_tags, axis=1)
    df['descricao'] = df.apply(_build_descricao, axis=1)
    
    lista_de_dfs = []
    data1_str = today.strftime("%d/%m")
    
    df = df.sort_values(by=['posto', 'sequencia']).reset_index(drop=True)

    for posto_name, group in df.groupby("posto"):
        header_row = {
            "Material": f"{group['turno'].iloc[0]}° TURNO - {data1_str}",
            "posto": posto_name,
            "tags": f"PONT. {str(posto_name)[-2:]}",
            "data": today,
            "linha": int(linha_str),
            "STATUS": 2
        }
        lista_de_dfs.append(pd.DataFrame([header_row]))
        lista_de_dfs.append(group)

    return pd.concat(lista_de_dfs, ignore_index=True)