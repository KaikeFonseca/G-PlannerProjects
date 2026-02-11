import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def montar_patan(letra_patan, linha, turno, excel_path):
    df_patan = pd.read_excel(excel_path)

    # Filtros iniciais
    df_filtered = df_patan[
        (df_patan["patan"] == letra_patan) &
        (df_patan["linha"] == f"LINHA {linha}") &
        (df_patan["turno"] == int(turno))
    ].copy()

    df_diario_de_bordo = pd.DataFrame(columns=["Material", "DescricaoDiarioDeBordo", "QuantidadeFaltante"])
    df_com_erros = pd.DataFrame(columns=df_filtered.columns.tolist() + ["ErroDescricao"])
    df_output_temp = [] # Usar uma lista para construir o DataFrame e depois concatenar

    for index, row in df_filtered.iterrows():
        material = row["Material"]
        
        # Tratar valores NaN
        cols_to_check = ["pcs/embalagem", "qtdCaixas", "tempoProd", "kanbanMax", "totalLivre"]
        if row[cols_to_check].isnull().any():
            df_com_erros.loc[len(df_com_erros)] = row.tolist() + ["Valores NaN em colunas críticas"]
            continue

        pcs_embalagem = row["pcs/embalagem"]
        qtd_caixas_original = row["qtdCaixas"]
        tempo_prod_original = row["tempoProd"]
        kanban_max = row["kanbanMax"]
        total_livre_original = row["totalLivre"]
        comp_comb_str = str(row["compComb"])
        op_value = int(row["op"]) # Get op value
        prod_em_linha = 0 if op_value == 10 else 1
        lote_patan = row["lotePatan"]
        lead_time = row["leadTime"]
        total_livre = total_livre_original if turno == 2 else (total_livre_original - lead_time) if turno == 3 else (total_livre_original - (lead_time*2))
        estoque_kanban_max = kanban_max*pcs_embalagem
        diff_estoque = total_livre - estoque_kanban_max
        obs = None
        STATUS = 2
        qtd_caixas_atual = qtd_caixas_original
        info_comp_faltante = []

        # Regra de Overproduction
        if diff_estoque > 0:
            obs = "Over"
            #INSERIR AQUI VALOR VERDADEIRO PARA A VARIAVEL QUE SERÁ RESPONSÁVEL PELA CARACTERIZAÇÃO DE OVERPRODUCTION
            #df_diario_de_bordo.loc[len(df_diario_de_bordo)] = [material, "Overproduction", np.nan]
            #continue

        # Regra para a prdoução
        if (total_livre+lote_patan) > estoque_kanban_max: #SE VERDADE, PRODUZ PATAN
            print(lote_patan)
            #ou
            qtd_caixas_atual = qtd_caixas_original #Desta forma, qtd de kanban é igual a do PATAN
        else:
            while(qtd_caixas_atual*pcs_embalagem < abs(diff_estoque)):
                qtd_caixas_atual += 1


        tempo_prod_atual = (tempo_prod_original * qtd_caixas_atual)/qtd_caixas_original

        #REGRA ANTIGA DE PRODUÇÃO
        """# Regra de Redução de Caixas
        qtd_caixas_atual = qtd_caixas_original
        tempo_prod_atual = tempo_prod_original
        
        while ((total_livre + (pcs_embalagem * qtd_caixas_atual)) / qtd_caixas_atual) > kanban_max:
            if qtd_caixas_atual - 1 <= 0: # Evitar qtdCaixas <= 0
                break
            
            # Calcular novo tempoProd proporcionalmente
            novo_tempo_prod_candidato = (tempo_prod_original * (qtd_caixas_atual - 1)) / qtd_caixas_original
            
            if novo_tempo_prod_candidato < 50:
                # Se o novo tempoProd for menor que 50, não reduz mais e mantém o valor atual de qtd_caixas_atual
                break 
            else:
                qtd_caixas_atual -= 1
                tempo_prod_atual = novo_tempo_prod_candidato"""

        qtd_pecas_serem_produzidas = pcs_embalagem * qtd_caixas_atual
        kanbans = qtd_pecas_serem_produzidas // pcs_embalagem

        # Processar compComb
        comp_comb_output = []
        if comp_comb_str and comp_comb_str != 'nan': # Check if string is not empty or 'nan'
            componentes = comp_comb_str.split('$$$$') # Split by newline followed by $$$$
            count_aux = 0
            for comp_info in componentes:
                count_aux+=1
                try:
                    comp_parts = comp_info.split('|')
                    componente_nome = comp_parts[0].strip()
                    qtd_comp_por_peca = int(comp_parts[1].strip())
                    descricao_componente = comp_parts[2].strip()
                    
                    # Handle empty string for estoqueComp
                    estoque_comp_str = comp_parts[3].strip()
                    estoque_comp = float(estoque_comp_str) if estoque_comp_str else 0.0

                    total_comp_necessario = qtd_comp_por_peca * qtd_pecas_serem_produzidas

                    if total_comp_necessario > estoque_comp:
                        quantidade_faltante = total_comp_necessario - estoque_comp
                        """df_diario_de_bordo.loc[len(df_diario_de_bordo)] = [
                            material, 
                            f'Falta de Componente: {componente_nome}', 
                            quantidade_faltante
                        ]"""
                        obs = "Falta comp."
                        #print(continuar amanha - FAZER  A PARTE Q ELE IRA SOMAR EM 'info_comp_faltante')
                        info_comp_faltante.append(f'{componente_nome}\n')

                    if count_aux == len(componentes):
                        comp_comb_output.append(f'{componente_nome} | {qtd_comp_por_peca} | {round(estoque_comp)}')
                    else:
                        comp_comb_output.append(f'{componente_nome} | {qtd_comp_por_peca} | {round(estoque_comp)}$$$$')

                except (ValueError, IndexError) as e:
                    df_com_erros.loc[len(df_com_erros)] = row.tolist() + [f'Erro ao processar compComb: {e} - {comp_info}']
                    break

        # Determinar a coluna de sequência
        sequencia_col = f'seq{letra_patan}'
        sequencia = row[sequencia_col] if pd.notna(row[sequencia_col]) else np.nan # Keep NaN for sorting
        df_output_temp.append({
            'Material': material, 
            'posto': row['posto'], 
            'patan': row['patan'], 
            'linha': linha, 
            'turno': row['turno'],
            'qtdPecasSeremProduzidas': qtd_pecas_serem_produzidas,
            'qtdPorKanban': pcs_embalagem, 
            'kanbans': kanbans,
            'tempPeca': row['tempPeca'], 
            'tempoProd': tempo_prod_atual, 
            'sequencia': sequencia, 
            'prodEmLinha': prod_em_linha,
            'compComb': ''.join(comp_comb_output),
            'estoqueMaterial': total_livre,
            'estoqueKanbanMax': estoque_kanban_max,
            'diff': diff_estoque,
            'obs': obs,
            'STATUS': STATUS
        })
    df_output = pd.DataFrame(df_output_temp)

    # Adicionar horaProdInicial, horaProdFinal e descricaoRefeicao
    if not df_output.empty:
        turn_start_times = {
            1: datetime.strptime('06:00', '%H:%M').time(),
            2: datetime.strptime('14:40', '%H:%M').time(),
            3: datetime.strptime('22:40', '%H:%M').time()
        }

        df_output['horaProdInicial'] = pd.NaT
        df_output['horaProdFinal'] = pd.NaT
        df_output['descricaoRefeicao'] = ''

        df_output = df_output.sort_values(by=["sequencia"])

        df_output["sequencia"] = df_output.groupby("posto").cumcount() + 1
        df_output = df_output.sort_values(by=["posto", "sequencia"]).reset_index(drop=True)

        # Função para calcular os horários por grupo (posto)
        def calculate_times_for_group(group_df):
            group_df = group_df.copy() # Work on a copy of the group to avoid SettingWithCopyWarning
            turno_atual = group_df.iloc[0]["turno"]
            meal_times = {
                1: datetime.strptime("11:00", "%H:%M").time(),
                2: datetime.strptime("18:00", "%H:%M").time(),
                3: datetime.strptime("03:00", "%H:%M").time()
            }

            for i in range(len(group_df)):
                tempo_prod_item = group_df.iloc[i]["tempoProd"]
                sequencia_atual = group_df.iloc[i]["sequencia"]

                if sequencia_atual == 1: # First item in the sequence for this posto
                    current_time_dt = datetime.combine(datetime.today().date(), turn_start_times[turno_atual])
                else:
                    # Get horaProdFinal of the previous item in the group
                    prev_hora_prod_final_time = group_df.iloc[i-1]["horaProdFinal"].time()
                    current_time_dt = datetime.combine(datetime.today().date(), prev_hora_prod_final_time)
                
                # Handle meal time crossing for the current day
                meal_datetime = datetime.combine(datetime.today().date(), meal_times[turno_atual])
                if current_time_dt.time() > meal_datetime.time():
                    meal_datetime += timedelta(days=1)

                # Check if the production time crosses the meal time
                if current_time_dt < meal_datetime and \
                    (current_time_dt + timedelta(minutes=tempo_prod_item)) > meal_datetime:
                    current_time_dt += timedelta(minutes=40)
                    group_df.loc[group_df.index[i], "descricaoRefeicao"] = "TEMPO DE REFEIÇÃO ADICIONADO AO CARTÃO"
                
                group_df.loc[group_df.index[i], "horaProdInicial"] = current_time_dt
                group_df.loc[group_df.index[i], "horaProdFinal"] = current_time_dt + timedelta(minutes=tempo_prod_item)
            
            return group_df            
        df_output = df_output.groupby("posto", group_keys=False).apply(calculate_times_for_group)
        # Formatar horaProdInicial e horaProdFinal para HH:MM
        df_output['horaProdInicial'] = df_output['horaProdInicial'].dt.strftime('%H:%M')
        df_output['horaProdFinal'] = df_output['horaProdFinal'].dt.strftime('%H:%M')

    return df_output, df_diario_de_bordo, df_com_erros

if __name__ == '__main__':
    # Exemplo de uso:
    # Crie um arquivo Excel de teste ou use o seu AutomacaoPlanner.xlsx
    # Certifique-se de que o caminho do arquivo está correto
    excel_file_path = r"\\sb2-fs\11_GESTAO_DA_LOGISTICA$\LOGISTICA\104 - AutomacaoPlanner\Automação Planner\AutomacaoPlanner.xlsx"
    
    # Simular inputs do usuário
    letra_patan_input = 'A'
    linha_input = '2'
    turno_input = '1'

    df_result, df_diary, df_errors = montar_patan(letra_patan_input, linha_input, turno_input, excel_file_path)

    print('\n--- DataFrame de Saída ---')
    print(df_result.to_string())

    print('\n--- Diário de Bordo ---')
    print(df_diary.to_string())

    print('\n--- Erros Encontrados ---')
    print(df_errors.to_string())

    # Salvar resultados em arquivos para verificação
    df_result.to_excel('output_patan.xlsx', index=False)
    df_diary.to_excel('diario_de_bordo.xlsx', index=False)
    df_errors.to_excel('erros_processamento.xlsx', index=False)