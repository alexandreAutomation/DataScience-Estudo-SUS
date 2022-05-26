#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------------
# Created By  : Alexandre Edson Silva Pereira
# Created Date: 2021/12/31 12:46
# version ='5.0'
# ---------------------------------------------------------------------------
"""
Esse script foi construído para o processo seletivo da Petz, onde é feito
a estimativa de dados dos meses faltantes usando a base de dados tratada
case_internacao_SUS.xlsx

Para executar esse script é necessario a instalação dos seguintes módulos:

pip install pandas
pip install xlrd
pip install xlsxwriter

Deus é fiel

"""
# ---------------------------------------------------------------------------

# Imports
import pandas as pd
import numpy as np

# Declaração das Variáveis
writer = pd.ExcelWriter('AlexandreEdsonEx1Desafio3.xlsx', engine = 'xlsxwriter')
months = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
years = range(17, 20)
sheets_df = []
months_missing = []
sheet_df = pd.DataFrame
sheet_missing_df = pd.DataFrame({})
list_states = []
state_general_list = []

# Loop for para gerar Sheets de cada ano
for year in years:

    # Loop para gerar Sheets de cada mês referente ao ano do loop anterior
    for month_num, month in enumerate(months, start=1):

        # try para ignora o erro das Sheets do mês/ano não encontrados na base de dados "case_internacao_SUS.xls"
        try:
            # Cria um DataFrame nomeado sheet_df
            sheet_df = pd.DataFrame(pd.read_excel('AlexandreEdsonEx1Desafio1.xlsx', sheet_name=f'{month}{year}'))

            # Insert uma nova coluna "Período", se o ano for menor que 10, acrescenta o digito 0, se não, segue.
            if month_num < 10:
                sheet_df.insert(0, 'Periodo', f'20{year}0{month_num}01', allow_duplicates = False)
            else:
                sheet_df.insert(0, 'Periodo', f'20{year}{month_num}01', allow_duplicates = False)

            # Acresenta na lista os dataframes
            sheets_df.append(sheet_df)
        except:
            # Cria uma lista com os períodos não encontrados
            if year == 17 or (year == 19 and month_num > 7):
                pass
            else:
                if month_num < 10:
                    months_missing.append(f'20{year}0{month_num}01')
                else:
                    months_missing.append(f'20{year}{month_num}01')

# Resgata o nome das coluas convertanto no tipo list
columns = list(sheet_df.columns.array)

# Cria um dataframe para unir outros datafrme em uma única tabela
sheets_merge_df = pd.DataFrame(columns=columns)

# União das tabelas
for x in range(0, len(sheets_df)):
    sheets_merge_df = pd.merge(sheets_merge_df, sheets_df[x], how='outer')

sheets_merge_df.sort_values(by = ['Periodo', 'Regiao_UF', 'Unidade_federacao'], ascending=False, inplace=True)

# Cria um dataframe somente com as colunas 2, 3 e remove duplicatas, para ser usada no no df Meses faltantes
missing_df = pd.DataFrame(sheets_merge_df.drop_duplicates(['Unidade_federacao']).
                                     drop(columns=columns[3:14]).drop(columns=columns[0]))

# Loop for para gerar períodos faltantes
for month_missing in months_missing:
    missing_df.insert(0, 'Periodo', month_missing, allow_duplicates = False)

    # Inserir estados para cada período faltantes
    for index in range(0, len(missing_df.index)):
        sheet_missing_df = sheet_missing_df.append({'Periodo': missing_df.iloc[index]['Periodo'],
                                                    'Unidade_federacao': missing_df.iloc[index]['Unidade_federacao'],
                                                    'Regiao_UF': missing_df.iloc[index]['Regiao_UF']}, ignore_index=True)
    missing_df.drop(columns=columns[0], inplace=True)

# União da tabela faltante com a tabela dados completos
database_gereral = pd.merge(sheets_merge_df, sheet_missing_df, how='outer')

# Converte Periodo para datatime
database_gereral['Periodo'] = database_gereral['Periodo'].astype('datetime64')

# Classifica dados em três variaveis
database_gereral.sort_values(by = ['Periodo', 'Regiao_UF', 'Unidade_federacao'], ascending=True, inplace=True)

# database_gereral['Periodo'] = database_gereral.Periodo.dt.strftime('%d/%m/%Y')

# Gera uma lista com as tabelas de cada estado
for state in list(missing_df.Unidade_federacao.values):
    database_state = database_gereral[database_gereral.Unidade_federacao.eq(state)]
    database_state.index = [x for x in range(0, database_state.shape[0])]
    list_states.append(database_state)

# Aplica o modelo de estimativa para cada estado
for list_state in list_states:

    # Seleciona linha a linha
    for line in list_state.index:

        # Seleciona coluna por coluna com valores NaN
        for list_column in ['Valor_total', 'AIH_aprovadas', 'Internacoes', 'Valor_serv_hosp', 'Valor_serv_prof',
                            'Dias_permanencia', 'Obitos']:
            # Verifica a exixtencia de NaN
            if str(list_state.loc[line, list_column]) == str(np.float64(np.nan)):
                loop1 = True
                line_sub = 1

                # Faça até conseguir achar os dois valores mais proximos (Um Anterior e outro posterior ao valor NaN
                while loop1:
                    # Se for a primeira linha, não verifica a operação porque não tem um valor anterior
                    if line > 0:
                        # Verifica se no mês anterior é diferente de NaN, se sim, usa ele para calcular a média
                        if str(list_state.loc[line - line_sub, list_column]) != str(np.float64(np.nan)):
                            loop1 = False
                        else:
                            line_sub += 1
                    elif str(list_state.loc[line + line_sub, list_column]) != str(np.float64(np.nan)):
                        loop1 = False
                    else:
                        line_sub += 1

                loop2 = True
                line_sum = 1

                # Mesmas consideraçoes do loop anterior, porém buscando o valor posterior ao NaN
                while loop2:
                    if line < len(list_state.index):
                        if str(list_state.loc[line + line_sum, list_column]) != str(np.float64(np.nan)):
                            loop2 = False
                        else:
                            line_sum += 1
                    elif str(list_state.loc[line - line_sum, list_column]) != str(np.float64(np.nan)):
                        loop2 = False
                    else:
                        line_sum += 1
                # Cálculo da média entre os valores anterior e posterior encontrados nos loops
                list_state.at[line, list_column] = (list_state.loc[line - line_sub, list_column] +
                                                    list_state.loc[line + line_sum, list_column]) / 2

    # Cria uma lista de dataframes estados
    state_general_list.append(list_state)

# Cria uma unica tabela
state_general_df = pd.concat(state_general_list, ignore_index=True)

# Cálculo das variáveis dependentes dos dados estimados
state_general_df['Valor_medio_AIH'] = state_general_df.apply(lambda x_df: x_df['Valor_total'] / x_df['AIH_aprovadas'], axis=1)
state_general_df['Valor_medio_intern'] = state_general_df.apply(lambda x_df: x_df['Valor_total'] / x_df['Internacoes'], axis=1)
state_general_df['Media_permanencia'] = state_general_df.apply(lambda x_df: x_df['Dias_permanencia'] / x_df['Internacoes'], axis=1)
state_general_df['Taxa_mortalidade'] = state_general_df.apply(lambda x_df: (x_df['Obitos'] / x_df['Internacoes']) * 100, axis=1)
state_general_df.sort_values(by = ['Periodo', 'Regiao_UF', 'Unidade_federacao'], ascending=False, inplace=True)
state_general_df.iloc[:, 4:15] = state_general_df.iloc[:, 4:15].apply(pd.to_numeric, errors='coerce')
# Converte Dataframe para excel
state_general_df.to_excel(writer, sheet_name='Geral', index=False)

# Atualiza os espaçamentos entre as colunas e os formatos de cada coluna
workbook = writer.book
worksheet = writer.sheets['Geral']
format1 = workbook.add_format({'num_format': '#,0.00'})
format2 = workbook.add_format({'num_format': '#'})
worksheet.set_column('D:E', 18, format2)
worksheet.set_column('K:K', 18, format2)
worksheet.set_column('M:M', 18, format2)
worksheet.set_column('F:J', 18, format1)
worksheet.set_column('L:L', 18, format1)
worksheet.set_column('N:N', 18, format1)

for i, col in enumerate(list(sheet_df.columns.array)):
    column_len = sheet_df[col].astype(str).str.len().max()
    worksheet.set_column(i, i, max(column_len, len(col) + 4))

worksheet.set_column(0, 0, 17)

# Salva os dados tratados na nova planilha 'AlexandreEdsonEx1Desafio3.xlsx'
writer.save()

# ----------------------------------------------------------------------------
# Grato pela Oportunidade e o conhecimento adquirido até aqui
