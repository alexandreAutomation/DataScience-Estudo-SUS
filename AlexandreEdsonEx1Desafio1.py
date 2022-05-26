#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------------
# Created By  : Alexandre Edson Silva Pereira
# Created Date: 2021/12/26 17:45
# version ='3.0'
# ---------------------------------------------------------------------------
"""
Esse script foi contruído para o processo seletivo da Petz, onde é feito
o tratamento dos dados providos da planilha case_internacao_SUS.xlsx

Para executar esse script é necessario a instalação dos seguintes modulos:

pip install pandas
pip install xlrd
pip install xlsxwriter
pip install requests

Deus é fiel
"""
# ---------------------------------------------------------------------------

# Imports
import pandas as pd
from json import loads
from requests import get

# Declaração das Variáveis
months = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
writer = pd.ExcelWriter('AlexandreEdsonEx1Desafio1.xlsx', engine = 'xlsxwriter')
years = range(17, 20)
region_dict = {}
regions = {}

# API IBGE regiões das UF, armazena dados na variável 'api_regioes'
while True:
    api_regioes = loads(get('https://servicodados.ibge.gov.br/api/v1/localidades/mesorregioes/').text)
    break

# Loop for para gerar Sheets de cada ano
for year in years:
    
    # Loop para gerar Sheets de cada mês referente ao ano do loop anterior
    for month in months:
        
        # try para ignora o erro das Sheets do mês/ano não encontrados na base de dados "case_internacao_SUS.xls"
        try:
            
            # Cria um DataFrame nomeado sheet_df
            sheet_df = pd.DataFrame(pd.read_excel('case_internacao_SUS.xls', sheet_name=f'{month}{year}'))
            
            # Remove as colunas com dados faltantes/Irrelevantes para a analise
            sheet_df = sheet_df.drop(columns=['Val_serv_hosp_-_compl_federal', 'Val_serv_hosp_-_compl_gestor',
                                              'Val_serv_prof_-_compl_federal', 'Val_serv_prof_-_compl_gestor'])
            
            # renomeia nome das colunas
            sheet_df.rename({'Região/Unidade da Federação': 'Unidade_federacao',
                             'Valor_serviços_hospitalares': 'Valor_serv_hosp',
                             'Valor_serviços_profissionais': 'Valor_serv_prof', 'Internações': 'Internacoes',
                             'Valor_médio_AIH': 'Valor_medio_AIH', 'Dias_permanência': 'Dias_permanencia',
                             'Óbitos': 'Obitos', 'Média_permanência': 'Media_permanencia',
                             'Valor_médio_intern': 'Valor_medio_intern'}, axis = 1, inplace = True)
            
            # Tratamentos dos dados
            sheet_df.Unidade_federacao.replace(regex=r'^\.. ', value='', inplace = True)
            sheet_df.iloc[:, 2:18] = sheet_df.iloc[:, 2:18].apply(pd.to_numeric, errors='coerce')
            sheet_df.dropna(axis = 0, how = 'all', inplace = True)
            sheet_df.drop(sheet_df.tail(1).index, inplace = True)
            sheet_df = sheet_df[sheet_df['Unidade_federacao'] != 'Total']
            
            # Cria uma nova coluna regiao
            sheet_df.insert(1, 'Regiao_UF', '', allow_duplicates = False)
            
            # Loop para remover os dados referentes as regiões (Dados irrelevantes pois já existem por UF)
            for local in sheet_df.Unidade_federacao.values:
                if 'região' in str(local).lower():
                    sheet_df = sheet_df[(sheet_df['Unidade_federacao'] != local)]
                    
            # Filtra dados consumidos da API_IBGE conforme a lista de cada UF na base de dados
            for uf in list(sheet_df.Unidade_federacao.values):
                for region in range(0, len(api_regioes)):
                    if uf == api_regioes[region]['UF']['nome']:
                        regions[f"{uf}"] = api_regioes[region]['UF']['regiao']['nome']
                        
            # Atualiza a coluna 'Regiao_UF' acrescentando as regiões respectivos a cada UF com os dados filtrados no loop
            # anterior
            for index, row in sheet_df.iterrows():
                for uf, region in regions.items():
                    if row['Unidade_federacao'] == uf:
                        sheet_df.loc[index, 'Regiao_UF'] = str(regions[f'{uf}'])
            
            # Ordenação das linhas por Regiao_UF e Unidade_federacao
            sheet_df.sort_values(by = ['Regiao_UF', 'Unidade_federacao'], ascending=False, inplace = True)

            # Converte Dataframe para excel
            sheet_df.to_excel(writer, sheet_name=f'{month}{year}', index = False)

            # Converte a saida para melhor visualização dos dados
            workbook = writer.book
            worksheet = writer.sheets[f'{month}{year}']
            format1 = workbook.add_format({'num_format': '#,0.00'})
            worksheet.set_column('E:I', 18, format1)
            worksheet.set_column('K:K', 18, format1)
            worksheet.set_column('M:M', 18, format1)

            # Atualiza os espaçamentos entre as colunas
            for i, col in enumerate(list(sheet_df.columns.array)):
                column_len = sheet_df[col].astype(str).str.len().max()
                worksheet.set_column(i, i, max(column_len, len(col) + 4))

        # Fim do Try except
        except Exception as e:
            pass

# Salva os dados tratados na nova planilha 'AlexandreEdsonEx1Desafio1.xlsx'
writer.save()

# ----------------------------------------------------------------------------
# Grato pela Oportunidade e o conhecimento adquirido até aqui
