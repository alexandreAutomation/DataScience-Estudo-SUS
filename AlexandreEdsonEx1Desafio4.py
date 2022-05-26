#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------------
# Created By  : Alexandre Edson Silva Pereira
# Created Date: 2022/01/03 00:20
# version ='6.0'
# ---------------------------------------------------------------------------
"""
Esse script foi construído para o processo seletivo da Petz, onde é feito
a previsão de 6 meses para cada estado com excessão ao estado do acre que teve sua
previsão para 7 meses devido não constas na base de dados original, usando a base
de dados tratada AlexandreEdsonEx1Desafio3.xlsx

Para executar esse script é necessario a instalação dos seguintes módulos:

pip install pandas
pip install xlrd
pip install xlsxwriter
pip install pmdarima

Deus é fiel
"""
# ---------------------------------------------------------------------------

# Imports
import pandas as pd
from pmdarima import auto_arima

# Declaração das Variáveis
model_arima = []
future_months = ['2019-08-01', '2019-09-01', '2019-10-01', '2019-11-01', '2019-12-01', '2020-01-01']
future_months_acre = ['2019-07-01', '2019-08-01', '2019-09-01', '2019-10-01', '2019-11-01', '2019-12-01', '2020-01-01']
writer = pd.ExcelWriter('AlexandreEdsonEx1Desafio4.xlsx', engine = 'xlsxwriter')
states_predict = []

# Importação da dase de dados
past_months = pd.DataFrame(pd.read_excel('AlexandreEdsonEx1Desafio3.xlsx', sheet_name='Geral',
                                                     index_col='Periodo', parse_dates=True))

# Seleciona as colunas para previsão
past_months = past_months[['Unidade_federacao', 'AIH_aprovadas', 'Valor_medio_AIH', 'Obitos']]

# Classifica os dados por Periodo
past_months.sort_index(inplace=True)

# Busca o nome das colunas para o loop for
states = past_months.drop_duplicates(['Unidade_federacao'])[['Unidade_federacao']].values

# Loop for para aplicar o metodo para cada estado
for state in states:

    # Seleciona o estado
    past_months_df = pd.DataFrame(past_months.loc[
                                           past_months.Unidade_federacao == f'{state[0]}']
                                           [['AIH_aprovadas', 'Valor_medio_AIH', 'Obitos']])

    # AIH_aprovadas = ARIMA(0,1,0)(0,1,0)[12] intercept   : AIC=147.474, Time=0.02 sec
    # Internacoes = ARIMA(0,1,0)(0,1,0)[12] intercept   : AIC=147.589, Time=0.02 sec
    # Valor_total = ARIMA(0,1,3)(0,1,0)[12] intercept   : AIC=249.854, Time=0.02 sec
    # Obitos = ARIMA(1,1,0)(0,1,0)[12] intercept   : AIC=111.114, Time=0.03 sec

    # Aplica o metodo para cada coluna de cada estado
    for past_column in list(past_months_df.columns.array):

        # Modelp ARIMA para prever dados temporais
        model = auto_arima(past_months_df[f'{past_column}'], start_p=1,    # se trata do valor inicial dentro do range de aprendizado do valor p (AR)
                                  start_d=1,    # se trata do valor inicial dentro do range de aprendizado do valor d (I)
                                  start_q=1,    # se trata do valor inicial dentro do range de aprendizado do valor q (MA)
                                  max_p=8,  # Valor máximo dentro do range de aprendizado do valor p (AR)
                                  max_d=8,  # Valor máximo dentro do range de aprendizado do valor d (I)
                                  max_q=8,  # Valor máximo dentro do range de aprendizado do valor q (MA)
                                  m=12,  # Valor que se referente ao período da diferenciação sazonal, 7 é igual a diário
                                  start_P=0,    # Com p maiúsculo é referente ao valor inicial do modelo AR para sazonalidade. Considere 0 por que o padrão é 1.
                                  seasonal=True,    # dica se deve usar um ARIMA ou SARIMA, ou seja, o conjunto possui sazonalidade ou não
                                  d=1,  # d e D Indicam a ordem da primeira e das demais diferenciação da sazonalidade, caso não sejam informados,
                                  D=1,  # os valores serão selecionados através de um teste de sazonalidade feito pelo próprio modelo
                                  trace=False,   # Indica se deve ser impresso o acompanhamento do aprendizado feito pelo modelo.
                                  error_action='ignore',    # Indica se deve notificar ou não caso o conjunto de dados tem um problema de dados estacionários (sem possibilidade de prever)
                                  stepwise=True)    # Indica se deve ser utilizado um algorítimo especifico chamado Stepwise para aprendizado dos parâmetros (Hyperparameter)

        # Condição para o estado Acre, quando for a vez dele, terá 7 meses de previsão(LEIA O CABEÇALHO DP SCRIPT)
        if state[0] != 'Acre':
            future_values = model.predict(n_periods=6)
            future_values = pd.DataFrame(future_values, index=future_months, columns=[f'{past_column}'])
        else:
            future_values = model.predict(n_periods=7)
            future_values = pd.DataFrame(future_values, index=future_months_acre, columns=[f'{past_column}'])
        future_values.index.name = 'Periodo'
        future_values.index = pd.to_datetime(future_values.index)
        future_values[f'{past_column}'] = round(future_values[f'{past_column}'], 0)

        # Guarda os resultados das colunas em uma lista
        model_arima.append(pd.DataFrame(future_values[f'{past_column}']))

    # União das colunas e armazena na variavel "states_predict"
    future_values = pd.merge(model_arima[0], model_arima[1], how='inner', on='Periodo')
    future_values1 = pd.merge(future_values, model_arima[2], how='inner', on='Periodo')
    future_values1.insert(0, 'Unidade_federacao', f'{state[0]}', allow_duplicates=False)
    states_predict.append(future_values1)
    print(future_values1)
    del model_arima[0:99]

# Cria um dataframe para ser convertido em excel
future_values_general = pd.DataFrame(pd.concat(states_predict, ignore_index=False))
future_values_general.reset_index(inplace=True)
future_values_general.sort_values(by = ['Periodo', 'Unidade_federacao'], ascending=False, inplace=True)

# Converte Dataframe para excel
future_values_general.to_excel(writer, sheet_name='teste', index=False)

# Atualiza os espaçamentos entre as colunas e os formatos de cada coluna
workbook = writer.book
worksheet = writer.sheets['teste']
format1 = workbook.add_format({'num_format': '#,0.00'})
format2 = workbook.add_format({'num_format': '#'})
worksheet.set_column('C:D', 18, format2)
worksheet.set_column('F:F', 18, format2)
worksheet.set_column('E:E', 18, format1)

for i, col in enumerate(list(future_values_general.columns.array)):
    column_len = future_values_general[col].astype(str).str.len().max()
    worksheet.set_column(i, i, max(column_len, len(col) + 6))

worksheet.set_column(0, 0, 20)

# Salva os dados tratados na nova planilha 'AlexandreEdsonEx1Desafio4.xlsx'
writer.save()

# ----------------------------------------------------------------------------
# Grato pela Oportunidade e o conhecimento adquirido até aqui
