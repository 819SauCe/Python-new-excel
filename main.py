import os
import pandas as pd
from datetime import datetime
from estilização import estilizar_planilha
import numpy as np
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
meses_para_calculo = pd.date_range(end=datetime.now(), periods=36, freq='MS').strftime('%B, %Y').tolist()
meses_exibicao = meses_para_calculo[-12:]
arquivo = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
status_relatorio = []
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def gerar_relatorio_status(mensagem, sucesso=True):
    status = 'Sucesso' if sucesso else 'Erro'
    status_relatorio.append(f"{status}: {mensagem}")
    if not sucesso:
        print(f"{status}: {mensagem}")
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def classificar_abc(percentual, tipo='geral'):
    thresholds = {'geral': (0.20, 0.75), 'marca_valor': (0.35, 0.75), 'marca': (0.10, 0.20)}
    a, b = thresholds.get(tipo, (0, 1))
    return 'A' if percentual <= a else ('B' if percentual <= b else 'C')
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def import_report(path, filename):
    try:
        df = pd.read_csv(os.path.join(path, filename), sep='|', encoding='ISO-8859-1', low_memory=False)
        gerar_relatorio_status(f"{filename} carregado com sucesso.")
        return df
    except Exception as e:
        gerar_relatorio_status(f"Erro ao ler {filename}: {e}", sucesso=False)
        return pd.DataFrame()
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def carregar_planilhas(diretorio):
    planilhas = [
        ('TI-PBI_EstoqueLote.txt', 'Nome do Produto'),
        ('TI-PBI_VendasProdutos.txt', 'Nome Prod/Serv'),
        ('CF_TI-PBI_Estoque.txt', 'Nome do Produto'),
        ('TI-PBI_ItensOrdensCompras.txt', 'Produto/Sv.')
    ]
    dfs = []
    for arquivo, nome_coluna in planilhas:
        df = import_report(diretorio, arquivo)
        if not df.empty:
            if nome_coluna in df.columns:
                df.rename(columns={nome_coluna: 'Nome Produto'}, inplace=True)
            if 'Código Produto' in df.columns:
                df = df[['Código Produto', 'Nome Produto']].drop_duplicates()
            dfs.append(df)
    
    if dfs:
        df_produtos = pd.concat(dfs, ignore_index=True).drop_duplicates(subset=['Código Produto'], keep='first')
        gerar_relatorio_status("Planilhas carregadas e consolidadas com sucesso.")
        return df_produtos
    gerar_relatorio_status("Erro ao carregar planilhas: dados ausentes ou inválidos.", sucesso=False)
    return pd.DataFrame()
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def df_vendas(diretorio):
    df = import_report(diretorio, 'TI-PBI_VendasProdutos.txt')
    if df.empty:
        return df

    try:
        df.rename(columns={'Código Prod/Serv': 'Código Produto', 'Quantidade': 'Quantidade', 'Data Emissão NF': 'Data Emissão NF'}, inplace=True)
        df['Data Emissão NF'] = pd.to_datetime(df['Data Emissão NF'], format='%d/%m/%Y', errors='coerce')
        df['Quantidade'] = pd.to_numeric(df['Quantidade'], errors='coerce').fillna(0).astype(int)
        
        df['Mês/Ano'] = df['Data Emissão NF'].dt.to_period('M').dt.strftime('%B, %Y')
        df_pivot = df.pivot_table(index=['Código Produto'], columns='Mês/Ano', values='Quantidade', fill_value=0, aggfunc='sum').reset_index()

        for mes in meses_para_calculo:
            if mes not in df_pivot.columns:
                df_pivot[mes] = 0

        df_pivot = df_pivot[['Código Produto'] + meses_exibicao]
        gerar_relatorio_status("Dados de vendas processados com sucesso.")
        return df_pivot
    
    except Exception as e:
        gerar_relatorio_status(f"Erro ao processar vendas: {e}", sucesso=False)
        return pd.DataFrame()
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def carregar_estoque(diretorio):
    df_estoque = import_report(diretorio, 'CF_TI-PBI_Estoque.txt')
    if 'Código Produto' in df_estoque.columns and 'Qtde Disponível' in df_estoque.columns and 'Inativo' in df_estoque.columns:
        df_estoque['Qtde Disponível'] = pd.to_numeric(df_estoque['Qtde Disponível'], errors='coerce').fillna(0).astype(int)
        df_estoque = df_estoque[['Código Produto', 'Qtde Disponível', 'Inativo']]
        df_estoque['Inativo'] = df_estoque['Inativo'].apply(lambda x: 'Inativo' if str(x).strip().lower() in ['1', 'inativo', 'sim'] else 'Ativo')
        gerar_relatorio_status("Dados de estoque carregados com sucesso.")
    else:
        gerar_relatorio_status("Erro ao carregar dados de estoque: colunas necessárias ausentes.", sucesso=False)
    return df_estoque
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def df_order(diretorio):
    df = import_report(diretorio, 'TI-PBI_ItensOrdensCompras.txt')
    if df.empty:
        return df

    try:
        # Selecionar apenas as colunas relevantes
        df = df[['Situação', 'Cód.Pr.Sv.', 'Nr.OC', 'Em Trânsito']]
        df['Código Produto'] = df['Cód.Pr.Sv.']
        df = df.drop(['Cód.Pr.Sv.'], axis=1)
        df = df[(df['Situação'] == 'Parcialm. Atend.') | (df['Situação'] == 'Pendente')]
        df['Ordem de Compra'] = df['Nr.OC']

        # Verificar os valores únicos na coluna 'Em Trânsito' para diagnóstico
        print("Valores únicos na coluna 'Em Trânsito' antes da conversão:")
        print(df['Em Trânsito'].unique())

        # Converter os valores 'Sim' e 'Não' para valores numéricos
        df['Em Trânsito'] = df['Em Trânsito'].str.strip().str.lower().apply(lambda x: 1 if 'sim' in str(x) else 0)

        # Imprimir os primeiros valores da coluna 'Em Trânsito' após a conversão para diagnóstico
        print("Valores da coluna 'Em Trânsito' após a conversão:")
        print(df['Em Trânsito'].head(10))

        gerar_relatorio_status("Dados de ordens processados com sucesso.")
    except Exception as e:
        gerar_relatorio_status(f"Erro ao processar ordens: {e}", sucesso=False)
        return pd.DataFrame()

    return df
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def validades(diretorio):
    df = import_report(diretorio, 'TI-PBI_EstoqueLote.txt')
    if df.empty:
        gerar_relatorio_status("Arquivo TI-PBI_EstoqueLote.txt está vazio ou não encontrado.", sucesso=False)
        return df

    try:
        if 'Marca' in df.columns:
            df['Marca'] = df['Marca'].astype(str).str.strip()
        else:
            gerar_relatorio_status("Coluna 'Marca' não encontrada no arquivo TI-PBI_EstoqueLote.txt.", sucesso=False)
            df['Marca'] = None

        if 'S' in df.columns:
            df['Disponível'] = pd.to_numeric(df['S'], errors='coerce').fillna(0).astype(int)

        # Garantir que a Data Última Compra é extraída e convertida corretamente
        if 'Data Última Compra' in df.columns:
            df['Data Última Compra'] = pd.to_datetime(df['Data Última Compra'], format='%d/%m/%y', errors='coerce')
        elif df.columns.size >= 8:
            df['Data Última Compra'] = pd.to_datetime(df.iloc[:, 7], format='%d/%m/%y', errors='coerce')

        df['Qtd. Lotes'] = 1
        df = df[['Código Produto', 'Marca', 'Validade', 'Qtd. Lotes', 'Disponível', 'Data Última Compra']]
        df = df.groupby('Código Produto').agg({
            'Marca': 'first',
            'Validade': 'first',
            'Qtd. Lotes': 'sum',
            'Disponível': 'sum',
            'Data Última Compra': 'max'  # Priorizar a data mais recente
        }).reset_index()

        gerar_relatorio_status("Dados de validades processados com sucesso.")
    except Exception as e:
        gerar_relatorio_status(f"Erro ao processar validades: {e}", sucesso=False)
        return pd.DataFrame()

    return df
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def completar_dados(df_vendas, df_estoque, df_produtos, df_order, df_lotes, df_marcas_vendas, df_marcas_estoque):
    # Padronizar os códigos dos produtos
    df_vendas['Código Produto'] = df_vendas['Código Produto'].astype(str).str.strip().str.upper()
    df_estoque['Código Produto'] = df_estoque['Código Produto'].astype(str).str.strip().str.upper()
    df_produtos['Código Produto'] = df_produtos['Código Produto'].astype(str).str.strip().str.upper()
    df_order['Código Produto'] = df_order['Código Produto'].astype(str).str.strip().str.upper()
    df_lotes['Código Produto'] = df_lotes['Código Produto'].astype(str).str.strip().str.upper()
    df_marcas_vendas['Código Produto'] = df_marcas_vendas['Código Produto'].astype(str).str.strip().str.upper()
    df_marcas_estoque['Código Produto'] = df_marcas_estoque['Código Produto'].astype(str).str.strip().str.upper()

    # Combinar os DataFrames
    df_completo = pd.merge(df_estoque, df_vendas, on='Código Produto', how='outer')
    df_completo = pd.merge(df_completo, df_produtos, on='Código Produto', how='left', suffixes=('', '_produtos'))
    df_completo = pd.merge(df_completo, df_order, on='Código Produto', how='outer', suffixes=('', '_order'))
    df_completo = pd.merge(df_completo, df_lotes, on='Código Produto', how='outer', suffixes=('', '_lotes'))

    # Resolver colunas conflitantes
    if 'Em Trânsito_order' in df_completo.columns:
        df_completo['Em Trânsito'] = df_completo['Em Trânsito'].combine_first(df_completo['Em Trânsito_order'])
        df_completo.drop(columns=['Em Trânsito_order'], inplace=True)

    # Preencher marcas faltantes com diferentes fontes, priorizando na ordem
    df_completo = pd.merge(df_completo, df_marcas_vendas, on='Código Produto', how='left', suffixes=('', '_vendas'))
    df_completo['Marca'] = df_completo['Marca'].combine_first(df_completo['Marca_vendas'])
    df_completo.drop(columns=['Marca_vendas'], inplace=True)

    df_completo = pd.merge(df_completo, df_marcas_estoque, on='Código Produto', how='left', suffixes=('', '_estoque'))
    df_completo['Marca'] = df_completo['Marca'].combine_first(df_completo['Marca_estoque'])
    df_completo.drop(columns=['Marca_estoque'], inplace=True)

    # Garantir que a Data Última Compra está correta
    if 'Data Última Compra' in df_lotes.columns:
        df_completo['Data Última Compra'] = df_completo['Data Última Compra'].combine_first(df_lotes['Data Última Compra'])

    # Tratar colunas numéricas
    df_completo['Qtde Disponível'] = pd.to_numeric(df_completo['Qtde Disponível'], errors='coerce').fillna(0)
    df_completo['Em Trânsito'] = pd.to_numeric(df_completo['Em Trânsito'], errors='coerce').fillna(0)
    df_completo['Qtd a Receber'] = df_completo['Em Trânsito']
    df_completo['Saldo Total'] = df_completo['Qtde Disponível'] + df_completo['Qtd a Receber']
    df_completo['Em Trânsito'] = df_completo['Em Trânsito'].map({1: 'Sim', 0: 'Não'})

    # Tratar colunas de status
    if 'Inativo' in df_completo.columns:
        df_completo['Inativo'] = df_completo['Inativo'].fillna('Ativo').apply(
            lambda x: 'Inativo' if str(x).strip().lower() in ['1', 'inativo', 'sim'] else 'Ativo'
        )
    else:
        df_completo['Inativo'] = 'Ativo'

    colunas_meses = [col for col in meses_exibicao if col in df_completo.columns]
    df_completo[colunas_meses] = df_completo[colunas_meses].fillna(0).astype(int)

    # Adicionar colunas adicionais conforme necessário
    for coluna in ['Validade', 'Qtd. Lotes', 'Ordem de Compra']:
        if coluna not in df_completo.columns:
            df_completo[coluna] = None

    # Calcular média de vendas
    if 'Media' not in df_completo.columns:
        if colunas_meses:
            df_completo['Media'] = df_completo[colunas_meses].mean(axis=1)
        else:
            df_completo['Media'] = None

    # Calcular tempo de estoque
    if 'Media' in df_completo.columns and 'Saldo Total' in df_completo.columns:
        df_completo['Tempo de Estoque'] = df_completo.apply(
            lambda row: int(round(row['Saldo Total'] / row['Media'])) if row['Media'] > 0 else None, axis=1
        )
    df_completo['Tempo de Estoque'] = df_completo['Tempo de Estoque'].apply(lambda x: x if x is not None and x > 0 else None)

    # Determinar necessidade de compra
    df_completo['Comprar(S/N)'] = df_completo['Tempo de Estoque'].apply(lambda x: 'Sim' if x is not None and x < 6 else 'Não')

    # Calcular unidades de compra necessárias
    if 'Media' in df_completo.columns and 'Saldo Total' in df_completo.columns:
        df_completo['Unidade de Compra'] = df_completo.apply(
            lambda row: max(row['Media'] * 12 - row['Saldo Total'], 0) if row['Media'] is not None else None, axis=1
        )

    gerar_relatorio_status("Dados completos gerados com sucesso.")
    return df_completo
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def adicionar_coluna_total(df):
    colunas_meses = [col for col in meses_exibicao if col in df.columns]
    df[colunas_meses] = df[colunas_meses].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
    df['Total'] = df[colunas_meses].sum(axis=1)
    return df
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def adicionar_coluna_media(df):
    colunas_meses = [col for col in meses_exibicao if col in df.columns]                                                     #Calcular a média apenas para valores maiores
    df['Media'] = df[colunas_meses].apply(lambda row: int(row[row > 0].mean()) if row[row > 0].count() > 0 else '-', axis=1) #que zero e arredondar para o inteiro mais próximo
    return df                                                                                                                #utiliza Dados para criar a coluna Média
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def calcular_curva_abc(df):
    colunas_meses = [col for col in meses_para_calculo if col in df.columns]                                #calcular curva ABC baseada em quantidade de meses
    df['Total Quantidade Meses'] = df[colunas_meses].sum(axis=1)                                            #Implementa a função `calcular_curva_abc(df)` que calcula
    total_qtd = df['Total Quantidade Meses'].sum()                                                          #a curva ABC para os produtos.
                                                                                                            #Soma as quantidades dos meses relevantes, calcula a porcentagem
    if total_qtd != 0:                                                                                      #acumulada e classifica os produtos
        df = df.sort_values(by='Total Quantidade Meses', ascending=False)                                   #nas categorias A, B ou C, considerando quantidade e valor.
        df['Qtd Acumulado'] = df['Total Quantidade Meses'].cumsum() / total_qtd                             #Adiciona classificações `ABC Qtd`, `ABC Valor` e `ABC MARCA Qtd`
        df['ABC Qtd'] = df['Qtd Acumulado'].apply(lambda x: classificar_abc(x, tipo='geral'))               #para os produtos.
        df['ABC Valor'] = df['Qtd Acumulado'].apply(lambda x: classificar_abc(x, tipo='marca_valor'))       #Ordena o DataFrame pela quantidade total e remove a coluna de
        df['ABC MARCA Qtd'] = df['Qtd Acumulado'].apply(lambda x: classificar_abc(x, tipo='marca'))         #quantidade acumulada após o uso.
                                                                                                            #Gera um relatório de status indicando o sucesso do cálculo da
    df = df.drop(columns=['Qtd Acumulado'], errors='ignore')                                                #curva ABC.
    gerar_relatorio_status("Curva ABC calculada com sucesso.")
    return df
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def adicionar_data_ultima_compra(diretorio, df_principal):
    df_lotes = import_report(diretorio, 'TI-PBI_EstoqueLote.txt')                                                           #adiciona a data da ultima compra
                                                                                                                            #coloca os valores em data
    if df_lotes.empty:                                                                                                      #puxa os dados da planilha PBI_EstoqueLote
        gerar_relatorio_status("Erro: Arquivo TI-PBI_EstoqueLote.txt não encontrado ou vazio.", sucesso=False)              #extremamente necessario
        return df_principal                                                                                                 #extremamente necessario
                                                                                                                            #talvez no futuro remover as horas
    try:
        if df_lotes.columns.size >= 8:
            df_lotes.rename(columns={df_lotes.columns[7]: 'Data Última Compra'}, inplace=True)
            df_lotes['Data Última Compra'] = pd.to_datetime(df_lotes['Data Última Compra'], format='%d/%m/%Y', errors='coerce')
            df_lotes = df_lotes[['Código Produto', 'Data Última Compra']].drop_duplicates()
            df_principal = pd.merge(df_principal, df_lotes, on='Código Produto', how='left')

            gerar_relatorio_status("Coluna 'Data Última Compra' adicionada com sucesso.")
        else:
            gerar_relatorio_status("Erro: Coluna H ('Data Última Compra') não encontrada no arquivo.", sucesso=False)

    except Exception as e:
        gerar_relatorio_status(f"Erro ao adicionar coluna 'Data Última Compra': {e}", sucesso=False)

    return df_principal
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def carregar_marcas_vendas(diretorio):
    df = import_report(diretorio, 'TI-PBI_VendasProdutos.txt')
    if df.empty:
        gerar_relatorio_status("Arquivo TI-PBI_VendasProdutos.txt está vazio ou não encontrado.", sucesso=False)    # feat: Implementa cálculo detalhado da curva ABC baseada em quantidade de meses
        return pd.DataFrame()                                                                                       # Filtra as colunas dos meses relevantes para o cálculo da curva ABC.
    try:                                                                                                            # Calcula a soma total da quantidade dos meses selecionados e armazena em uma
        if 'Código Prod/Serv' in df.columns and 'Marca' in df.columns:                                              # nova coluna 'Total Quantidade Meses'
            df.rename(columns={'Código Prod/Serv': 'Código Produto'}, inplace=True)                                 # Ordena o DataFrame com base na quantidade total em ordem decrescente para
            df['Código Produto'] = df['Código Produto'].astype(str).str.strip()                                     # priorizar os itens mais relevantes.
            df['Marca'] = df['Marca'].astype(str).str.strip()                                                       # Calcula a quantidade acumulada em relação ao total e armazena em uma coluna
            df_marcas = df[['Código Produto', 'Marca']].drop_duplicates()                                           #  'Qtd Acumulado'
            gerar_relatorio_status("Marcas carregadas do arquivo TI-PBI_VendasProdutos.txt com sucesso.")
            return df_marcas
        else:
            gerar_relatorio_status("Colunas necessárias ('Código Prod/Serv' e 'Marca') não encontradas no arquivo TI-PBI_VendasProdutos.txt.", sucesso=False)
            return pd.DataFrame()

    except Exception as e:
        gerar_relatorio_status(f"Erro ao processar marcas do arquivo TI-PBI_VendasProdutos.txt: {e}", sucesso=False)
        return pd.DataFrame()
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#    
def carregar_marcas_estoque(diretorio):
    df = import_report(diretorio, 'CF_TI-PBI_Estoque.txt')
    if df.empty:
        gerar_relatorio_status("Arquivo CF_TI-PBI_Estoque.txt está vazio ou não encontrado.", sucesso=False)
        return pd.DataFrame()

    try:
        if 'Código Produto' in df.columns and 'Marca' in df.columns:
            df['Código Produto'] = df['Código Produto'].astype(str).str.strip()
            df['Marca'] = df['Marca'].astype(str).str.strip()

            df_marcas_estoque = df[['Código Produto', 'Marca']].drop_duplicates()
            gerar_relatorio_status("Marcas carregadas do arquivo CF_TI-PBI_Estoque.txt com sucesso.")
            return df_marcas_estoque
        else:
            gerar_relatorio_status("Colunas necessárias ('Código Produto' e 'Marca') não encontradas no arquivo CF_TI-PBI_Estoque.txt.", sucesso=False)
            return pd.DataFrame()

    except Exception as e:
        gerar_relatorio_status(f"Erro ao processar marcas do arquivo CF_TI-PBI_Estoque.txt: {e}", sucesso=False)
        return pd.DataFrame()
    
def executar_relatorio_completo(diretorio):
    df_produtos = carregar_planilhas(diretorio)
    df_sales = df_vendas(diretorio)
    df_estoque = carregar_estoque(diretorio)
    df_order_data = df_order(diretorio)
    df_lotes = validades(diretorio)
    df_marcas_vendas = carregar_marcas_vendas(diretorio)
    df_marcas_estoque = carregar_marcas_estoque(diretorio)

    # Verifica se todas as tabelas estão presente
    if df_sales.empty or df_estoque.empty or df_produtos.empty or df_marcas_vendas.empty or df_marcas_estoque.empty:
        gerar_relatorio_status("Dados insuficientes para gerar o relatório.", sucesso=False)
        return False

    # Julta todas as tabelas
    df_completo = completar_dados(
        df_sales, df_estoque, df_produtos, df_order_data, df_lotes, df_marcas_vendas, df_marcas_estoque
    )
    df_completo = adicionar_coluna_total(df_completo)
    df_completo = adicionar_coluna_media(df_completo)
    df_final = calcular_curva_abc(df_completo)

    # Selecionar colunas finais para o relatório
    colunas_finais = (
        ['Código Produto', 'Nome Produto', 'Marca', 'Inativo', 'Em Trânsito', 'Qtd. Lotes', 'Validade', 'Ordem de Compra',
         'Qtd a Receber', 'Saldo Total', 'Tempo de Estoque', 'Comprar(S/N)', 'Unidade de Compra', 'Data Última Compra'] +
        ['ABC Qtd', 'ABC Valor', 'ABC MARCA Qtd'] + meses_exibicao +
        ['Media', 'Total', 'Qtde Disponível']
    )
    colunas_finais = [col for col in colunas_finais if col in df_final.columns]
    df_final = df_final[colunas_finais]

    # Gerar o arquivo final
    caminho_saida = f'C:/Users/venda/Desktop/Relatorio_{arquivo}.xlsx'
    try:
        df_final.to_excel(caminho_saida, index=False)
        gerar_relatorio_status(f"Relatório gerado e salvo em {caminho_saida}.")
        estilizar_planilha(caminho_saida)
    except Exception as e:
        gerar_relatorio_status(f"Erro ao salvar o relatório: {e}", sucesso=False)
        return False

    # Resumo do status de execução
    print("\nResumo do Relatório de Status:")
    for item in status_relatorio:
        print(item)

    return True
