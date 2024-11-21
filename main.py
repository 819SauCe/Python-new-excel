import os
import pandas as pd
from datetime import datetime
from estilização import estilizar_planilha
import numpy as np

meses_para_calculo = pd.date_range(end=datetime.now(), periods=36, freq='MS').strftime('%B, %Y').tolist()
meses_exibicao = meses_para_calculo[-12:]
arquivo = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
status_relatorio = []

def gerar_relatorio_status(mensagem, sucesso=True):
    status = 'Sucesso' if sucesso else 'Erro'
    status_relatorio.append(f"{status}: {mensagem}")
    if not sucesso:
        print(f"{status}: {mensagem}")

def classificar_abc(percentual, tipo='geral'):
    thresholds = {'geral': (0.20, 0.75), 'marca_valor': (0.35, 0.75), 'marca': (0.10, 0.20)}
    a, b = thresholds.get(tipo, (0, 1))
    return 'A' if percentual <= a else ('B' if percentual <= b else 'C')

def import_report(path, filename):
    try:
        df = pd.read_csv(os.path.join(path, filename), sep='|', encoding='ISO-8859-1', low_memory=False)
        gerar_relatorio_status(f"{filename} carregado com sucesso.")
        return df
    except Exception as e:
        gerar_relatorio_status(f"Erro ao ler {filename}: {e}", sucesso=False)
        return pd.DataFrame()

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

def validades(diretorio):
    df = import_report(diretorio, 'TI-PBI_EstoqueLote.txt')
    if df.empty:
        gerar_relatorio_status("Arquivo TI-PBI_EstoqueLote.txt está vazio ou não encontrado.", sucesso=False)
        return df

    try:
        # Garantir que 'Marca' está presente
        if 'Marca' in df.columns:
            df['Marca'] = df['Marca'].astype(str).str.strip()
        else:
            gerar_relatorio_status("Coluna 'Marca' não encontrada no arquivo TI-PBI_EstoqueLote.txt.", sucesso=False)
            df['Marca'] = None

        # Processar 'Disponível'
        if 'S' in df.columns:
            df['Disponível'] = pd.to_numeric(df['S'], errors='coerce').fillna(0).astype(int)

        # Processar 'Data Última Compra'
        if 'Data Última Compra' in df.columns:
            df['Data Última Compra'] = pd.to_datetime(df['Data Última Compra'], format='%d/%m/%y', errors='coerce')
        elif df.columns.size >= 8:
            df['Data Última Compra'] = pd.to_datetime(df.iloc[:, 7], errors='coerce')

        # Adicionar 'Qtd. Lotes'
        df['Qtd. Lotes'] = 1

        # Selecionar e agrupar colunas relevantes
        df = df[['Código Produto', 'Marca', 'Validade', 'Qtd. Lotes', 'Disponível', 'Data Última Compra']]
        df = df.groupby('Código Produto').agg({
            'Marca': 'first',
            'Validade': 'first',
            'Qtd. Lotes': 'sum',
            'Disponível': 'sum',
            'Data Última Compra': 'max'
        }).reset_index()

        gerar_relatorio_status("Dados de validades processados com sucesso.")
    except Exception as e:
        gerar_relatorio_status(f"Erro ao processar validades: {e}", sucesso=False)
        return pd.DataFrame()

    return df

def completar_dados(df_vendas, df_estoque, df_produtos, df_order, df_lotes):
    # Padronizar os códigos dos produtos
    df_vendas['Código Produto'] = df_vendas['Código Produto'].astype(str).str.strip().str.upper()
    df_estoque['Código Produto'] = df_estoque['Código Produto'].astype(str).str.strip().str.upper()
    df_produtos['Código Produto'] = df_produtos['Código Produto'].astype(str).str.strip().str.upper()
    df_order['Código Produto'] = df_order['Código Produto'].astype(str).str.strip().str.upper()
    df_lotes['Código Produto'] = df_lotes['Código Produto'].astype(str).str.strip().str.upper()

    # Combinar os DataFrames
    df_completo = pd.merge(df_estoque, df_vendas, on='Código Produto', how='outer')
    df_completo = pd.merge(df_completo, df_produtos, on='Código Produto', how='left', suffixes=('', '_produtos'))
    df_completo = pd.merge(df_completo, df_order, on='Código Produto', how='outer', suffixes=('', '_order'))
    df_completo = pd.merge(df_completo, df_lotes, on='Código Produto', how='outer', suffixes=('', '_lotes'))

    # Resolver colunas conflitantes
    if 'Em Trânsito_order' in df_completo.columns:
        df_completo['Em Trânsito'] = df_completo['Em Trânsito'].combine_first(df_completo['Em Trânsito_order'])
        df_completo.drop(columns=['Em Trânsito_order'], inplace=True)

    # Garantir que Marca e Data Última Compra estão presentes
    if 'Marca' in df_lotes.columns:
        df_completo['Marca'] = df_lotes['Marca']
    if 'Data Última Compra' in df_lotes.columns:
        df_completo['Data Última Compra'] = df_lotes['Data Última Compra']

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

    # Preencher dados dos meses de vendas
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


def adicionar_coluna_total(df):
    colunas_meses = [col for col in meses_exibicao if col in df.columns]
    df[colunas_meses] = df[colunas_meses].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
    df['Total'] = df[colunas_meses].sum(axis=1)
    return df

def adicionar_coluna_media(df):
    colunas_meses = [col for col in meses_exibicao if col in df.columns]
    # Calcular a média apenas para valores maiores que zero e arredondar para o inteiro mais próximo
    df['Media'] = df[colunas_meses].apply(lambda row: int(row[row > 0].mean()) if row[row > 0].count() > 0 else '-', axis=1)
    return df

def calcular_curva_abc(df):
    colunas_meses = [col for col in meses_para_calculo if col in df.columns]
    df['Total Quantidade Meses'] = df[colunas_meses].sum(axis=1)
    total_qtd = df['Total Quantidade Meses'].sum()

    if total_qtd != 0:
        df = df.sort_values(by='Total Quantidade Meses', ascending=False)
        df['Qtd Acumulado'] = df['Total Quantidade Meses'].cumsum() / total_qtd
        df['ABC Qtd'] = df['Qtd Acumulado'].apply(lambda x: classificar_abc(x, tipo='geral'))
        df['ABC Valor'] = df['Qtd Acumulado'].apply(lambda x: classificar_abc(x, tipo='marca_valor'))
        df['ABC MARCA Qtd'] = df['Qtd Acumulado'].apply(lambda x: classificar_abc(x, tipo='marca'))

    df = df.drop(columns=['Qtd Acumulado'], errors='ignore')
    gerar_relatorio_status("Curva ABC calculada com sucesso.")
    return df


def adicionar_data_ultima_compra(diretorio, df_principal):
    # Carregar o arquivo TI-PBI_EstoqueLote.txt
    df_lotes = import_report(diretorio, 'TI-PBI_EstoqueLote.txt')
    
    if df_lotes.empty:
        gerar_relatorio_status("Erro: Arquivo TI-PBI_EstoqueLote.txt não encontrado ou vazio.", sucesso=False)
        return df_principal

    try:
        # Garantir que existe a coluna H (índice 7)
        if df_lotes.columns.size >= 8:
            df_lotes.rename(columns={df_lotes.columns[7]: 'Data Última Compra'}, inplace=True)  # Coluna H = index 7
            df_lotes['Data Última Compra'] = pd.to_datetime(df_lotes['Data Última Compra'], format='%d/%m/%Y', errors='coerce')

            # Manter apenas Código Produto e Data Última Compra
            df_lotes = df_lotes[['Código Produto', 'Data Última Compra']].drop_duplicates()

            # Mesclar com a planilha principal
            df_principal = pd.merge(df_principal, df_lotes, on='Código Produto', how='left')
            gerar_relatorio_status("Coluna 'Data Última Compra' adicionada com sucesso.")
        else:
            gerar_relatorio_status("Erro: Coluna H ('Data Última Compra') não encontrada no arquivo.", sucesso=False)

    except Exception as e:
        gerar_relatorio_status(f"Erro ao adicionar coluna 'Data Última Compra': {e}", sucesso=False)

    return df_principal

def executar_relatorio_completo(diretorio):
    df_produtos = carregar_planilhas(diretorio)
    df_sales = df_vendas(diretorio)
    df_estoque = carregar_estoque(diretorio)
    df_order_data = df_order(diretorio)
    df_lotes = validades(diretorio)

    if df_sales.empty or df_estoque.empty or df_produtos.empty:
        gerar_relatorio_status("Dados insuficientes para gerar o relatório.", sucesso=False)
        return False

    df_completo = completar_dados(df_sales, df_estoque, df_produtos, df_order_data, df_lotes)
    df_completo = adicionar_coluna_total(df_completo)
    df_completo = adicionar_coluna_media(df_completo)
    df_final = calcular_curva_abc(df_completo)

    colunas_finais = (
        ['Código Produto', 'Nome Produto','Marca', 'Inativo', 'Em Trânsito', 'Qtd. Lotes', 'Validade', 'Ordem de Compra',
         'Qtd a Receber', 'Saldo Total', 'Tempo de Estoque', 'Comprar(S/N)', 'Unidade de Compra', 'Data Última Compra'] +
        ['ABC Qtd', 'ABC Valor', 'ABC MARCA Qtd'] + meses_exibicao +
        ['Media', 'Total', 'Qtde Disponível']
    )

    colunas_finais = [col for col in colunas_finais if col in df_final.columns]
    df_final = df_final[colunas_finais]

    caminho_saida = f'C:/Users/venda/Desktop/Relatorio_{arquivo}.xlsx'
    try:
        df_final.to_excel(caminho_saida, index=False)
        gerar_relatorio_status(f"Relatório gerado e salvo em {caminho_saida}.")
        estilizar_planilha(caminho_saida)
    except Exception as e:
        gerar_relatorio_status(f"Erro ao salvar o relatório: {e}", sucesso=False)
        return False

    print("\nResumo do Relatório de Status:")
    for item in status_relatorio:
        print(item)
    
    return True
