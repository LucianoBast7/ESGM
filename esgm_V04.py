# Importar Bibliotecas
print("IMPORTANDO BIBLIOTECAS")
import os
import pandas as pd
import pyodbc
import warnings
warnings.filterwarnings("ignore")
from variaveis.variaveis_esgm import Variaveis
from queries.queries_esgm import query_sinqia, query_data_atual_carteira_sinqia, query_cadastro_cotista
from datetime import datetime, date, timedelta
import numpy as np
from requirements.funcoes_selenium import SeleniumAutomator
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import base64
import logging
import zipfile
import shutil
print("BIBLIOTECAS IMPORTADAS")
time.sleep(1)

# Carregando Variáveis
print("CARREGANDO VARIÁVEIS")
time.sleep(1)

# Carregando classe com Variáveis
variaveis = Variaveis()

# Acessos Sinqia
server_sinqia = variaveis.server_sinqia
database_sinqia = variaveis.database_sinqia
user_sinqia = variaveis.user_sinqia
password_sinqia = variaveis.password_sinqia
lista_sinqia = variaveis.lista_sinqia
lista_data_carteira_atual = variaveis.lista_data_carteira_atual

# Url do Sinqia
url_sinqia = variaveis.url

# Parâmetros de Data
hoje = datetime.today()
data_atual = date.today()
ano = hoje.strftime("%Y")
mes = hoje.strftime("%m")
dia = hoje.strftime("%d")

caminho_log = f"W:\\PY_016 - ESGM\\0. Log\\{ano}\\{mes}\\{dia}\\ESGM_{data_atual}.log"

log_format = '%(asctime)s:%(levelname)s:%(filename)s:%(message)s'
logging.basicConfig(filename=caminho_log, filemode='a', level=logging.INFO, format=log_format)

def add_log(msg: str, level: str = 'info'):
    getattr(logging, level)(msg)

# Caminho arquivo ESGM
caminho_esgm = f"W:\\PY_016 - ESGM\\97. ESGM\\{ano}\\{mes}\\{dia}\\"
arquivo_comparacao = f"{caminho_esgm}comparacao_sinqia_esgm.csv"
arquivo_comparacao_cpf_novos = f"{caminho_esgm}comparacao_sinqia_esgm_cpf_novos.csv"
arquivo_comparacao_cpf_novos_excel = f"{caminho_esgm}conciliacao_sinqia_esgm (CPF NOVOS).xlsx"
arquivo_comparacao_excel_fake = f"{caminho_esgm}conciliacao_sinqia_esgm (FAKE).xlsx"
arquivo_comparacao_excel_real = f"{caminho_esgm}conciliacao_sinqia_esgm (REAL).xlsx"
reconciliacao = f"{caminho_esgm}Reconciliacao.xlsx"
cpfs_novos_sem_merge = f"{caminho_esgm}InconsistenciaCpfNovos.xlsx"
cpfs_novos_com_merge = f"{caminho_esgm}ReconciliacaoCpfNovos.xlsx"

# Caminho arquivo de Compra e Venda
arquivo_sinqia_esgm_fake = f"{caminho_esgm}TransferenciaTitularidade - FAKE.txt"
arquivo_sinqia_esgm_real = f"{caminho_esgm}TransferenciaTitularidade - REAL.txt"
arquivo_sinqia_esgm_cpf_novos = f"{caminho_esgm}TransferenciaTitularidade - CPF NOVOS.txt"

# Parâmetros String Formatada
cotista_fake = "9346601"
padrao = "0 S N"
data = "31/01/2025"

# Temporais
tempo_muito_longo = 12000

# Acessos Sinqia
user_sinqia_importar = variaveis.user_sinqia_importar
senha = variaveis.senha_sinqia_importar
decoded_bytes = base64.b64decode(senha)
senha_sinqia = decoded_bytes.decode("utf-8")

print("VARIÁVEIS CARREGADAS")
time.sleep(1)

# Conexão com Base de Dados
def connection_string(server, database, user, password):
    # Connection String
    connection = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={user};PWD={password};Encrypt=yes;TrustServerCertificate=yes;"
    conn = pyodbc.connect(connection)
    return conn

print("INICIANDO EXECUÇÃO")
add_log("INICIANDO EXECUÇÃO")
time.sleep(1)

def comparacao_sinqia_esgm():
    # Le o arquivo ESGM do BPA005
    print("LENDO ESGM")
    add_log("LENDO ESGM")
    time.sleep(1)
    def ler_esgm():
        arquivos = [f for f in os.listdir(caminho_esgm) if "ESGM" in f and f.endswith(".TXT")]
        if arquivos:
            caminho = os.path.join(caminho_esgm, arquivos[0])
        try:
            df = pd.read_csv(caminho, delimiter=";", encoding="latin1", engine="python", on_bad_lines="skip")
            return df
        except UnicodeDecodeError:
            print("ERRO DE CODIFICAÇÃO! TENTANDO OUTRA...")
            add_log("ERRO DE CODIFICAÇÃO! TENTANDO OUTRA...")
            try:
                df = pd.read_csv(caminho, delimiter=";", encoding="ISO-8859-1", engine="python", on_bad_lines="skip")
                return df
            except UnicodeDecodeError:
                print("NENHUMA CODIFICAÇÃO FUNCIONOU!")
                add_log("NENHUMA CODIFICAÇÃO FUNCIONOU!")

    arquivo_esgm = ler_esgm()
    print("LEITURA CONCLUIDA!")
    add_log("LEITURA CONCLUIDA!")
    time.sleep(1)

    print("TRATANDO DADOS ESGM")
    add_log("TRATANDO DADOS ESGM")
    time.sleep(1)
    def tratamento_dados_esgm():
        df_inicial = pd.DataFrame(columns=[
            "CPF/CNPJ - Cotistas - ESGM",
            "CPF - CNPJ - Cotistas (TESTE)",
            "ISIN do Fundo (TESTE)",
            "ISIN do Fundo",
            "Qtde de Cotas (TESTE)",
            "Qtde de Cotas",
            "Gravame"
        ])

        if arquivo_esgm is None or arquivo_esgm.empty:
            print("ERRO: O ARQUIVO ESGM NÃO FOI CARREGADO CORRETAMENTE OU ESTÁ VAZIO")
            add_log("ERRO: O ARQUIVO ESGM NÃO FOI CARREGADO CORRETAMENTE OU ESTA VAZIO")
            return pd.DataFrame()
        
        df_final = pd.concat([df_inicial, arquivo_esgm], axis=1)

        df_final["ISIN do Fundo (TESTE)"] = df_final.iloc[:, 7].astype(str).str[-136:]
        df_final["Qtde de Cotas (TESTE)"] = df_final.iloc[:, 7].astype(str).str[-102:]

        # Calculando formulas com pandas
        df_final["CPF - CNPJ - Cotistas (TESTE)"] = df_final.iloc[:, 7].astype(str).str.slice(2, 17)
        df_final["CPF/CNPJ - Cotistas - ESGM"] = pd.to_numeric(df_final["CPF - CNPJ - Cotistas (TESTE)"], errors="coerce")
        df_final["ISIN do Fundo"] = df_final["ISIN do Fundo (TESTE)"].astype(str).str[:12]
        df_final["Qtde de Cotas"] = pd.to_numeric(df_final["Qtde de Cotas (TESTE)"].astype(str).str[:15], errors="coerce")
        df_final["Gravame"] = df_final.iloc[:, 7].astype(str).str.slice(418, 421)
        df_final["CPF/CNPJ - Cotistas - ESGM"] = df_final["CPF/CNPJ - Cotistas - ESGM"].astype(str)

        df_final = df_final[
            [
                "CPF/CNPJ - Cotistas - ESGM",
                "ISIN do Fundo",
                "Qtde de Cotas",
                "Gravame"
            ]
        ]

        df_final["CPF/CNPJ - Cotistas - ESGM"] = df_final["CPF/CNPJ - Cotistas - ESGM"].fillna(0).replace([np.inf, -np.inf], 0).astype(str).str.replace(r"\.0$", "", regex=True)
        df_final["Qtde de Cotas"] = df_final["Qtde de Cotas"].fillna(0).astype(float).astype(int)

        return df_final
    
    arquivo_esgm = tratamento_dados_esgm()
    if not arquivo_esgm.empty:
        print("ARQUIVO ESGM TRATADO COM SUCESSO!")
        add_log("ARQUIVO ESGM TRATADO COM SUCESSO!")
        time.sleep(1)

    # Seleciona as colunas filtradas
    print("FILTRANDO COLUNAS ESGM")
    add_log("FILTRANDO COLUNAS ESGM")
    time.sleep(1)
    def filtrando_colunas_esgm(df):
        df["Cotas ESGM"] = df.groupby(["CPF/CNPJ - Cotistas - ESGM", "ISIN do Fundo"])["Qtde de Cotas"].transform("sum")
        df = df.drop_duplicates(subset=["CPF/CNPJ - Cotistas - ESGM", "ISIN do Fundo"], keep="first").copy()
        df = df[["CPF/CNPJ - Cotistas - ESGM", "ISIN do Fundo", "Cotas ESGM", "Gravame"]]       
        return df

    df_esgm = filtrando_colunas_esgm(arquivo_esgm)
    add_log("FILTRAGEM CONCLUIDA")
    time.sleep(1)

    # cpf_list = df_esgm["CPF/CNPJ - Cotistas - ESGM"].replace("nan", np.nan).dropna().astype(str).tolist()

    print("EXECUTANDO QUERY SINQIA")
    add_log("EXECUTANDO QUERY SINQIA")
    time.sleep(1)
    def executa_query():
        query = query_sinqia(lista_sinqia)
        conn = connection_string(server_sinqia, database_sinqia, user_sinqia, password_sinqia)
        cotas_df_sinqia = pd.read_sql(query, conn)
        conn.close()
        return cotas_df_sinqia
    
    def executar_query_data_carteira():
        query = query_data_atual_carteira_sinqia(lista_data_carteira_atual)
        conn = connection_string(server_sinqia, database_sinqia, user_sinqia, password_sinqia)
        df_data_atual_carteira = pd.read_sql(query, conn)
        conn.close()
        return df_data_atual_carteira
    
    def formatar_df():
        df_data_atual_carteira = executar_query_data_carteira()
        df_data_atual_carteira = df_data_atual_carteira[[
            "Código do Fundo",
            "DataAtual"
        ]]
        return df_data_atual_carteira
    
    df_data_atual_carteira = formatar_df()        
    
    def analise_sinqia():
        cotas_df_sinqia = executa_query()
        print("QUERY EXECUTADA COM SUCESSO")
        add_log("QUERY EXECUTADA COM SUCESSO")
        time.sleep(1)
        print("ANALISANDO DF SINQIA")
        add_log("ANALISANDO DF SINQIA")
        time.sleep(1)
        cotas_df_sinqia = cotas_df_sinqia[
            [
                "Carteira",
                "SaldoCotasAtual",
                "IDCliente",
                "Fundo",
                "CodigoCliente",
                "Código ISIN",
            ]
        ]
        return cotas_df_sinqia

    cotas_df_sinqia = analise_sinqia()
    print("ANALISE CONCLUIDA")
    add_log("ANALISE CONCLUIDA")
    time.sleep(1)

    print("ANALISANDO SINQIA FOUND")
    add_log("ANALISANDO SINQIA FOUND")
    time.sleep(1)
    def analise_sinqia_found():
        if not cotas_df_sinqia.empty:
            cotas_investidor_sinqia = pd.DataFrame(cotas_df_sinqia)
            cotas_investidor_sinqia.rename(
                columns={"SaldoCotasAtual": "SaldoCotas_Sinqia"}, inplace=True
            )
            return cotas_investidor_sinqia 
        else:
            print("NÃO EXISTE COTAS POR INVESTIDOR")
            add_log("NÃO EXISTEM COTAS POR INVESTIDOR")
            time.sleep(1)

    cotas_investidor_sinqia = analise_sinqia_found()
    print("ANALISE CONCLUIDA")
    add_log("ANALISE CONCLUIDA")
    time.sleep(1)

    print("FILTRANDO COTAS_INVESTIDOR_SINQIA")
    add_log("FILTRANDO COTAS_INVESTIDOR_SINQIA")
    time.sleep(1)
    def filtro_cotas_investidor_sinqia(df):
        df["Cotas Sinqia"] = df.groupby(["IDCliente", "Código ISIN", "Carteira"])["SaldoCotas_Sinqia"].transform("sum")
        df = df[
            [
                "Carteira",
                "IDCliente",
                "Cotas Sinqia",
                "Código ISIN",
                "CodigoCliente",
            ]
        ]
        df["Cotas Sinqia"] = df["Cotas Sinqia"].fillna(0).astype(float).astype(int)
        return df
        
    cotas_investidor_sinqia = filtro_cotas_investidor_sinqia(cotas_investidor_sinqia)
    print("COTAS_INESTIDOR_SINQIA PRONTO PARA MERGE")
    add_log("COTAS_INVESTIDOR_SINQIA PRONTO PARA MERGE")
    time.sleep(1)

    print("INICIANDO CONCILIAÇÃO SINQIA X ESGM")
    add_log("INICIANDO CONCILIAÇÃO SINQIA X ESGM")
    time.sleep(1)
    def juncao_sinqia_x_esgm():
        try:
            # Verifica se os DataFrames não estão vazios
            if not cotas_df_sinqia.empty and not df_esgm.empty:
                
                # Garantir que as colunas de chave primária estão no mesmo formato (string)
                cotas_investidor_sinqia["IDCliente"] = (
                    cotas_investidor_sinqia["IDCliente"]
                    .fillna("")
                    .astype(str)
                    .str.replace(r'\.0$', '', regex=True)
                    .str.strip()
                )

                df_isin_carteira = cotas_investidor_sinqia[[
                    "Carteira",
                    "Código ISIN"
                ]]
                df_isin_carteira = df_isin_carteira.drop_duplicates(subset=["Carteira", "Código ISIN"])
                
                df_esgm["CPF/CNPJ - Cotistas - ESGM"] = (
                    df_esgm["CPF/CNPJ - Cotistas - ESGM"]
                    .astype(str)
                    .str.strip()
                )

                # Merge dos DataFrames
                df_final = pd.merge(
                    cotas_investidor_sinqia,
                    df_esgm,
                    how="outer",
                    left_on=["IDCliente", "Código ISIN"],
                    right_on=["CPF/CNPJ - Cotistas - ESGM", "ISIN do Fundo"]
                )

                # Remover duplicatas
                df_final = df_final.drop_duplicates()

                # Substituir valores NaN
                df_final.fillna({"Cotas Sinqia": 0, "Cotas ESGM": 0}, inplace=True)

                # Calcular diferenças
                df_final["Diferença de cotas"] = df_final["Cotas Sinqia"] - df_final["Cotas ESGM"]

                df_final["Diferença de CPF/CNPJ"] = df_final.apply(
                    lambda row: "CORRETO" if row["IDCliente"] == row["CPF/CNPJ - Cotistas - ESGM"] else "INCORRETO",
                    axis=1
                )

                df_final_correto = df_final.copy()
                df_final_incorreto = df_final.copy()

                # Filtrar apenas registros corretos
                df_final_correto = df_final_correto[df_final_correto["Diferença de CPF/CNPJ"] == "CORRETO"]
                df_final_atualizado = pd.merge(
                    df_final_correto,
                    df_data_atual_carteira,
                    how="outer",
                    left_on="Carteira",
                    right_on="Código do Fundo"
                )
                # df_final_atualizado = df_final_atualizado[df_final_atualizado["Carteira"] == 200325.0]

                # Exportar resultado
                df_final_atualizado.to_csv(arquivo_comparacao, index=False) 
                df_reconciliacao = df_final_atualizado[[
                    "Código do Fundo",
                    "IDCliente",
                    "Código ISIN",
                    "CodigoCliente",
                    "Gravame",
                    "DataAtual",
                    "Cotas Sinqia",
                    "Cotas ESGM",
                    "Diferença de cotas"

                ]] 
                df_reconciliacao.to_excel(reconciliacao, index=False, sheet_name="Reconciliação")

                # Filtra dados de CPF/CNPJ novos e cria Lista para query
                def query_cpf_novos():
                    add_log("INICIANDO CONCILIAÇÃO SINQIA X CPF NOVOS ESGM")
                    df_cpf_novos = df_final_incorreto[df_final_incorreto["Diferença de CPF/CNPJ"] == "INCORRETO"]
                    cpf_list = df_cpf_novos["CPF/CNPJ - Cotistas - ESGM"].replace("nan", np.nan).dropna().astype(str).to_list()
                    
                    # Executar Query
                    add_log("EXECUTANDO QUERY SINQIA")
                    query = query_cadastro_cotista(cpf_list)
                    conn = connection_string(server_sinqia, database_sinqia, user_sinqia, password_sinqia)
                    df_lista_query = pd.read_sql(query, conn)
                    add_log("QUERY EXECUTADA COM SUCESSO")
                    
                    df_lista_query["IDCliente"] = (
                        df_lista_query["IDCliente"]
                        .astype(str)
                        .str.strip()
                        .str.replace(r"\.0$", "", regex=True)  # remove apenas se terminar com .0
                    )

                    # Merge com ESGM
                    df_cpf_novos_merge = pd.merge(
                        df_lista_query,
                        df_esgm,
                        how="outer",
                        left_on="IDCliente",
                        right_on="CPF/CNPJ - Cotistas - ESGM"                        
                    )

                    # Verificação de CPF/CNPJ
                    df_cpf_novos_merge["Diferença de CPF/CNPJ"] = df_cpf_novos_merge.apply(
                        lambda row: "CORRETO" if row["IDCliente"] == row["CPF/CNPJ - Cotistas - ESGM"] else "INCORRETO",
                        axis=1
                    )
                    
                    # Definição de Cpf sem cadastro no Sinqia
                    df_cpf_novos_incorreto = df_cpf_novos_merge[df_cpf_novos_merge["Diferença de CPF/CNPJ"] == "INCORRETO"]
                    df_cpf_novos_incorreto = df_cpf_novos_incorreto[[
                        "CPF/CNPJ - Cotistas - ESGM",
                        "ISIN do Fundo",
                        "Cotas ESGM",
                        "Gravame"
                    ]]
                    df_cpf_novos_incorreto.to_excel(cpfs_novos_sem_merge, index=False, sheet_name="Inconsistências ESGM")

                    # Definição de Cpf com cadastro no Sinqia
                    df_cpf_novos_correto = df_cpf_novos_merge[df_cpf_novos_merge["Diferença de CPF/CNPJ"] == "CORRETO"]
                    df_cpf_novos_correto = df_cpf_novos_correto[[
                        "IDCliente",
                        "CodigoCliente",
                        "ISIN do Fundo",
                        "Cotas ESGM",
                        "Gravame"
                    ]].astype(str)
                    df_cpf_novos_correto["ISIN do Fundo"] = df_cpf_novos_correto["ISIN do Fundo"].fillna("").astype(str).str.strip()
                    df_cpf_novos_correto = df_cpf_novos_correto[df_cpf_novos_correto["ISIN do Fundo"] != ""]

                    # Merge com cotas Cotas Investidor Sinqia ISIN
                    df_cpf_novos_final = pd.merge(
                        df_cpf_novos_correto,
                        df_isin_carteira,
                        how="outer",
                        left_on="ISIN do Fundo",
                        right_on="Código ISIN"
                    )

                    df_cpf_novos_final = df_cpf_novos_final[[
                        "Carteira",
                        "IDCliente",
                        "CodigoCliente",
                        "ISIN do Fundo",
                        "Cotas ESGM",
                        "Gravame"
                    ]].astype(str)
                    df_cpf_novos_final = df_cpf_novos_final.dropna()
                    df_cpf_novos_final["Carteira"] = df_cpf_novos_final["Carteira"].fillna("").astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                    df_data_atual_carteira["Código do Fundo"] = df_data_atual_carteira["Código do Fundo"].astype(str)

                    # Merge com Data Atual
                    df_cpfs_novos_final = pd.merge(
                        df_cpf_novos_final,
                        df_data_atual_carteira,
                        how="outer",
                        left_on="Carteira",
                        right_on="Código do Fundo"
                    )
                    
                    # Removendo carteiras que não existem na query
                    # Remove linhas onde a Carteira está como NaN ou em branco
                    try:
                        add_log("EXPORTANDO ARQUIVOS")
                        df_cpfs_novos_final["Carteira"] = df_cpfs_novos_final["Carteira"].astype(str).str.strip()
                        df_cpfs_novos_final = df_cpfs_novos_final[~df_cpfs_novos_final["Carteira"].isin(["", "nan", "NaN"])]
                        df_cpfs_novos_final["IDCliente"] = df_cpfs_novos_final["IDCliente"].astype(str).str.strip()
                        df_cpfs_novos_final = df_cpfs_novos_final[~df_cpfs_novos_final["IDCliente"].isin(["", "nan", "NaN"])]
                        df_cpfs_novos_final = df_cpfs_novos_final.drop_duplicates()
                        df_cpfs_novos_final.to_csv(arquivo_comparacao_cpf_novos, index=False)
                        df_exportacao_cpf_novos = df_cpfs_novos_final[[
                            "Carteira",
                            "IDCliente",
                            "CodigoCliente",
                            "ISIN do Fundo",
                            "Cotas ESGM",
                            "Gravame",
                            "DataAtual"
                        ]]
                        # df_exportacao_cpf_novos.to_excel(cpfs_novos_com_merge, index=False, sheet_name="Conciliação")
                        add_log("ARQUIVOS EXPORTADOS COM SUCESSO")
                    except Exception as e:
                        add_log(f"ERRO AO EXPORTAR ARQUIVOS: {e}")               
                    conn.close()
                
                # query_cpf_novos()                         

                print("BATIMENTO CONCLUÍDO!")
                add_log("BATIMENTO CONCLUIDO")
                time.sleep(1)
            else:
                print("NÃO FOI POSSÍVEL FAZER O BATIMENTO. UM DOS DF ESTA VAZIO.")     
                add_log("NÃO FOI POSSÍVEL FAZER O BATIMENTO. UM DOS DF ESTA VAZIO")
        except Exception as e:
            print(f"ERRO: {e}")

    juncao_sinqia_x_esgm()

comparacao_sinqia_esgm()

print("INICIANDO COMPRA E VENDA DE COTAS")
add_log("INICIANDO COMPRA E VENDA DE COTAS")
time.sleep(1)
def venda_compra_cotas():

    def remover_ponto_zero(valor):
        valor_str = str(valor)
        return valor_str[:-2] if valor_str.endswith(".0") else valor_str
    
    def formatar_coluna_txt(row):
        return (
            (" " * (10 - len(remover_ponto_zero(row["Carteira"])))) + remover_ponto_zero(row["Carteira"])
            + " " + row["Data"]
            + " " + (" " * (15 - len(remover_ponto_zero(row["De"])))) + remover_ponto_zero(row["De"])
            + " " + (" " * (15 - len(remover_ponto_zero(row["Para"])))) + remover_ponto_zero(row["Para"])
            + " 0 S N"
            + (" " * (15 - len(remover_ponto_zero(row["Cota"])))) + remover_ponto_zero(row["Cota"])
            + (" " * (18 - len(remover_ponto_zero(row["Qtde"])))) + remover_ponto_zero(row["Qtde"])
            + "0       "
        )

    def cria_arquivo():
        try:
            try:
                # Lê o arquivo csv com a comparação entre Sinqia x ESGM
                df_final = pd.read_csv(arquivo_comparacao, dtype=str).fillna("")
                excel_real = []
                excel_fake = []
                print("ARQUIVO CONCILIAÇÃO LIDO")
                add_log("ARQUIVO CONCILIAÇÃO LIDO")
            except Exception as e:
                print(f"ERRO AO LER ARQUIVO{e}")
                add_log("ERRO AO LER ARQUIVO")

            for _, row in df_final.iterrows():
                carteira = remover_ponto_zero(row["Carteira"])  
                id_cliente = remover_ponto_zero(row["CodigoCliente"])
                data_atual = row["DataAtual"]
                cota = "0"
 
                # Obtém valores de cotas e converte corretamente
                cotas_sinqia_num = int(float(row["Cotas Sinqia"])) if row["Cotas Sinqia"] else 0
                cotas_esgm_num = int(float(row["Cotas ESGM"])) if row["Cotas ESGM"] else 0

                if cotas_sinqia_num > cotas_esgm_num:
                    diferenca_final = remover_ponto_zero(cotas_sinqia_num - cotas_esgm_num) + "00000000"
                    excel_real.append({
                        "Carteira": carteira,
                        "Data": data_atual,
                        "De": id_cliente,
                        "Para": cotista_fake,
                        "Cota": cota,
                        "Qtde": diferenca_final
                    })

                elif cotas_sinqia_num < cotas_esgm_num:
                    diferenca_final = remover_ponto_zero(cotas_esgm_num - cotas_sinqia_num) + "00000000"
                    excel_fake.append({
                        "Carteira": carteira,
                        "Data": data_atual,
                        "De": cotista_fake,
                        "Para": id_cliente,
                        "Cota": cota,
                        "Qtde": diferenca_final
                    })

            try:
                # Criar DataFrames e gerar Excel
                df_excel_fake = pd.DataFrame(excel_fake)
                df_excel_real = pd.DataFrame(excel_real)
                print("DATAFRAME DE COMPRA E VENDA CRIADO COM SUCESSO")
                add_log("DATAFRAME DE COMPRA E VENDA CRIADO COM SUCESSO")
            except Exception as e:
                print(f"ERRO AO CRIAR DATAFRAMES: {e}")
                add_log("ERRO AO CRIAR DATAFRAME")

            try:
                if not df_excel_fake.empty:
                    df_excel_fake["Colar TXT"] = df_excel_fake.apply(formatar_coluna_txt, axis=1)
                    print("COLUNA (COLAR TXT) FORMATADA COM SUCESSO")
                    add_log("COLUNA (COLAT TXT) FORMATADA COM SUCESSO")
                elif df_excel_fake.empty:
                    df_excel_fake.to_csv(arquivo_sinqia_esgm_fake, index=False, header=False)
                    add_log("SEM DIFERENÇAS PARA CONCILIAR, ARQUIVO VAZIO")
                
                if not df_excel_real.empty:
                    df_excel_real["Colar TXT"] = df_excel_real.apply(formatar_coluna_txt, axis=1)
                    print("COLUNA (COLAR TXT) FORMATADA COM SUCESSO")
                    add_log("COLUNA (COLAT TXT) FORMATADA COM SUCESSO")
                elif df_excel_fake.empty:
                    df_excel_real.to_csv(arquivo_sinqia_esgm_real, index=False, header=False)
                    add_log("SEM DIFERENÇAS PARA CONCILIAR, ARQUIVO VAZIO")
            except Exception as e:
                print(f"ERRO AO FORMATAR COLUNA (COLAR TXT): {e}")
                add_log("ERRO AO FORMATAR COLUNA (COLAR TXT)")

            try:
                df_excel_fake.to_excel(arquivo_comparacao_excel_fake, index=False, sheet_name="Movimentações")
                df_excel_real.to_excel(arquivo_comparacao_excel_real, index=False, sheet_name="Movimentações")
                print("ARQUIVO EXCEL MOVIMENTAÇÕES EXPORTADOS COM SUCESSO")
                add_log("ARQUIVO EXCEL MOVIMENTAÇÕES EXPORTADOS COM SUCESSO")
            except Exception as e:
                print(f"ERRO AO TENTAR EXPORTAR MOVIMENTAÇÕES: {e}")
                add_log("ERRO AO TENTAR EXPORTAR MOVIMENTAÇÕES")

            try:
                if not df_excel_fake.empty:
                    # Exportar a coluna "Colar TXT" para arquivos .txt
                    df_excel_fake["Colar TXT"].to_csv(arquivo_sinqia_esgm_fake, index=False, header=False)
                    print("ARQUIVO TXT DE COMPRA E VENDA FINALIZADO COM SUCESSO")
                    add_log("ARQUIVO TXT DE COMPRA E VENDA FINZALIZADO COM SUCESSO")
                elif df_excel_fake.empty:
                    pass
                
                if not df_excel_real.empty:
                    df_excel_real["Colar TXT"].to_csv(arquivo_sinqia_esgm_real, index=False, header=False)
                    print("ARQUIVO TXT DE COMPRA E VENDA FINALIZADO COM SUCESSO")
                    add_log("ARQUIVO TXT DE COMPRA E VENDA FINZALIZADO COM SUCESSO")
                elif df_excel_real.empty:
                    pass
            except Exception as e:
                print(f"ERRO AO FINALIZAR FORMATAÇÃO DO ARQUIVO: {e}")
                add_log("ERRO AO FINALIZAR FORMATAÇÃO DO ARQUIVO")

            # Arquivos Finalizados
            print("COMPRA E VENDA DE COTAS FINALIZADA")
            add_log("COMPRA E VENDA DE COTAS FINALIZADA")

            def cria_arquivo_cpf_novos():
                add_log("CRIANDO ARQUIVO DE IMPORTAÇÃO COM CPF NOVOS")
                df_cpf_novos = pd.read_csv(arquivo_comparacao_cpf_novos, dtype=str).fillna("")
                df_cpf_novos = df_cpf_novos[df_cpf_novos["Cotas ESGM"] != "0"]
                df_cpf_novos = df_cpf_novos[df_cpf_novos["Carteira"] == "200325"]
                excel_cpf_novos = []

                for _, row in df_cpf_novos.iterrows():
                    carteira = remover_ponto_zero(row["Carteira"])
                    codigo_cleinte = remover_ponto_zero(row["CodigoCliente"])
                    data_atual_cpf_novos = row["DataAtual"]
                    cota_cpf_novos = "0"

                    cotas_esgm_cpf_novos = row["Cotas ESGM"]
                    
                    if cotas_esgm_cpf_novos != 0:
                        diferenca_final_cpf_novos = remover_ponto_zero(cotas_esgm_cpf_novos) + "00000000"
                        excel_cpf_novos.append({
                            "Carteira": carteira,
                            "Data": data_atual_cpf_novos,
                            "De": cotista_fake,
                            "Para": codigo_cleinte,
                            "Cota": cota_cpf_novos,
                            "Qtde": diferenca_final_cpf_novos
                        })
                try:
                    df_excel_cpf_novos = pd.DataFrame(excel_cpf_novos)
                    df_excel_cpf_novos["Colar TXT"] = df_excel_cpf_novos.apply(formatar_coluna_txt, axis=1)
                    df_excel_cpf_novos.to_excel(arquivo_comparacao_cpf_novos_excel, index=False)
                    df_excel_cpf_novos["Colar TXT"].to_csv(arquivo_sinqia_esgm_cpf_novos, index=False, header=False)
                    add_log("ARQUIVO EXPORTADO COM SUCESSO")
                except Exception as e:
                    add_log("ERRO AO CRIAR AQUIVO COM CPF NOVOS")
            # cria_arquivo_cpf_novos()

        except Exception as e:
            print(f"ERRO: {e}")
    cria_arquivo()
venda_compra_cotas()

# Inserir arquivo de Compra e Venda no Sinqia
def inserir_arquivo_sinqia():
    print("INSERINDO ARQUIVO DE COMPRA E VENDA NO SINQIA")
    add_log("INSERINDO ARQUIVO DE COMPRA E VENDA NO SINQIA")
    if __name__ == "__main__":
        bot = SeleniumAutomator()

        def importar_arquivo(arquivo, tipo):
            def selecionar_por_lista():
                try:
                    # Clicando em Selecionar por lista
                    print("CLICANDO EM 'SELECIONAR POR LISTA'")
                    add_log("CLICANDO EM 'SELECIONAR POR LISTA'")    
                    bot.clicar_por_xpath("//*[contains(text(), 'Selecionar por lista')]")
                    time.sleep(1)
                    print("CLICK REALIZADO")
                    add_log("CLICK REALIZADO")
                except Exception as e:
                    print("ERRO AO CLICAR EM 'SELECIONAR POR LISTA'")
                    add_log("ERRO AO CLICAR EM 'SELECIONAR POR LISTA'")

            selecionar_por_lista()

            def todas_as_carteiras():
                try:
                    # Clicar em Todas as carteiras
                    print("CLICANDO EM 'TODAS AS CARTEIRAS'")
                    add_log("cLICANDO EM 'TODAS AS CARTEIRAS'")
                    bot.clicar_por_id("chk0")
                    time.sleep(1)
                    bot.aguardar_estado_documento()
                    print("CLICK REALIZADO")
                    add_log("CLICK REALIZADO")
                except Exception as e:
                    print("ERRO AO CLICAR EM 'TODAS AS CARTEIRAS'")
                    add_log("ERRO AO CLICAR EM 'TODAS AS CARTEIRAS'")
            
            todas_as_carteiras()

            def lista_fii_esgm_passi():
                try:
                    # Selecionando lista
                    print("SELECIONANDO LISTA '#FII ESGM PASSI'")
                    add_log("SELECIONANDO LISTA '#FII ESGM PASSI'")
                    bot.executar_javascript('document.querySelector("#select_strIp_cbList-button > span.ui-selectmenu-text").click()')
                    time.sleep(1)
                    bot.aguardar_estado_documento()
                    bot.clicar_por_xpath("//div[contains(translate(text(), '\xa0', ' '), '#FII ESGM Passi')]")
                    time.sleep(1)
                    bot.aguardar_estado_documento()
                    print("LISTA SELECIONADA")
                    add_log("LISTA SELECIONADA")
                except Exception as e:
                    print("ERRO AO SELECIONAR LISTA")
                    add_log("ERRO AO SELECIONAR LISTA")

            lista_fii_esgm_passi()

            def transferencia_qtd():
                try:
                    # Selecionando tipo de tranferencia
                    print("SELECIONANDO ITEM 'TRANSFERENCIA QTD'")
                    add_log("SELECIONANDO ITEM 'TRANSFERENCIA QTD'")
                    bot.executar_javascript('document.querySelector("#select_cbProcesso-button > span.ui-selectmenu-text").click()')
                    time.sleep(1)
                    bot.aguardar_estado_documento()
                    bot.clicar_por_xpath("//div[contains(translate(text(), '\xa0', ' '), 'Transferência QTD')]")
                    time.sleep(1)
                    bot.aguardar_estado_documento()
                    print("ITEM CLICADO")
                    add_log("ITEM CLICADO")
                except Exception as e:
                    print("ERRO AO SELECIONAR ITEM NA LISTA")
                    add_log("ERRO AO SELECIONAR ITEM NA LISTA")

            transferencia_qtd()

            def processar_arquivo():
                try:
                    # Processa envio do arquivo
                    print("PROCESSANDO ENVIO")
                    add_log("PROCESSANDO ENVIO")
                    bot.enviar_arquivo_por_id("iu_file", arquivo)
                    time.sleep(1)
                    bot.aguardar_estado_documento()
                    bot.clicar_por_xpath("//button[contains(translate(text(), '\xa0', ' '), 'Processar')]")
                    time.sleep(5)
                    WebDriverWait(bot.driver, tempo_muito_longo).until(
                        EC.invisibility_of_element_located((By.CLASS_NAME, "ui-draggable"))
                    )
                    time.sleep(3)
                    print("ARQUIVO PROCESSADO")
                    add_log("ARQUIVO PROCESSADO")
                except Exception as e:
                    print("ERRO AO PROCESSAR O ARQUIVO")
                    add_log("ERRO AO PROCESSAR O ARQUIVO")

            processar_arquivo()

            def salvar_tela():
                try:
                    # Salvando tela de processamento
                    print("SALVANDO TELA DE PROCESSAMENTO")
                    add_log("SALVANDO TELA DE PROCESSAMENTO")
                    bot.trocar_para_nova_aba()
                    time.sleep(1)
                    # Baixar Documento de processamento dos dados
                    WebDriverWait(bot.driver, 60).until(
                        EC.presence_of_element_located((By.XPATH, "//b[contains(text(), 'Importação de Transferência de Titularidade')]"))
                    )
                    bot.driver.execute_script("window.print();")

                    time.sleep(5)

                    # Renomar Nome .pdf
                    nome_atual = "Importação de Transferência de Titularidade.pdf"
                    nome_novo = f"ImportacaoTransferenciaTitularidade - {tipo}.pdf"
                    origem = os.path.join(caminho_esgm, nome_atual)
                    destino = os.path.join(caminho_esgm, nome_novo)
                    if os.path.exists(origem):
                        os.rename(origem, destino)

                    bot.aguardar_estado_documento()
                    time.sleep(2)
                    bot.fechar_nova_aba()
                    time.sleep(1)
                    print("TELA SALVA")
                    add_log("TELA SALVA")
                except Exception as e:
                    print("ERRO AO SALVAR TELA")
                    add_log("ERRO AO SALVAR TELA")

            salvar_tela()

        try:
            def abrir_sinqia():
                try:
                    # Acessar site
                    print("ABRINDO SINQIA")
                    add_log("ABRINDO SINQIA")
                    bot.navegar_para(url_sinqia)
                    bot.aguardar_estado_documento()
                    bot.maximizar_janela()
                    print("SINQIA ABERTO COM SUCESSO")
                    add_log("SINQIA ABERTO COM SUCESSO")
                except Exception as e:
                    print("ERRO AO ACESSAR SINQIA")
                    add_log("ERRO AO ACESSAR SINQIA")
            
            abrir_sinqia()

            def realizar_login(user, senha):
                try:
                    # Inserir Usuário e Senha
                    print("FAZENDO LOGIN")
                    add_log("FAZENDO LOGIN")
                    bot.digitar_por_id("Usr", user)
                    time.sleep(1)
                    bot.digitar_por_id("Pwd", senha)
                    time.sleep(1)
                    bot.clicar_por_id("btnSubmit")
                    time.sleep(8)
                    bot.aguardar_estado_documento()
                    print("LOGIN REALIZADO COM SUCESSO")
                    add_log("LOGIN REALIZADO COM SUCESSO")
                except Exception as e:
                    print("ERRO AO FAZER LOGIN")
                    add_log("ERRO AO FAZER LOGIN")

            realizar_login(user_sinqia_importar, senha_sinqia)

            def acessar_processo():
                try:
                    # Acessando processo 1476
                    print("ACESSANDO PROCESSO 1476")
                    add_log("ACESSANDO PROCESSO 1476")
                    bot.digitar_por_id("txArgBusca", "1476")
                    time.sleep(1)
                    bot.executar_javascript('document.querySelector("#search-bar > table > tbody > tr > td:nth-child(2) > a > img").click()')
                    time.sleep(1)
                    bot.aguardar_estado_documento()
                    time.sleep(1)
                    print("PROCESSO ABERTO COM SUCESSO")
                    add_log("PROCESSO ABERTO COM SUCESSO")
                except Exception as e:
                    print("ERRO AO ACESSAR PROCESSO 1476")
                    add_log("ERRO AO ACESSAR PROCESSO 1476")

            acessar_processo()

            def importar_arquivo_venda():
                try:
                    # Importando Arquivo de Venda
                    print("IMPORTANDO ARQUIVO DE VENDA")
                    add_log("IMPORTANDO ARQUIVO DE VENDA")
                    importar_arquivo(arquivo_sinqia_esgm_real, "REAL")
                    time.sleep(1)
                    bot.voltar_para_aba_anterior()
                    time.sleep(1)
                    bot.aguardar_estado_documento()
                    time.sleep(1)
                    bot.driver.refresh()
                    time.sleep(1)
                    bot.aguardar_estado_documento()
                    time.sleep(1)
                    print("PROCESSO DE IMPORTAÇÃO CONCLUIDO")
                    add_log("PROCESSO DE IMPORTAÇÃO CONCLUIDO")
                except Exception as e:
                    print("ERRO AO IMPORTAR ARQUIVO DE VENDA")
                    add_log("ERRO AO IMPORTAR ARQUIVO DE VENDA")

            importar_arquivo_venda()

            def importar_arquivo_compra():
                try:
                    # Importando Arquivo de 
                    print("IMPORTANDO ARQUIVO DE COMPRA")
                    add_log("IMPORTANDO ARQUIVO DE COMPRA")
                    importar_arquivo(arquivo_sinqia_esgm_fake, "FAKE")
                    time.sleep(1)
                    bot.voltar_para_aba_anterior()
                    time.sleep(1)
                    bot.aguardar_estado_documento()
                    time.sleep(1)
                    bot.driver.refresh()
                    time.sleep(1)
                    bot.aguardar_estado_documento()
                    time.sleep(1)
                    print("PROCESSO DE IMPORTAÇÃO CONCLUIDO")
                    add_log("PROCESSO DE IMPORTAÇÃO CONCLUIDO")
                except Exception as e:
                    print("ERRO AO IMPORTAR ARQUIVO DE COMPRA")
                    add_log("ERRO AO IMPORTAR ARQUIVO DE COMPRA")

            importar_arquivo_compra()

            def importar_arquivo_cpf_novos():
                try:
                    # Importando Arquivo de 
                    print("IMPORTANDO ARQUIVO DE CPF NOVOS")
                    add_log("IMPORTANDO ARQUIVO DE CPF NOVOS")
                    importar_arquivo(arquivo_sinqia_esgm_cpf_novos, "CPF NOVOS")
                    time.sleep(1)
                    bot.voltar_para_aba_anterior()
                    time.sleep(1)
                    bot.aguardar_estado_documento()
                    time.sleep(1)
                    bot.driver.refresh()
                    time.sleep(1)
                    bot.aguardar_estado_documento()
                    time.sleep(1)
                    print("PROCESSO DE IMPORTAÇÃO CONCLUIDO")
                    add_log("PROCESSO DE IMPORTAÇÃO CONCLUIDO")
                except Exception as e:
                    print("ERRO AO IMPORTAR ARQUIVO DE CPF NOVOS")
                    add_log("ERRO AO IMPORTAR ARQUIVO DE CPF NOVOS")

            # importar_arquivo_cpf_novos()

            try:
                # Fecha o navegador
                bot.fechar_navegador()
                print("LOGOUT CONCLUIDO")
                add_log("LOGOUT CONCLUIDO")
            except Exception as e:
                bot.driver.quit()
        except Exception as e:
            print(f"ERRO NO PROCESSO DE IMPORTAÇÃO {e}")
            bot.fechar_navegador()

inserir_arquivo_sinqia()

# Iniciando Reconciliação
comparacao_sinqia_esgm()
    