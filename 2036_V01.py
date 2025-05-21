import os
import pandas as pd
import warnings
warnings.filterwarnings("ignore")
from variaveis.variaveis_esgm import Variaveis
from datetime import datetime, date
from requirements.funcoes_selenium import SeleniumAutomator
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import base64
import logging
import glob

variaveis = Variaveis()
bot = SeleniumAutomator()

tempo_muito_longo = 12000

# Url do Sinqia
url_sinqia = variaveis.url

# Acessos Sinqia
user_sinqia_importar = variaveis.user_sinqia_importar
senha = variaveis.senha_sinqia_importar
decoded_bytes = base64.b64decode(senha)
senha_sinqia = decoded_bytes.decode("utf-8")

# Parâmetros de Data
hoje = datetime.today()
data_atual = date.today()
ano = hoje.strftime("%Y")
mes = hoje.strftime("%m")
dia = hoje.strftime("%d")

caminho_esgm = f"W:\\PY_016 - ESGM\\97. ESGM\\{ano}\\{mes}\\{dia}\\"
arquivos = [f for f in os.listdir(caminho_esgm) if "ESGM" in f and f.endswith(".TXT")]

if arquivos:
    nome_arquivo = arquivos[0]  # <-- nome do arquivo
    caminho = os.path.join(caminho_esgm, nome_arquivo)  # <-- caminho completo

arquivo = f"{caminho_esgm}{nome_arquivo}"

caminho_log = f"W:\\PY_016 - ESGM\\0. Log\\{ano}\\{mes}\\{dia}\\2036_{data_atual}.log"

log_format = '%(asctime)s:%(levelname)s:%(filename)s:%(message)s'
logging.basicConfig(filename=caminho_log, filemode='a', level=logging.INFO, format=log_format)

def add_log(msg: str, level: str = 'info'):
    getattr(logging, level)(msg)


def importar_esgm_2036():
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
            # Acessando processo 2036
            print("ACESSANDO PROCESSO 2036")
            add_log("ACESSANDO PROCESSO 2036")
            bot.digitar_por_id("txArgBusca", "2036")
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

    def check_box_7():
        try:
            # Clicar em Concilia Cadastro de Cliente
            print("CLICANDO EM 'CONCILIA CADASTRO DE CLIENTE'")
            add_log("cLICANDO EM 'CONCILIA CADASTRO DE CLIENTE'")
            bot.clicar_por_id("chk7")
            time.sleep(1)
            bot.aguardar_estado_documento()
            print("CLICK REALIZADO")
            add_log("CLICK REALIZADO")
        except Exception as e:
            print("ERRO AO CLICAR EM 'CONCILIA CADASTRO DE CLIENTE'")
            add_log("ERRO AO CLICAR EM 'CONCILIA CADASTRO DE CLIENTE'")

    check_box_7()

    def check_box_8():
        try:
            # Clicar em Sobrescreve Cadastro de Cliente
            print("CLICANDO EM 'SOBRESCREVE CADASTRO DE CLIENTE'")
            add_log("cLICANDO EM 'SOBRESCREVE CADASTRO DE CLIENTE'")
            bot.clicar_por_id("chk8")
            time.sleep(1)
            bot.aguardar_estado_documento()
            print("CLICK REALIZADO")
            add_log("CLICK REALIZADO")
        except Exception as e:
            print("ERRO AO CLICAR EM 'SOBRESCREVE CADASTRO DE CLIENTE'")
            add_log("ERRO AO CLICAR EM 'SOBRESCREVE CADASTRO DE CLIENTE'")

    check_box_8()

    # def check_box_9():
    #     try:
    #         # Clicar em Considerar Hist. de Saldos de Gravames
    #         print("CLICANDO EM 'CONSIDERAR HIST. DE SALDOS DE GRAVAMES'")
    #         add_log("cLICANDO EM 'CONSIDERAR HIST. DE SALDOS DE GRAVAMES'")
    #         bot.clicar_por_id("chk9")
    #         time.sleep(1)
    #         bot.aguardar_estado_documento()
    #         print("CLICK REALIZADO")
    #         add_log("CLICK REALIZADO")
    #     except Exception as e:
    #         print("ERRO AO CLICAR EM 'CONSIDERAR HIST. DE SALDOS DE GRAVAMES'")
    #         add_log("ERRO AO CLICAR EM 'CONSIDERAR HIST. DE SALDOS DE GRAVAMES'")

    # check_box_9()

    def check_box_11():
        try:
            # Clicar em Importae INR sem CPF/CNPJ informado
            print("CLICANDO EM 'IMPORTAR INR SEM CPF/CNPJ INFORMADO'")
            add_log("CLICANDO EM 'IMPORTAR INR SEM CPF/CNPJ INFORMADO'")
            bot.clicar_por_id("chk11")
            time.sleep(1)
            bot.aguardar_estado_documento()
            print("CLICK REALIZADO")
            add_log("CLICK REALIZADO")
        except Exception as e:
            print("ERRO AO CLICAR EM 'IMPORTAR INR SEM CPF/CNPJ INFORMADO'")
            add_log("ERRO AO CLICAR EM 'IMPORTAR INR SEM CPF/CNPJ INFORMADO'")

    check_box_11()

    def zerar_du():
        try:
            # Zerando DU
            print("ZERANDO DU")
            add_log("ZERANDO DU")
            bot.limpar_campo_por_xpath("//*[contains(@internalname, 'iiDULiqMov')]")
            bot.digitar_por_xpath("//*[contains(@internalname, 'iiDULiqMov')]", "0")
            time.sleep(1)
            bot.aguardar_estado_documento()
            print("DU ZERADA")
            add_log("DU ZERADA")
        except Exception as e:
            print("ERRO AO ZERAR DU")
            add_log("ERRO AO ZERAR DU")

    zerar_du()

    def lista_ESGM_1():
        try:
            # Selecionando Processo ESGM
            print("SELECIONANDO LISTA ESGM")
            add_log("SELECIONANDO LISTA ESGM")
            bot.executar_javascript('document.querySelector("#select_cbImpPosCot-button").click()')
            time.sleep(1)
            bot.aguardar_estado_documento()
            bot.clicar_por_xpath("//ul[@id='select_cbImpPosCot-menu']/li/div[text() = 'ESGM']")
            time.sleep(1)
            bot.aguardar_estado_documento()
            print("LISTA SELECIONADA")
            add_log("LISTA SELECIONADA")
        except Exception as e:
            print("ERRO AO SELECIONAR LISTA")
            add_log("ERRO AO SELECIONAR LISTA")

    lista_ESGM_1()

    def lista_ESGM_2():
        try:
            # Selecionando Processo ESGM
            print("SELECIONANDO LISTA ESGM")
            add_log("SELECIONANDO LISTA ESGM")
            bot.executar_javascript('document.querySelector("#select_cbImpCadClt-button").click()')
            time.sleep(1)
            bot.aguardar_estado_documento()
            bot.clicar_por_xpath("//ul[@id='select_cbImpCadClt-menu']/li/div[text() = 'ESGM']")
            time.sleep(1)
            bot.aguardar_estado_documento()
            print("LISTA SELECIONADA")
            add_log("LISTA SELECIONADA")
        except Exception as e:
            print("ERRO AO SELECIONAR LISTA")
            add_log("ERRO AO SELECIONAR LISTA")

    lista_ESGM_2()

    def processar_arquivo():
        try:
            # Processa envio do arquivo
            print("PROCESSANDO ENVIO")
            add_log("PROCESSANDO ENVIO")
            bot.enviar_arquivo_por_id("iuProcessar_file", arquivo)
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
            print(f"ERRO AO PROCESSAR O ARQUIVO: {e}")
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
                EC.presence_of_element_located((By.XPATH, "//b[contains(text(), 'Importação de Posição de Cotista - B3')]"))
            )
            bot.driver.execute_script("window.print();")
            # Renomar Nome .pdf
            time.sleep(5)
            nome_atual = "Importação de Posição de Cotista - B3.pdf"
            nome_novo = f"ImportacaoPosicaoCotista - B3.pdf"
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
        # Fecha o navegador
        bot.fechar_navegador()
        print("LOGOUT CONCLUIDO")
        add_log("LOGOUT CONCLUIDO")
    except Exception as e:
        bot.driver.quit()

importar_esgm_2036()