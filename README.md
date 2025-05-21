ESGM – Sistema de Conciliação e Importação Sinqia
📌 Visão Geral
Este projeto tem como objetivo automatizar o processo de reconciliação de dados entre os arquivos do ESGM (posição de cotistas) e os dados do sistema Sinqia, além de gerar e importar os arquivos de ajuste de cotas por meio da automação com Selenium.

O fluxo completo envolve:

Leitura e tratamento de arquivos TXT (posição de cotistas).

Comparação com base de dados SQL Server.

Geração de relatórios conciliatórios e arquivos de ajuste de cotas.

Envio automatizado dos arquivos para o sistema Sinqia via navegador.

🗂️ Estrutura dos Scripts
🔹 2036_V01.py
Realiza o envio manual do arquivo ESGM no processo 2036 do Sinqia via automação com Selenium.

🔹 esgm_V04.py
Processo completo:

Leitura e tratamento do arquivo ESGM.

Execução de queries SQL via pyodbc.

Conciliação das cotas com a base do Sinqia.

Geração de arquivos Excel e TXT para importação.

Importação dos arquivos gerados no processo 1476 do Sinqia.

🔧 Pré-requisitos
Python 3.9+

ODBC Driver 17 for SQL Server

Instalação das bibliotecas listadas em requirements.txt:

bash
Copiar
Editar
pip install pandas pyodbc selenium numpy python-dotenv
ChromeDriver compatível com a versão do navegador.

🔐 Variáveis de Ambiente
As credenciais e caminhos estão centralizados no módulo:

bash
Copiar
Editar
variaveis/variaveis_esgm.py
Certifique-se de configurar:

Credenciais do Sinqia

URL do sistema

Lista de carteiras

Informações do banco SQL Server

🚀 Execução
Importação manual (2036):

bash
Copiar
Editar
python 2036_V01.py
Processo completo de conciliação e importação (1476):

bash
Copiar
Editar
python esgm_V04.py
🗃️ Arquivos Gerados
conciliacao_sinqia_esgm (REAL).xlsx / FAKE.xlsx: Excel com movimentações detectadas.

TransferenciaTitularidade - REAL.txt / FAKE.txt: Arquivos formatados para importar no Sinqia.

Reconciliacao.xlsx: Relatório principal de divergências por fundo.

Logs por data no caminho:

css
Copiar
Editar
W:\PY_016 - ESGM\0. Log\{ano}\{mes}\{dia}\
📌 Observações
A reconciliação considera apenas os dados com CPF/CNPJ válidos e realiza merge por Código ISIN.

Há tratamento específico para cotistas novos não encontrados na base Sinqia.

O sistema inclui exportações para Excel e .txt com formatação customizada para integração Sinqia.

🤖 Tecnologias Utilizadas
Python

Selenium WebDriver

Pandas / NumPy

pyodbc (SQL Server)

Automação de Web UI
