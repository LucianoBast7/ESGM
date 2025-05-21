# Armazenamento de Variáveis de Ambiente

# Importação de Bibliotecas
import os
from dotenv import load_dotenv
import ast

class Variaveis():
    def __init__(self):
        
        # Carregar Variáveis .env
        load_dotenv()

        # Acessos Sinqia
        self.server_sinqia = os.getenv("SERVER_SINQIA")
        self.database_sinqia = os.getenv("DATABASE_SINQIA")
        self.user_sinqia = os.getenv("USER_SINQIA")
        self.password_sinqia = os.getenv("PASSWORD_SINQIA")

        # Listas
        self.lista_sinqia = os.getenv("LISTA_SINQIA")
        self.lista_data_carteira_atual = os.getenv("LISTA_DATA_CARTEIRA_ATUAL")

        # Acessos Sinqia Importar Arquivo
        self.senha_sinqia_importar = os.getenv("SENHA_SINQIA_IMPORTAR")
        self.user_sinqia_importar = os.getenv("USER_SINQIA_IMPORTAR")

        # Url do Sinqia
        self.url = os.getenv("URL")
