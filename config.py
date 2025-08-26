import os
from pathlib import Path
from datetime import datetime, time
from dotenv import load_dotenv


VIEWPORT_PADRAO = {'width': 1920, 'height': 1080}

CAMINHO_PASTA_RAIZ = Path(__file__).resolve().parent.parent

CAMINHO_PASTA_PRINTS = Path(CAMINHO_PASTA_RAIZ, 'Prints')

CAMINHO_PASTA_DOCX = Path(CAMINHO_PASTA_RAIZ, 'Histórico de monitoramentos')

CAMINHO_PASTA_LOGS = Path(CAMINHO_PASTA_RAIZ, 'Logs')

CAMINHO_PASTA_DADOS_MENSAIS = Path(CAMINHO_PASTA_RAIZ, 'Dados Mensais')


FORMATACAO_LOGGING = '%(asctime)s - %(name)s - %(levelname)s - %(message)s \n'


AGORA = datetime.now()

DATA_ATUAL = AGORA.date()

HORARIO_ATUAL = AGORA.time()

HORARIO_PARA_INSERIR_PRINTS = time(hour=17, minute=30, second=0)


load_dotenv(encoding='utf-8', verbose=True)

sites = {
    'Solis': {
        'url': 'https://www.soliscloud.com/#/homepage',
        'login': os.getenv('LOGIN_SOLIS'),
        'senha': os.getenv('SENHA_SOLIS')
    },

    'Solplanet': {
        'url': 'https://internation-pro-cloud.solplanet.net/user/login',
        'login':  os.getenv('LOGIN_SOLPLANET'),
        'senha': os.getenv('SENHA_SOLPLANET')
    },

    'Sungrow': {
        'url': 'https://web3.isolarcloud.com.hk/#/login', 
        'login': os.getenv('LOGIN_SUNGROW'),
        'senha': os.getenv('SENHA_SUNGROW')
    },

    'Shine': {
        'url': 'https://www.renovigi.solar/cus/renovigi/index_po.html?1724337076408',
        'login': os.getenv('LOGIN_SHINE'),
        'senha': os.getenv('SENHA_SHINE')
    },

    'Growatt': {
        'url': 'https://server.growatt.com/?lang=pt',
        'login': os.getenv('LOGIN_GROWATT'),
        'senha': os.getenv('SENHA_GROWATT')
    },

    'PHB': {
        'url': 'http://www.phbsolar.com.br/home/login',
        'login_Imebras': os.getenv('LOGIN_PHB_IMEBRAS'),
        'senha_Imebras': os.getenv('SENHA_PHB_IMEBRAS')
    }
}


remetente = os.getenv('REMETENTE_AVISOS_MONITORAMENTO')

destinatario = os.getenv('DESTINATARIO')

host = os.getenv('HOST')

porta = os.getenv('PORTA')
