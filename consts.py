import os
import logging
from datetime import datetime
from google_play_scraper import app
import configparser
import json

class Functions:

    def load_json(json_file):
        if os.path.exists(os.path.join(JSON_PATH, json_file)):
            with open(os.path.join(JSON_PATH, json_file)) as ip_js_file:
                return json.load(ip_js_file)

    def configure_logs(log, log_path):
        log.setLevel(logging.INFO)
        handler = logging.FileHandler(log_path, encoding="utf-8")
        handler.setLevel(logging.INFO)
        handler.setFormatter(logging.Formatter('%(asctime)s | %(levelname)s | %(message)s'))
        log.addHandler(handler)

    def _version_app():
        _package_name = PACKAGE_APP
        try:
            _app_info = app(_package_name)
            return _app_info
        except Exception as e:
            raise Exception
        
    def check_exists(path):
        if not os.path.exists(path):
            os.makedirs(path)
            _pastaFalta = path.split("\\")[-1]
            LOG_GERAL.warning(
            f"Pasta {_pastaFalta}, não encontrada. Criada em momento de execução... | *Risco de falha na integridade de dados do BOT.*")

#PASTAS PADRÃO DA APLICAÇÃO
MAIN_PATH = os.getcwd()

CONFIG_PATH = os.path.join(MAIN_PATH, "config")
Functions.check_exists(CONFIG_PATH)

SQLS_PATH = os.path.join(CONFIG_PATH, 'sqls')
Functions.check_exists(SQLS_PATH)

JSON_PATH = os.path.join(CONFIG_PATH, 'json')
Functions.check_exists(JSON_PATH)

LOGS_PATH = os.path.join(MAIN_PATH,"logs")
Functions.check_exists(LOGS_PATH)

EXTRACT_PATH = os.path.join(MAIN_PATH, "extraidos")
Functions.check_exists(EXTRACT_PATH)

OUTPUT_PATH = os.path.join(MAIN_PATH, "saida")
Functions.check_exists(OUTPUT_PATH)

PACKAGE_APP = "br.com.novetech.atendsaude.atendsaudev4"

VERSAO_MONITOR_ATUAL = '2.0.1'

COLUNAS_ATEND = [   "UNIDADE", 
                    "LOTE",
                    "DATA",
                    "PERIODO", 
                    "VALIDAS", 
                    "INVALIDAS"]

COLUNAS_PEC = [ "UNIDADE EXTRACAO",
                "UNIDADE",
                "LOTE",
                "DATA", 
                "HORA",
                "PERIODO",
                "ENVIADO POR",
                "VALIDAS", 
                "INVALIDAS", 
                "TIPO"]

CONFIG_INI = os.path.join(MAIN_PATH, CONFIG_PATH, 'config.ini')

CONFIG = configparser.ConfigParser()
CONFIG.read(CONFIG_INI)
ESUS_IPS = Functions.load_json('esus.json')
EAS_IPS = Functions.load_json('eas.json')
EMAIL_OPTS = Functions.load_json('emails.json')

LOG_GERAL = logging.getLogger("main_log")
LOG_GERAL_PATH = os.path.join(LOGS_PATH, f'process{datetime.today().strftime("%d-%m-%Y")}.log')
Functions.configure_logs(LOG_GERAL, LOG_GERAL_PATH)

try:
    VERSAO_ATUAL = Functions._version_app()['version']
except:
    VERSAO_ATUAL = CONFIG["UTILS"]["version_app"]


TEXTO_PADRAO_MAIL = f"\
Prezados,\n\n\
\
Por meio deste e-mail, realizamos a entrega do relatório semanal de integração dos lotes.\n\
Devemos destacar desde já que a visualização do relatório segue o modelo de comparação 'de-para', ou seja,\n\
a análise dos dados referentes aos lotes é realizada do AtendSaúde para o eSUS, sempre considerando que o dado\n\
mais consistente e fidedigno é proveniente do próprio AtendSaúde, e a expectativa é que este lote seja devidamente integrado ao eSUS.\
\
Em caso de identificação de falhas ou divergências nos dados apresentados,\n\
solicitamos cordialmente que notifiquem ao Apoio Institucional para as devidas correções.\n\n\
\
Agradecemos pela colaboração e permanecemos à disposição para quaisquer esclarecimentos adicionais.\n\n\
\
Dados Técnicos:\n\
- Nome: Monitoramento-Integração\n\
- Versão: {VERSAO_MONITOR_ATUAL}\n\n\
\
Este e-mail tem caráter informativo e foi gerado automaticamente. Por favor, não responda a este e-mail.\n\n\
\
Atenciosamente,\n\n\
\
Gabriel\n\
Apoio Institucional Novetech."