from os import _exit
from consts import CONFIG_PATH, LOG_GERAL

class BadConection(Exception):
    def __init__(self, message):
        self.message = message

class BOTMonitoramentoException(Exception):
    def __init__(self, message):
        self.message = message

class BadBDConnection(Exception):
    def __init__(self, message):
        self.message = message

class ProcessClosed(Exception):
    def __init__(self, message):
        self.message = message

class outputLogException():
    def iniError():
        LOG_GERAL.warning(f"{CONFIG_PATH} fora de conformidade. Abra novamente a aplicação.")
        _exit(0)

class IntegridadeComprometida(Exception):
    def __init__(self, message):
        self.message = message