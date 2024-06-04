import psycopg2
import mariadb
from ExceptionClass import *
import pandas as pd
from consts import CONFIG

#Por Gabriel Nogueira em virtude da Novetech Soluções Technológicas
#v2.0.1

class easBD:
    def __init__(self) -> None:
        self._bdAddress = CONFIG["CREDENCIAISEAS"]["address"]
        self._port = CONFIG["CREDENCIAISEAS"]["port"]
        self._login = CONFIG["CREDENCIAISEAS"]["login"]
        self._senha = CONFIG["CREDENCIAISEAS"]["senha"]

        try:
            self._conn = mariadb.connect(
                user = self._login,
                password= self._senha,
                host = self._bdAddress,
                port = int(self._port)
            )
            
        except Exception as e:
            raise BadBDConnection('Falha ao conectar ao host do banco de dados ', e)
    
    def cursor_conn(self):
        return self._conn.cursor()

    def change_db(self, dbname):
        self._conn.select_db(dbname)

    def close_db(self):
        self._conn.close()
    
    def execute_SQL(self, SQL, data = None):
        with self.cursor_conn() as cursor:
            try:
                cursor.execute(SQL, data)
                return cursor.fetchall()
            except Exception as e:
                    return e

class esusBD:

    def __init__(self, unidade:str, namedb:str, host:str, port:str,login:str, senha:str):
        
        self._unidade = unidade
        self._host = host
        self._port = port
        self._login = login
        self._namedb = namedb 
        self._senha = senha
        
        self._conexao = None

        try:
            self._conexao = self.connection()
        except BadBDConnection as e:
            self.close_conn()
            raise BadBDConnection(e)
            

    def cursor_conn(self):
        return self._conexao.cursor()

    def connection(self):
        try:       
            return psycopg2.connect(
                host=self._host,
                port=self._port,
                dbname=self._namedb,
                user=self._login,
                password=self._senha)
        except psycopg2.OperationalError as e:
            self.close_conn()
            raise BadBDConnection(f"Erro de conexão: Login, senha ou Nome do Banco não reconhecido.") 
    
    def start_query(self, sql):
    
        with self.cursor_conn() as cursor:
            sql_query = sql
            try:
                print("Extraindo...")
                cursor.execute(sql_query)
                return cursor.fetchall()
            except Exception as e:
                return e
            
    def close_conn(self):
        if self._conexao:
            self._conexao.close()
            print("Conexão encerrada.")

class Includer:

    def __init__(self, df_colunas):
        self.colunas = df_colunas
        self._df = pd.DataFrame(columns=self.colunas)
        self._linhas_preenchidas = 0

    def insert(self, dados):

        for valores in dados:
            self._linhas_preenchidas += 1
            for coluna, valor in enumerate(valores):
                self._df.loc[self._linhas_preenchidas , self._df.columns[coluna]] = valor
        return 'Extração concluida.'
    
    def getDf(self):
        return self._df

# --> Para extrações unitárias.
if __name__ == '__main__':
    pass
