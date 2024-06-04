from Agents import esusBD, Includer, easBD
import os
from consts import *
import traceback
from Sistematizar import Sistematizar
from datetime import datetime
from ExceptionClass import BadBDConnection
from util import check_log
import mailsend

#Por Gabriel Nogueira em virtude da Novetech Soluções Technológicas
#v2.0.1

def load_sql(sql):
    try:
        sql = os.path.join(SQLS_PATH, sql)
        with open(sql, 'r', encoding='utf-8-sig') as sql_file:
            return sql_file.read()
    except Exception as e:
        pass

class Extraction:
 
    def bd_esus(mes=datetime.today().strftime('%m'), ano=datetime.today().strftime('%Y')):

        _port = CONFIG["CREDENCIAISBD"]["port"]
        _nameDb = CONFIG["CREDENCIAISBD"]["namedb"]
        _login = CONFIG["CREDENCIAISBD"]["login"]
        _senha = CONFIG["CREDENCIAISBD"]["senha"]
        

        for estado, municipios in ESUS_IPS.items():
            for municipio, unidades in municipios.items():  

                municipio_path = os.path.join(EXTRACT_PATH, estado.upper(), municipio.upper())
                if not os.path.exists(municipio_path):
                        os.makedirs(municipio_path)

                df_municipio = Includer(COLUNAS_PEC)
                for unidade, host in unidades.items():
                    print(f"\nCONECTANDO: {(municipio, unidade, host)}")
                    try:
                        senha = _senha
                        port = _port
                        user = _login
                            
                        agentBD = esusBD(unidade, _nameDb, host, port, user, senha)
                        extracao = agentBD.start_query(load_sql("lotes.sql"))
                    except BadBDConnection as e:
                        print(e)
                        extracao = [(f'{unidade}',)+('SEM REDE',) * (len(COLUNAS_PEC)-2)]
                    
                    for ind, valor in enumerate(extracao):
                        extracao[ind] = (f'{unidade}',) + extracao[ind]

                    df_municipio.insert(extracao)
                df_municipio.getDf().to_excel(os.path.join(municipio_path, f'LOTES-{municipio}-PEC-{datetime.today().strftime("%d-%m-%Y")}.xlsx'), index=False)


    def bd_eas(mes=datetime.today().strftime('%m'), ano=datetime.today().strftime('%Y')):
        
        try:
            BDAtend = easBD()
        except BadBDConnection as e:
            raise BadBDConnection(e)
        
        SQL = load_sql('loteseas.sql')

        for estado, municipios in EAS_IPS.items():
            for municipio, bdname in municipios.items():
                print(f"\nCONECTANDO: {(municipio, bdname)}")

                municipio_path = os.path.join(EXTRACT_PATH, estado.upper(), municipio.upper())
                if not os.path.exists(municipio_path):
                        os.makedirs(municipio_path)
                    
                try:
                    BDAtend.change_db(bdname)
                    result = BDAtend.execute_SQL(SQL, {'comp': f"{f'{int(mes)}'}/{ano}"})
                    df = Includer(COLUNAS_ATEND)
                    df.insert(result)
                    df.getDf().to_excel(os.path.join(municipio_path, f'LOTES-{municipio}-ATEND-{datetime.today().strftime("%d-%m-%Y")}.xlsx'), index=False)
                except Exception as e:
                    print(e)


if __name__ == "__main__":

    try:
        check_log()
        Extraction.bd_eas()
        Extraction.bd_esus()
        LOG_GERAL.info("SISTEMATIZANDO...")
        Sistematizar()
        LOG_GERAL.info("SISTEMATIZAÇÃO COMPLETA.")
        mailsend.MailSend("Relatório Monitoramento Integração", TEXTO_PADRAO_MAIL)
        os._exit(0)
    except KeyboardInterrupt:
        LOG_GERAL.critical("Um usuário pressionou Ctrl+C. Encerrando o script.")
        os._exit(0)
    except SystemExit:
        LOG_GERAL.critical("O sistema solicitou a saída do script.")
        os._exit(0)
    except Exception as e:
        traceback_str = traceback.format_exc()
        LOG_GERAL.critical(f"Ocorreu uma exceção inesperada: {str(e)}\n{traceback_str}")
        print(traceback_str)
        os._exit(0)
        

    