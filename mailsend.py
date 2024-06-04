import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from consts import OUTPUT_PATH, TEXTO_PADRAO_MAIL, LOG_GERAL, EMAIL_OPTS, CONFIG
from datetime import datetime
from os import path
import traceback


#Por Gabriel Nogueira em virtude da Novetech Soluções Technológicas
#v2.0.1

SMTP_SERVERS = {
    "gmail": {"s_address": "smtp.gmail.com", "port": 587}
}

REMETENTE = {"automacao":
             {"name": "MONITORAMENTO ATENDSAUDE", "email": CONFIG["CREDENCIALSMTP"]["email"], "senha": CONFIG["CREDENCIALSMTP"]["auth"]}}

DESTINATARIO = EMAIL_OPTS['DESTINATARIOS']

class MailSend:

    def __init__(self, assunto, msg) -> None:
        LOG_GERAL.info("Tentativa de envio por e-mail...")

        self.msg = msg
        self.assunto = assunto
        self.remetente = REMETENTE["automacao"]
        self.destinatario = DESTINATARIO

        mensagem = MIMEMultipart()
        mensagem["From"] = f'{self.remetente["name"]} <{self.remetente["email"]}>'
        mensagem["To"] = ', '.join(self.destinatario.values())
        mensagem["Subject"] = self.assunto
        mensagem.attach(MIMEText(self.msg, "plain", "utf-8"))

        localArquivo = f"{OUTPUT_PATH}\\MONITORAMENTOLOTES-{datetime.today().strftime('%d-%m-%Y')}.xlsx"

        if path.exists(localArquivo):
            with open(localArquivo, 'rb') as relatorio:

                anexoTexto = MIMEApplication(relatorio.read())
                anexoTexto[
                    "Content-Disposition"] = f"attachment; filename=MONITORAMENTOLOTES{datetime.today().strftime('%d-%m-%Y')}.xlsx"
                mensagem.attach(anexoTexto)
        else:
            LOG_GERAL.error("Nenhum relatório atual encontrado para envio.")
            quit()

        self.send_mail(mensagem.as_string())

    def send_mail(self, mail):
        try:
            with smtplib.SMTP(SMTP_SERVERS["gmail"]["s_address"], SMTP_SERVERS["gmail"]["port"]) as servidor:
                servidor.ehlo()
                servidor.starttls()
                servidor.login(self.remetente["email"], self.remetente["senha"])
                servidor.sendmail(self.remetente["email"], list(self.destinatario.values()), mail)
        except Exception as e:
            #print(traceback.format_exc())
            LOG_GERAL.critical("Erro de acesso ao servidor SMTP.")
            LOG_GERAL.critical(f"{SMTP_SERVERS['gmail']}")

if __name__ == "__main__":
    pass
    MailSend("Relatório Monitoramento LOTE | Integração", TEXTO_PADRAO_MAIL)
