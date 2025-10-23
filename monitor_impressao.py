import os
import email
import imaplib
import time
import schedule
from dotenv import load_dotenv
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

# Esta é a linha mágica:
# 1. Path(__file__) -> Pega o caminho do script atual (ex: C:\Users\...\script.py)
# 2. .resolve()   -> Converte para um caminho absoluto e limpo.
# 3. .parent      -> Pega o diretório pai (a pasta onde o script está).
BASE_DIR = Path(__file__).resolve().parent

# Agora, BASE_DIR é um objeto que representa o caminho absoluto
# para a pasta onde seu script está.
print(f"O script está rodando em: {BASE_DIR}")

# --- Como usar BASE_DIR para criar outros caminhos ---

# 1. Para criar o caminho para o seu arquivo de log
#    O operador "/" é usado para juntar caminhos de forma inteligente.
LOG_FILE_PATH = BASE_DIR / "log.txt"
print(f"O caminho do log será: {LOG_FILE_PATH}")

# 2. Para criar o caminho para a pasta de downloads
DOWNLOADS_FOLDER_PATH = BASE_DIR / "pdfs_para_imprimir"
print(f"A pasta de downloads será: {DOWNLOADS_FOLDER_PATH}")

# Você pode até criar a pasta se ela não existir:
DOWNLOADS_FOLDER_PATH.mkdir(exist_ok=True)

ENV_PATH = BASE_DIR / ".env"
CYCLE_TIME = 30

load_dotenv(dotenv_path=ENV_PATH)

IMAP_SERVER = os.getenv("IMAP_SERVER")
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# --- Configuração do Logging ---
# Cria um logger. Isso irá capturar todas as saídas.
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

# Usa um handler que rotaciona o log para não ficar gigante (ex: 5MB por arquivo, mantém 2 backups)
my_handler = RotatingFileHandler(LOG_FILE_PATH, mode='a', maxBytes=5*1024*1024, 
                                 backupCount=2, encoding=None, delay=0)
my_handler.setFormatter(log_formatter)
my_handler.setLevel(logging.INFO)

app_log = logging.getLogger('root')
app_log.setLevel(logging.INFO)
app_log.addHandler(my_handler)

# --- Funções ---

def imprimir_pdf(caminho_do_arquivo):
    """Envia um arquivo para a impressora padrão do Windows."""
    try:
        app_log.info(f"  -> Enviando '{os.path.basename(caminho_do_arquivo)}' para a impressora...")
        os.startfile(caminho_do_arquivo, "print")
        app_log.info("  -> Arquivo enviado para a fila de impressão.")
        return True
    except Exception as e:
        app_log.error(f"  -> Erro ao tentar imprimir o arquivo: {e}")
        return False

def verificar_e_imprimir_emails():
    """Conecta ao e-mail, baixa anexos PDF e os imprime."""

    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        mail.select("inbox")

        status, data = mail.search(None, '(UNSEEN)')

        if status == 'OK':
            email_ids = data[0].split()
            if not email_ids:
                return

            app_log.info(f"Encontrados {len(email_ids)} novos e-mails.")

            for email_id in email_ids:
                status, email_data = mail.fetch(email_id, '(RFC822)')
                if status == 'OK':
                    msg = email.message_from_bytes(email_data[0][1])
                    sender = msg['from']
                    subject = msg['subject']
                    app_log.info(f"Processando e-mail de: {sender} | Assunto: {subject}")

                    if msg.get_content_maintype() == 'multipart':
                        for part in msg.walk():
                            if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
                                continue

                            filename = part.get_filename()
                            if filename and filename.lower().endswith('.pdf'):
                                app_log.info(f"  -> Anexo PDF encontrado: {filename}")

                                filepath = os.path.join(DOWNLOADS_FOLDER_PATH, filename)
                                with open(filepath, 'wb') as f:
                                    f.write(part.get_payload(decode=True))
                                
                                if imprimir_pdf(filepath):
                                    mail.store(email_id, '+FLAGS', '\\Seen')
                                    app_log.info("  -> E-mail marcado como lido.")
        
        mail.logout()

    except imaplib.IMAP4.error as e:
        app_log.error(f"Erro de IMAP: {e}")
    except Exception as e:
        app_log.error(f"Ocorreu um erro inesperado: {e}")

# --- Execução Principal ---
if __name__ == "__main__":
    app_log.info("--- Serviço de Impressão Automática por E-mail INICIADO ---")
    
    verificar_e_imprimir_emails()
    schedule.every(CYCLE_TIME).seconds.do(verificar_e_imprimir_emails)

    while True:
        schedule.run_pending()
        time.sleep(1)