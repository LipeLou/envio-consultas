"""
Configurações de teste do sistema de envio de emails
Use este arquivo para testar enviando emails para você mesmo
"""
import os
from pathlib import Path
from dotenv import load_dotenv

# Carregar variáveis de ambiente do arquivo .env (se existir)
load_dotenv()

# Configurações de Email (TESTE - substitua pelo seu email pessoal)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_REMETENTE = "seu_email_teste@gmail.com"  # ALTERE AQUI: Configure seu email para testes

# Senha do email (será lida do arquivo .env)
EMAIL_SENHA = os.getenv("EMAIL_SENHA", "")

# Email de teste - SUBSTITUA PELO SEU EMAIL PESSOAL
EMAIL_TESTE = "seu_email_pessoal@gmail.com"  # ALTERE AQUI: Configure seu email pessoal para receber testes

# Caminhos dos arquivos
BASE_DIR = Path(__file__).parent.parent  # Voltar para a raiz do projeto
DATA_DIR = BASE_DIR / "data"
LOGS_DIR = BASE_DIR / "logs"
CSV_EMAILS_PATH = DATA_DIR / "email.csv"
XLSX_CONSULTAS_PATH = DATA_DIR / "consultas.xlsx"

# Configurações de processamento
IGNORE_VALUES = ["inrrelevante", "Evento", "Prestador", "", "nan", "NaN", "None", "none"]

# Configurações de logging
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"

