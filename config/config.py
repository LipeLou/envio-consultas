"""
Configurações do sistema de envio de emails
"""
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# CONFIGURE AQUI!
# ------------------------------------------------------------

# Configurações de Email
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_REMETENTE = "seu_email@gmail.com"  # ALTERE AQUI: Configure seu email remetente

# Mês/ano a ser exibido no relatório
MES_RELATORIO = ""  # ALTERE AQUI: Configure o mês/ano do relatório

# Senha do email (será lida do arquivo .env)
EMAIL_SENHA = os.getenv("EMAIL_SENHA", "")

# ------------------------------------------------------------







# Caminhos dos arquivos
BASE_DIR = Path(__file__).parent.parent 
DATA_DIR = BASE_DIR / "data"
LOGS_DIR = BASE_DIR / "logs"
CSV_EMAILS_PATH = DATA_DIR / "email.csv"
XLSX_CONSULTAS_PATH = DATA_DIR / "consultas.xls"

# Configurações de processamento
IGNORE_VALUES = ["inrrelevante", "Evento", "Prestador", "", "nan", "NaN", "None", "none"]

# Configurações de logging
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"

