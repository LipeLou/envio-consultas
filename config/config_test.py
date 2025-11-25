"""
Configurações de teste do sistema de envio de emails
Use este arquivo para testar enviando emails para você mesmo
"""
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# CONFIGURE AQUI!
# ------------------------------------------------------------

# Configurações de Email (TESTE - substitua pelo seu email teste)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_REMETENTE = "seu_email_remetente@gmail.com"  # ALTERE AQUI: Configure seu email para testes

# Mês/ano a ser exibido no relatório
MES_RELATORIO = "Mẽs teste"  # ALTERE AQUI: Configure o mês/ano do relatório

# Email de teste - SUBSTITUA PELO SEU EMAIL PESSOAL
EMAIL_TESTE = "seu_email_pessoal@gmail.com"  # ALTERE AQUI: Configure seu email pessoal para receber testes

# ------------------------------------------------------------







# Senha do email (será lida do arquivo .env)
EMAIL_SENHA = os.getenv("EMAIL_SENHA", "")

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

