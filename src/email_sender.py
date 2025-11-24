"""
Módulo para envio de emails
"""
import smtplib
import logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import Optional
from config.config import SMTP_SERVER, SMTP_PORT, EMAIL_REMETENTE, EMAIL_SENHA

logger = logging.getLogger(__name__)


def validar_email(email: str) -> bool:
    """
    Valida formato básico de email
    
    Args:
        email: Endereço de email a validar
    
    Returns:
        True se o email parece válido
    """
    if not email or not isinstance(email, str):
        return False
    
    email = email.strip()
    if "@" not in email or "." not in email.split("@")[1]:
        return False
    
    return True


def enviar_email(destinatario: str, assunto: str, corpo_html: str, 
                 remetente: Optional[str] = None, senha: Optional[str] = None) -> bool:
    """
    Envia email HTML para o destinatário
    
    Args:
        destinatario: Email do destinatário
        assunto: Assunto do email
        corpo_html: Corpo do email em HTML
        remetente: Email remetente (usa padrão do config se None)
        senha: Senha do email (usa padrão do config se None)
    
    Returns:
        True se enviado com sucesso, False caso contrário
    """
    if not validar_email(destinatario):
        logger.warning(f"Email inválido: {destinatario}")
        return False
    
    if remetente is None:
        remetente = EMAIL_REMETENTE
    
    if senha is None:
        senha = EMAIL_SENHA
    
    if not senha:
        logger.error("Senha do email não configurada. Configure a variável EMAIL_SENHA.")
        return False
    
    try:
        # Criar mensagem
        msg = MIMEMultipart('alternative')
        msg['Subject'] = assunto
        msg['From'] = remetente
        msg['To'] = destinatario
        
        # Adicionar corpo HTML
        parte_html = MIMEText(corpo_html, 'html', 'utf-8')
        msg.attach(parte_html)
        
        # Conectar e enviar
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(remetente, senha)
            server.send_message(msg)
        
        logger.info(f"Email enviado com sucesso para: {destinatario}")
        return True
    
    except smtplib.SMTPAuthenticationError as e:
        logger.error(f"Erro de autenticação ao enviar email para {destinatario}: {e}")
        return False
    except smtplib.SMTPException as e:
        logger.error(f"Erro SMTP ao enviar email para {destinatario}: {e}")
        return False
    except Exception as e:
        logger.error(f"Erro inesperado ao enviar email para {destinatario}: {e}")
        return False

