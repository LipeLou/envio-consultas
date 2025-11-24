"""
Utilitários para leitura de emails e envio
"""
import pandas as pd
import re
import logging
from pathlib import Path
from typing import Dict, Optional
from config.config import CSV_EMAILS_PATH

logger = logging.getLogger(__name__)


def ler_emails_csv(caminho_csv: Optional[Path] = None) -> Dict[str, str]:
    """
    Lê o arquivo CSV de emails e retorna um dicionário {nome: email}
    Prioriza sempre o email principal, usando o outro email apenas se o principal não existir
    
    Args:
        caminho_csv: Caminho para o arquivo CSV. Se None, usa o padrão do config.
    
    Returns:
        Dicionário com nomes normalizados (uppercase) como chaves e emails como valores
    """
    if caminho_csv is None:
        caminho_csv = CSV_EMAILS_PATH
    
    try:
        # Tentar ler CSV com diferentes encodings (comum no Brasil: ISO-8859-1, Windows-1252)
        encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252', 'windows-1252']
        df = None
        encoding_usado = None
        
        for encoding in encodings:
            try:
                df = pd.read_csv(caminho_csv, sep=';', encoding=encoding)
                encoding_usado = encoding
                logger.debug(f"CSV lido com sucesso usando encoding: {encoding}")
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        
        if df is None:
            raise ValueError(f"Não foi possível ler o arquivo CSV com nenhum dos encodings testados: {encodings}")
        
        # Criar dicionário nome -> email
        emails_dict = {}
        
        for _, row in df.iterrows():
            nome = str(row.get('NOME', '')).strip()
            email_principal = str(row.get('E-MAIL PRINCIPAL', '')).strip()
            outro_email = str(row.get('OUTRO E-MAIL', '')).strip()
            
            # Ignorar linhas vazias ou sem nome
            if not nome or nome == 'nan' or nome.lower() == 'nan' or 'TOTAL' in nome.upper():
                continue
            
            # Limpar valores NaN do pandas
            if email_principal.lower() == 'nan' or email_principal == '':
                email_principal = ''
            if outro_email.lower() == 'nan' or outro_email == '':
                outro_email = ''
            
            # Priorizar email principal, se não tiver, usar outro email
            if email_principal:
                email = email_principal
            elif outro_email:
                email = outro_email
            else:
                # Sem email, pular esta linha
                continue
            
            if email:
                # Normalizar nome para uppercase e remover espaços extras
                nome_normalizado = normalizar_nome(nome)
                # Se já existe, manter o primeiro (email principal tem prioridade)
                if nome_normalizado not in emails_dict:
                    emails_dict[nome_normalizado] = email
                else:
                    # Se já existe, verificar se o atual é email principal
                    # (não sobrescrever email principal com outro email)
                    logger.debug(f"Nome duplicado no CSV: {nome_normalizado} (mantendo primeiro email)")
        
        logger.info(f"Carregados {len(emails_dict)} emails do arquivo CSV")
        return emails_dict
    
    except Exception as e:
        logger.error(f"Erro ao ler arquivo CSV de emails: {e}")
        raise


def extrair_nome_titular(texto: str) -> Optional[str]:
    """
    Extrai o nome do titular do padrão "Cód Titular: XXXXXX - NOME"
    Retorna o nome normalizado (mesma normalização usada no CSV)
    
    Args:
        texto: Texto da célula
    
    Returns:
        Nome do titular normalizado ou None se não encontrar o padrão
    """
    if not texto or not isinstance(texto, str):
        return None
    
    # Padrão: "Cód Titular: XXXXXX - NOME"
    match = re.search(r'Cód Titular:\s*\d+\s*-\s*(.+)', texto, re.IGNORECASE)
    if match:
        nome = match.group(1).strip()
        # Usar a mesma função de normalização usada no CSV
        return normalizar_nome(nome)
    
    return None


def extrair_nome_beneficiario(texto: str) -> Optional[str]:
    """
    Extrai o nome do beneficiário do padrão "Beneficiário: XXXXXX - NOME (XX)"
    Retorna o nome normalizado
    
    Args:
        texto: Texto da célula
    
    Returns:
        Nome do beneficiário normalizado ou None se não encontrar o padrão
    """
    if not texto or not isinstance(texto, str):
        return None
    
    # Padrão: "Beneficiário: XXXXXX - NOME (XX)"
    match = re.search(r'Beneficiário:\s*\d+\s*-\s*([^(]+)', texto, re.IGNORECASE)
    if match:
        nome = match.group(1).strip()
        # Usar a mesma função de normalização
        return normalizar_nome(nome)
    
    return None


def normalizar_nome(nome: str) -> str:
    """
    Normaliza nome para comparação (uppercase, remove espaços extras)
    
    Args:
        nome: Nome a normalizar
    
    Returns:
        Nome normalizado
    """
    if not nome:
        return ""
    return " ".join(nome.upper().split())

