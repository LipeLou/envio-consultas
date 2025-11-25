#!/usr/bin/env python3
"""
Script de TESTE - Envia apenas os 3 primeiros emails para o email de teste configurado
Use este script para testar o sistema sem enviar emails para todos os titulares
Suporta arquivos Excel (.xls ou .xlsx)
"""
import logging
import sys
from pathlib import Path
from typing import Dict, Optional

# Importar configurações de teste ANTES de importar outros módulos
from config import config_test as config

# Sobrescrever o módulo config padrão com as configurações de teste
from config import config as config_original
config_original.SMTP_SERVER = config.SMTP_SERVER
config_original.SMTP_PORT = config.SMTP_PORT
config_original.EMAIL_REMETENTE = config.EMAIL_REMETENTE
config_original.EMAIL_SENHA = config.EMAIL_SENHA
config_original.CSV_EMAILS_PATH = config.CSV_EMAILS_PATH
config_original.IGNORE_VALUES = config.IGNORE_VALUES
config_original.XLSX_CONSULTAS_PATH = config.XLSX_CONSULTAS_PATH
config_original.MES_RELATORIO = config.MES_RELATORIO

# Importar funções para ler emails e fazer match (apenas para verificação no log)
from src.email_utils import ler_emails_csv, normalizar_nome
from src.xlsx_parser import processar_xlsx, Titular
from src.html_generator import gerar_html_consultas
from src.email_sender import enviar_email

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format=config.LOG_FORMAT,
    handlers=[
        logging.FileHandler(config.LOGS_DIR / 'test_emails.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)


def encontrar_email_titular(nome_titular: str, emails_dict: Dict[str, str]) -> Optional[str]:
    """
    Encontra o email do titular no dicionário de emails
    Para teste, retorna o email real do titular (apenas para verificação no log)
    O envio sempre será para EMAIL_TESTE
    
    Args:
        nome_titular: Nome do titular (já normalizado)
        emails_dict: Dicionário de emails
    
    Returns:
        Email real do titular (para exibição no log) ou None se não encontrado
    """
    nome_normalizado = normalizar_nome(nome_titular)
    
    # Tentar match exato primeiro
    if nome_normalizado in emails_dict:
        return emails_dict[nome_normalizado]
    
    # Tentar match parcial removendo acentos e caracteres especiais
    nome_sem_especiais = nome_normalizado.replace('Ç', 'C').replace('Ã', 'A').replace('Õ', 'O')
    
    for nome_csv, email in emails_dict.items():
        nome_csv_sem_especiais = nome_csv.replace('Ç', 'C').replace('Ã', 'A').replace('Õ', 'O')
        
        # Match exato sem caracteres especiais
        if nome_sem_especiais == nome_csv_sem_especiais:
            logger.debug(f"Match encontrado (sem especiais): '{nome_titular}' -> '{nome_csv}'")
            return email
        
        # Match parcial (um contém o outro)
        if nome_sem_especiais in nome_csv_sem_especiais or nome_csv_sem_especiais in nome_sem_especiais:
            logger.warning(f"Match parcial encontrado: '{nome_titular}' -> '{nome_csv}'")
            return email
    
    return None


def processar_e_enviar_teste(caminho_xlsx: Path, caminho_csv: Optional[Path] = None, limite: int = 3):
    """
    Processa o arquivo Excel (.xls ou .xlsx) e envia emails de TESTE (apenas os primeiros N titulares)
    
    Args:
        caminho_xlsx: Caminho para o arquivo Excel de consultas (.xls ou .xlsx)
        caminho_csv: Caminho para o arquivo CSV de emails (opcional)
        limite: Número máximo de titulares para processar (padrão: 3)
    """
    logger.info("=" * 60)
    logger.info("MODO DE TESTE - Enviando apenas para email de teste")
    logger.info(f"Email de teste: {config.EMAIL_TESTE}")
    logger.info(f"Limite de titulares: {limite}")
    logger.info("=" * 60)
    
    # Verificar se email de teste está configurado
    if not config.EMAIL_TESTE or config.EMAIL_TESTE == "seu_email_pessoal@gmail.com":
        logger.error("ERRO: Email de teste não configurado!")
        logger.error("Por favor, edite o arquivo config_test.py e defina EMAIL_TESTE com seu email pessoal")
        return
    
    # Verificar se arquivo Excel existe
    if not caminho_xlsx.exists():
        logger.error(f"Arquivo Excel não encontrado: {caminho_xlsx}")
        return
    
    # Ler emails do CSV para verificação (mostrar email real no log)
    # Mas todos os emails serão enviados para EMAIL_TESTE
    logger.info("Lendo arquivo CSV de emails para verificação...")
    try:
        emails_dict = ler_emails_csv(caminho_csv)
        logger.info(f"Carregados {len(emails_dict)} emails do CSV para verificação")
    except Exception as e:
        logger.warning(f"Erro ao ler CSV de emails (continuando sem verificação): {e}")
        emails_dict = {}
    
    # Processar arquivo Excel
    logger.info(f"Processando arquivo Excel: {caminho_xlsx}")
    try:
        titulares = processar_xlsx(caminho_xlsx)
    except Exception as e:
        logger.error(f"Erro ao processar arquivo Excel: {e}")
        return
    
    if not titulares:
        logger.warning("Nenhum titular encontrado no arquivo Excel")
        return
    
    # Limitar número de titulares para teste
    titulares_teste = titulares[:limite]
    total_titulares = len(titulares_teste)
    
    logger.info(f"\n⚠️  MODO TESTE: Processando apenas {total_titulares} de {len(titulares)} titulares")
    logger.info(f"⚠️  Todos os emails serão enviados para: {config.EMAIL_TESTE}")
    logger.info("-" * 60)
    
    # Estatísticas
    emails_enviados = 0
    emails_falhados = 0
    titulares_sem_consultas = 0
    
    # Processar cada titular (limitado)
    for idx, titular in enumerate(titulares_teste, 1):
        # Buscar email real do titular para verificação
        email_real_titular = None
        if emails_dict:
            email_real_titular = encontrar_email_titular(titular.nome, emails_dict)
            if not email_real_titular:
                # Log de debug para ajudar a identificar o problema
                nome_buscado = normalizar_nome(titular.nome)
                logger.debug(f"Email não encontrado para: '{titular.nome}' (normalizado: '{nome_buscado}')")
                # Mostrar alguns nomes do CSV para comparação
                if len(emails_dict) > 0:
                    exemplos = list(emails_dict.keys())[:3]
                    logger.debug(f"Exemplos de nomes no CSV: {exemplos}")
        
        # Montar mensagem com email real (se encontrado)
        if email_real_titular:
            logger.info(f"\n[{idx}/{total_titulares}] TESTE - Processando titular: {titular.nome} ({email_real_titular})")
        else:
            logger.warning(f"\n[{idx}/{total_titulares}] TESTE - Processando titular: {titular.nome} (⚠️ EMAIL NÃO ENCONTRADO NO CSV)")
        
        # Verificar se titular tem consultas
        total_consultas = sum(len(b.consultas) for b in titular.beneficiarios)
        if total_consultas == 0:
            logger.warning(f"Titular {titular.nome} não possui consultas. Pulando...")
            titulares_sem_consultas += 1
            continue
        
        logger.info(f"  - {len(titular.beneficiarios)} beneficiário(s)")
        logger.info(f"  - {total_consultas} consulta(s) total(is)")
        if email_real_titular:
            logger.info(f"  - ✓ Email encontrado no CSV: {email_real_titular}")
        else:
            logger.warning(f"  - ✗ Email NÃO encontrado no CSV para este titular")
        logger.info(f"  - Email de TESTE (destino): {config.EMAIL_TESTE} (enviando para email de teste)")
        
        # Gerar HTML
        try:
            html = gerar_html_consultas(titular)
        except Exception as e:
            logger.error(f"  - Erro ao gerar HTML: {e}")
            emails_falhados += 1
            continue
        
        # Enviar email para o email de teste
        assunto = f"[TESTE] Relatório de Consultas - {titular.nome}"
        logger.info(f"  - Enviando email de TESTE...")
        
        if enviar_email(config.EMAIL_TESTE, assunto, html):
            emails_enviados += 1
            logger.info(f"  - ✓ Email de TESTE enviado com sucesso!")
        else:
            emails_falhados += 1
            logger.error(f"  - ✗ Falha ao enviar email de TESTE")
    
    # Resumo final
    logger.info("\n" + "=" * 60)
    logger.info("RESUMO DO TESTE")
    logger.info("=" * 60)
    logger.info(f"Total de titulares processados: {total_titulares}")
    logger.info(f"Emails de teste enviados com sucesso: {emails_enviados}")
    logger.info(f"Emails de teste falhados: {emails_falhados}")
    logger.info(f"Titulares sem consultas: {titulares_sem_consultas}")
    logger.info(f"Todos os emails foram enviados para: {config.EMAIL_TESTE}")
    logger.info("=" * 60)


def main():
    """Função principal de teste"""
    # Usar sempre os caminhos padrão configurados
    caminho_xlsx = config.XLSX_CONSULTAS_PATH
    logger.info(f"Usando arquivo Excel: {caminho_xlsx}")
    
    # Verificar se foi passado limite como argumento (opcional)
    limite = 3  # Padrão
    if len(sys.argv) > 1:
        try:
            limite = int(sys.argv[1])
            logger.info(f"Usando limite customizado: {limite}")
        except ValueError:
            logger.warning(f"Limite inválido '{sys.argv[1]}', usando padrão: 3")
    
    # Processar e enviar (modo teste)
    try:
        processar_e_enviar_teste(caminho_xlsx, None, limite)
    except KeyboardInterrupt:
        logger.info("\nProcessamento interrompido pelo usuário")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Erro fatal: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()

