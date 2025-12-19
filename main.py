#!/usr/bin/env python3
"""
Sistema de Envio de Emails de Consultas
Processa planilha Excel (.xls ou .xlsx) com consultas e envia emails para cada titular
"""
import logging
import sys
import csv
import datetime
from pathlib import Path
from typing import Dict, Optional, List, Any
from config.config import LOG_FORMAT, XLSX_CONSULTAS_PATH, LOGS_DIR
from src.email_utils import ler_emails_csv, normalizar_nome
from src.xlsx_parser import processar_xlsx, Titular
from src.html_generator import gerar_html_consultas
from src.email_sender import enviar_email

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format=LOG_FORMAT,
    handlers=[
        logging.FileHandler(LOGS_DIR / 'envio_emails.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)


def encontrar_email_titular(nome_titular: str, emails_dict: Dict[str, str]) -> Optional[str]:
    """
    Encontra o email do titular no dicionário de emails
    
    Args:
        nome_titular: Nome do titular (já normalizado)
        emails_dict: Dicionário de emails
    
    Returns:
        Email do titular ou None se não encontrado
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


def processar_e_enviar(caminho_xlsx: Path, caminho_csv: Optional[Path] = None):
    """
    Processa o arquivo Excel (.xls ou .xlsx) e envia emails para os titulares
    
    Args:
        caminho_xlsx: Caminho para o arquivo Excel de consultas (.xls ou .xlsx)
        caminho_csv: Caminho para o arquivo CSV de emails (opcional)
    """
    logger.info("=" * 60)
    logger.info("Iniciando processamento de consultas e envio de emails")
    logger.info("=" * 60)
    
    # Verificar se arquivo Excel existe
    if not caminho_xlsx.exists():
        logger.error(f"Arquivo Excel não encontrado: {caminho_xlsx}")
        return
    
    # Ler emails do CSV
    logger.info("Lendo arquivo CSV de emails...")
    try:
        emails_dict = ler_emails_csv(caminho_csv)
    except Exception as e:
        logger.error(f"Erro ao ler CSV de emails: {e}")
        return
    
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
    
    # Estatísticas
    total_titulares = len(titulares)
    emails_enviados = 0
    emails_falhados = 0
    titulares_sem_email = 0
    
    # Lista de erros para relatório
    erros_relatorio = []
    
    # Processar cada titular
    logger.info(f"\nProcessando {total_titulares} titulares...")
    logger.info("-" * 60)
    
    for idx, titular in enumerate(titulares, 1):
        logger.info(f"\n[{idx}/{total_titulares}] Processando titular: {titular.nome}")
        
        # Verificar se titular tem consultas
        total_consultas = sum(len(b.consultas) for b in titular.beneficiarios)
        if total_consultas == 0:
            logger.warning(f"Titular {titular.nome} não possui consultas. Pulando...")
            continue
        
        logger.info(f"  - {len(titular.beneficiarios)} beneficiário(s)")
        logger.info(f"  - {total_consultas} consulta(s) total(is)")
        
        # Encontrar email do titular
        email_titular = encontrar_email_titular(titular.nome, emails_dict)
        
        if not email_titular:
            nome_buscado = normalizar_nome(titular.nome)
            logger.warning(f"  - ⚠️ EMAIL NÃO ENCONTRADO para: {titular.nome}")
            logger.warning(f"  - Nome normalizado buscado: '{nome_buscado}'")
            titulares_sem_email += 1
            erros_relatorio.append({
                "titular": titular.nome,
                "motivo": "Email não encontrado",
                "detalhe": f"Nome normalizado: {nome_buscado}",
                "email": ""
            })
            continue
        
        logger.info(f"  - Email encontrado: {email_titular}")
        
        # Gerar HTML
        try:
            html = gerar_html_consultas(titular)
        except Exception as e:
            logger.error(f"  - Erro ao gerar HTML: {e}")
            emails_falhados += 1
            erros_relatorio.append({
                "titular": titular.nome,
                "motivo": "Erro na geração do HTML",
                "detalhe": str(e),
                "email": email_titular
            })
            continue
        
        # Enviar email
        assunto = f"Relatório de Consultas - {titular.nome}"
        logger.info(f"  - Enviando email...")
        
        if enviar_email(email_titular, assunto, html):
            emails_enviados += 1
            logger.info(f"  - ✓ Email enviado com sucesso!")
        else:
            emails_falhados += 1
            logger.error(f"  - ✗ Falha ao enviar email")
            erros_relatorio.append({
                "titular": titular.nome,
                "motivo": "Falha no envio SMTP",
                "detalhe": "Verificar logs para erro específico do SMTP",
                "email": email_titular
            })
    
    # Gerar relatório de erros se houver falhas
    if erros_relatorio:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        arquivo_relatorio = LOGS_DIR / f"nao_enviados_{timestamp}.csv"
        
        try:
            with open(arquivo_relatorio, mode='w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=["titular", "motivo", "detalhe", "email"])
                writer.writeheader()
                writer.writerows(erros_relatorio)
            
            logger.info(f"\n⚠️  Relatório de emails não enviados salvo em:")
            logger.info(f"   -> {arquivo_relatorio}")
        except Exception as e:
            logger.error(f"Erro ao salvar relatório de erros: {e}")

    # Resumo final
    logger.info("\n" + "=" * 60)
    logger.info("RESUMO DO PROCESSAMENTO")
    logger.info("=" * 60)
    logger.info(f"Total de titulares processados: {total_titulares}")
    logger.info(f"Emails enviados com sucesso: {emails_enviados}")
    logger.info(f"Emails falhados: {emails_falhados}")
    logger.info(f"Titulares sem email: {titulares_sem_email}")
    logger.info("=" * 60)


def main():
    """Função principal"""
    # Usar sempre os caminhos padrão configurados
    caminho_xlsx = XLSX_CONSULTAS_PATH
    logger.info(f"Usando arquivo Excel: {caminho_xlsx}")
    
    # Processar e enviar
    try:
        processar_e_enviar(caminho_xlsx, None)
    except KeyboardInterrupt:
        logger.info("\nProcessamento interrompido pelo usuário")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Erro fatal: {e}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()

