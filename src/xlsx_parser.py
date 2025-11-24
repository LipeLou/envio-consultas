"""
Parser para arquivo Excel (.xls ou .xlsx) de consultas
"""
import pandas as pd
import logging
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from config.config import IGNORE_VALUES
from src.email_utils import extrair_nome_titular, extrair_nome_beneficiario, normalizar_nome

logger = logging.getLogger(__name__)


class Consulta:
    """Classe para representar uma consulta"""
    def __init__(self, prestador: str, quantidade: str, data: str, servico: str, valor: str):
        self.prestador = prestador.strip() if prestador else ""
        self.quantidade = quantidade.strip() if quantidade else ""
        self.data = data.strip() if data else ""
        self.servico = servico.strip() if servico else ""
        self.valor = valor.strip() if valor else ""
    
    def __repr__(self):
        return f"Consulta(prestador={self.prestador}, data={self.data}, valor={self.valor})"


class Beneficiario:
    """Classe para representar um beneficiário e suas consultas"""
    def __init__(self, nome: str):
        self.nome = nome
        self.consultas: List[Consulta] = []
    
    def adicionar_consulta(self, consulta: Consulta):
        self.consultas.append(consulta)
    
    def __repr__(self):
        return f"Beneficiario(nome={self.nome}, consultas={len(self.consultas)})"


class Titular:
    """Classe para representar um titular e seus beneficiários"""
    def __init__(self, nome: str):
        self.nome = nome
        self.beneficiarios: List[Beneficiario] = []
        self.beneficiario_atual: Optional[Beneficiario] = None
    
    def adicionar_beneficiario(self, nome: str):
        """Adiciona um novo beneficiário e o torna o atual"""
        beneficiario = Beneficiario(nome)
        self.beneficiarios.append(beneficiario)
        self.beneficiario_atual = beneficiario
    
    def adicionar_consulta_ao_beneficiario_atual(self, consulta: Consulta):
        """Adiciona uma consulta ao beneficiário atual"""
        if self.beneficiario_atual:
            self.beneficiario_atual.adicionar_consulta(consulta)
        else:
            # Se não há beneficiário atual, cria um com o nome do titular
            self.adicionar_beneficiario(self.nome)
            self.beneficiario_atual.adicionar_consulta(consulta)
    
    def __repr__(self):
        return f"Titular(nome={self.nome}, beneficiarios={len(self.beneficiarios)})"


def is_ignored_value(valor: str) -> bool:
    """Verifica se um valor deve ser ignorado"""
    # Verificar se é NaN do pandas antes de converter para string
    if pd.isna(valor):
        return True
    
    if not valor or not isinstance(valor, str):
        return True
    
    valor_limpo = valor.strip()
    
    # Ignorar valores "nan" (case-insensitive) que vêm do pandas
    if valor_limpo.lower() in ["nan", "none", ""]:
        return True
    
    # Ignorar se for apenas números (pode ser código ou número de linha)
    if valor_limpo.isdigit():
        return True
    
    return valor_limpo in IGNORE_VALUES


def is_prestador(linha: pd.Series, coluna_evento: int, coluna_quantidade: int, 
                 coluna_data: int, coluna_servico: int, coluna_valor: int) -> bool:
    """
    Verifica se uma linha contém um prestador (nome comum, não é titular nem beneficiário)
    
    Args:
        linha: Linha do DataFrame
        coluna_evento: Índice da coluna "Evento"
        coluna_quantidade: Índice da coluna de quantidade
        coluna_data: Índice da coluna de data
        coluna_servico: Índice da coluna de serviço
        coluna_valor: Índice da coluna de valor
    
    Returns:
        True se for um prestador
    """
    evento = str(linha.iloc[coluna_evento]) if len(linha) > coluna_evento else ""
    evento = evento.strip()
    
    # Não é prestador se:
    # - Está vazio
    # - É valor ignorado
    # - Contém "Cód Titular:"
    # - Contém "Beneficiário:"
    # - É exatamente "Evento" ou "Prestador"
    
    if not evento or is_ignored_value(evento):
        return False
    
    if "Cód Titular:" in evento or "Beneficiário:" in evento:
        return False
    
    if evento.upper() in ["EVENTO", "PRESTADOR"]:
        return False
    
    # Verificar se tem dados válidos nas outras colunas (pelo menos uma deve ter valor)
    quantidade = str(linha.iloc[coluna_quantidade]) if len(linha) > coluna_quantidade else ""
    data = str(linha.iloc[coluna_data]) if len(linha) > coluna_data else ""
    servico = str(linha.iloc[coluna_servico]) if len(linha) > coluna_servico else ""
    valor = str(linha.iloc[coluna_valor]) if len(linha) > coluna_valor else ""
    
    # Se tem pelo menos um dado válido além do nome, é provavelmente um prestador
    tem_dados = (not is_ignored_value(quantidade) or 
                 not is_ignored_value(data) or 
                 not is_ignored_value(servico) or 
                 not is_ignored_value(valor))
    
    return tem_dados


def processar_xlsx(caminho_xlsx: Path) -> List[Titular]:
    """
    Processa o arquivo Excel (.xls ou .xlsx) e retorna lista de titulares com suas consultas
    
    Args:
        caminho_xlsx: Caminho para o arquivo Excel (.xls ou .xlsx)
    
    Returns:
        Lista de objetos Titular
    """
    try:
        # Detectar extensão do arquivo e escolher engine apropriado
        extensao = caminho_xlsx.suffix.lower()
        if extensao == '.xls':
            engine = 'xlrd'
        elif extensao in ['.xlsx', '.xlsm']:
            engine = 'openpyxl'
        else:
            # Tentar detectar automaticamente
            engine = None
        
        # Ler o arquivo Excel
        df = pd.read_excel(caminho_xlsx, engine=engine, header=None)
        
        # Identificar colunas (primeira linha geralmente tem os nomes)
        # Procurar pela linha com "Evento" na primeira coluna
        coluna_evento = None
        coluna_quantidade = None
        coluna_data = None
        coluna_servico = None
        coluna_valor = None
        
        # Procurar cabeçalhos
        for idx, row in df.iterrows():
            primeira_col = str(row.iloc[0]).strip() if len(row) > 0 else ""
            if primeira_col.upper() == "EVENTO":
                # Esta linha tem os nomes das colunas
                coluna_evento = 0
                coluna_quantidade = 1
                coluna_data = 2
                coluna_servico = 3
                coluna_valor = 4
                break
        
        # Se não encontrou cabeçalho, assumir posições padrão
        if coluna_evento is None:
            coluna_evento = 0
            coluna_quantidade = 1
            coluna_data = 2
            coluna_servico = 3
            coluna_valor = 4
        
        titulares: List[Titular] = []
        titular_atual: Optional[Titular] = None
        
        # Processar linha por linha
        for idx, row in df.iterrows():
            # Verificar se a célula existe e não é NaN antes de converter para string
            if len(row) <= coluna_evento:
                continue
            
            valor_evento = row.iloc[coluna_evento]
            
            # Ignorar se for NaN do pandas
            if pd.isna(valor_evento):
                continue
            
            # Ignorar se for apenas números (códigos, números de linha, etc.)
            try:
                # Tentar converter para número - se conseguir, é apenas número
                float(str(valor_evento).strip())
                continue
            except (ValueError, TypeError):
                # Não é número, continuar processamento
                pass
            
            evento = str(valor_evento)
            evento_upper = evento.upper().strip()
            
            # Ignorar linhas de cabeçalho (podem aparecer múltiplas vezes)
            if evento_upper in ["EVENTO", "PRESTADOR"]:
                continue
            
            # Ignorar linhas vazias ou com valores irrelevantes (incluindo "nan")
            if is_ignored_value(evento):
                continue
            
            # Verificar se é um titular
            nome_titular = extrair_nome_titular(evento)
            if nome_titular:
                titular_atual = Titular(nome_titular)
                titulares.append(titular_atual)
                logger.debug(f"Encontrado titular: {nome_titular}")
                continue
            
            # Verificar se é um beneficiário
            nome_beneficiario = extrair_nome_beneficiario(evento)
            if nome_beneficiario and titular_atual:
                titular_atual.adicionar_beneficiario(nome_beneficiario)
                logger.debug(f"Encontrado beneficiário: {nome_beneficiario} (titular: {titular_atual.nome})")
                continue
            
            # Verificar se é um prestador (linha com dados de consulta)
            if is_prestador(row, coluna_evento, coluna_quantidade, coluna_data, 
                           coluna_servico, coluna_valor) and titular_atual:
                prestador = evento
                quantidade = str(row.iloc[coluna_quantidade]) if len(row) > coluna_quantidade else ""
                data = str(row.iloc[coluna_data]) if len(row) > coluna_data else ""
                servico = str(row.iloc[coluna_servico]) if len(row) > coluna_servico else ""
                valor = str(row.iloc[coluna_valor]) if len(row) > coluna_valor else ""
                
                # Ignorar se todos os valores são irrelevantes
                if (is_ignored_value(quantidade) and is_ignored_value(data) and 
                    is_ignored_value(servico) and is_ignored_value(valor)):
                    continue
                
                consulta = Consulta(prestador, quantidade, data, servico, valor)
                titular_atual.adicionar_consulta_ao_beneficiario_atual(consulta)
                logger.debug(f"Adicionada consulta: {prestador} - {data} - {valor}")
        
        logger.info(f"Processados {len(titulares)} titulares do arquivo Excel")
        return titulares
    
    except Exception as e:
        logger.error(f"Erro ao processar arquivo Excel: {e}")
        raise

