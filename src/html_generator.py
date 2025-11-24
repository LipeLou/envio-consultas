"""
Gerador de HTML para emails de consultas
"""
import logging
from typing import List
from src.xlsx_parser import Titular, Beneficiario, Consulta
from datetime import datetime

logger = logging.getLogger(__name__)


def formatar_valor(valor: str) -> str:
    """Formata valor monetário"""
    if not valor or valor.strip() == "":
        return "R$ 0,00"
    
    try:
        # Tentar converter para float e formatar
        valor_float = float(str(valor).replace(",", "."))
        return f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(valor)


def formatar_data(data: str) -> str:
    """Formata data para exibição"""
    if not data or data.strip() == "":
        return ""
    
    # Tentar manter o formato original se já estiver formatado
    return data.strip()


def gerar_html_consultas(titular: Titular) -> str:
    """
    Gera HTML formatado com as consultas do titular
    
    Args:
        titular: Objeto Titular com beneficiários e consultas
    
    Returns:
        String HTML formatada
    """
    mes_atual = datetime.now().strftime("%B de %Y").title()
    mes_atual = mes_atual.replace("January", "Janeiro").replace("February", "Fevereiro").replace("March", "Março")
    mes_atual = mes_atual.replace("April", "Abril").replace("May", "Maio").replace("June", "Junho")
    mes_atual = mes_atual.replace("July", "Julho").replace("August", "Agosto").replace("September", "Setembro")
    mes_atual = mes_atual.replace("October", "Outubro").replace("November", "Novembro").replace("December", "Dezembro")
    
    html = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Consultas - {titular.nome}</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f4f4f4;
        }}
        .container {{
            background-color: #ffffff;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #2c3e50;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }}
        h2 {{
            color: #34495e;
            margin-top: 25px;
            margin-bottom: 15px;
            font-size: 1.2em;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 25px;
            background-color: #fff;
        }}
        th {{
            background-color: #3498db;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: bold;
        }}
        td {{
            padding: 10px;
            border-bottom: 1px solid #ddd;
        }}
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        tr:hover {{
            background-color: #f1f1f1;
        }}
        .valor {{
            text-align: right;
            font-weight: bold;
            color: #27ae60;
        }}
        .data {{
            white-space: nowrap;
        }}
        .servico {{
            color: #555;
        }}
        .prestador {{
            font-weight: 500;
            color: #2c3e50;
        }}
        .total {{
            background-color: #ecf0f1;
            font-weight: bold;
            font-size: 1.1em;
        }}
        .footer {{
            margin-top: 30px;
            padding-top: 20px;
            border-top: 2px solid #ecf0f1;
            color: #7f8c8d;
            font-size: 0.9em;
            text-align: center;
        }}
        .beneficiario-section {{
            margin-bottom: 30px;
            padding: 15px;
            background-color: #f8f9fa;
            border-left: 4px solid #3498db;
            border-radius: 5px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>Relatório de Consultas - {mes_atual}</h1>
        <p><strong>Titular:</strong> {titular.nome}</p>
"""
    
    # Calcular total geral
    total_geral = 0.0
    
    # Processar cada beneficiário
    for beneficiario in titular.beneficiarios:
        if not beneficiario.consultas:
            continue
        
        html += f"""
        <div class="beneficiario-section">
            <h2>Beneficiário: {beneficiario.nome}</h2>
            <table>
                <thead>
                    <tr>
                        <th>Prestador</th>
                        <th>Data</th>
                        <th>Serviço</th>
                        <th>Quantidade</th>
                        <th class="valor">Valor</th>
                    </tr>
                </thead>
                <tbody>
"""
        
        total_beneficiario = 0.0
        
        for consulta in beneficiario.consultas:
            # Calcular valor total da consulta
            try:
                qtd = float(str(consulta.quantidade).replace(",", ".")) if consulta.quantidade else 1.0
                valor_unit = float(str(consulta.valor).replace(",", ".")) if consulta.valor else 0.0
                valor_total = qtd * valor_unit
                total_beneficiario += valor_total
                total_geral += valor_total
            except:
                valor_total = 0.0
            
            html += f"""
                    <tr>
                        <td class="prestador">{consulta.prestador}</td>
                        <td class="data">{formatar_data(consulta.data)}</td>
                        <td class="servico">{consulta.servico}</td>
                        <td>{consulta.quantidade}</td>
                        <td class="valor">{formatar_valor(consulta.valor)}</td>
                    </tr>
"""
        
        html += f"""
                </tbody>
                <tfoot>
                    <tr class="total">
                        <td colspan="4" style="text-align: right;"><strong>Total do Beneficiário:</strong></td>
                        <td class="valor">{formatar_valor(str(total_beneficiario))}</td>
                    </tr>
                </tfoot>
            </table>
        </div>
"""
    
    # Total geral
    html += f"""
        <table>
            <tfoot>
                <tr class="total">
                    <td colspan="4" style="text-align: right;"><strong>TOTAL GERAL:</strong></td>
                    <td class="valor">{formatar_valor(str(total_geral))}</td>
                </tr>
            </tfoot>
        </table>
        
        <div class="footer">
            <p>Este é um email automático. Por favor, não responda.</p>
            <p>Relatório gerado automaticamente pelo sistema de envio de consultas.</p>
        </div>
    </div>
</body>
</html>
"""
    
    return html

