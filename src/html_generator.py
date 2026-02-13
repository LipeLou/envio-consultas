"""
Gerador de HTML para emails de consultas
"""
import logging
from src.xlsx_parser import Titular, Beneficiario, Consulta
from config.config import MES_RELATORIO, LOGO_URL

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
    # Usar mês configurado manualmente
    mes_atual = MES_RELATORIO
    
    # Construir tag da logo
    img_tag = ""
    if LOGO_URL:
        img_tag = f'<img src="{LOGO_URL}" alt="Logo" style="max-width: 100px; max-height: 60px; float: right; margin-left: 20px;">'

    html = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Consultas UNIMED</title>
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
        .header-container {{
            border-bottom: 3px solid #06679a;
            padding-bottom: 10px;
            margin-bottom: 20px;
            overflow: hidden;
        }}
        h1 {{
            color: #2c3e50;
            margin: 0;
            padding: 0;
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
            background-color: #006600;
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
        }}
        .data {{
            white-space: nowrap;
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
            border-left: 4px solid #006600;
            border-radius: 5px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header-container">
            {img_tag}
            <h1>Relatório de Consultas UNIMED - {mes_atual}</h1>
        </div>
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
                        <th>Data</th>
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
                valor = float(str(consulta.valor).replace(",", ".")) if consulta.valor else 0.0
                total_beneficiario += valor
                total_geral += valor
            except:
                valor = 0.0
            
            html += f"""
                    <tr>
                        <td class="data">{formatar_data(consulta.data)}</td>
                        <td>{consulta.quantidade}</td>
                        <td class="valor">{formatar_valor(consulta.valor)}</td>
                    </tr>
"""
        
        html += f"""
                </tbody>
                <tfoot>
                    <tr class="total">
                        <td colspan="2" style="text-align: right;"><strong>Total do Beneficiário:</strong></td>
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
                    <td colspan="2" style="text-align: right;"><strong>TOTAL GERAL:</strong></td>
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

