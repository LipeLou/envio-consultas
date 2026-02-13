# Sistema de Envio de Emails de Consultas

[![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Pandas-1.3%2B-150458?logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![Licença: MIT](https://img.shields.io/badge/Licen%C3%A7a-MIT-green.svg)](LICENSE)

Sistema automatizado para processar planilhas Excel (.xls ou .xlsx) com consultas médicas e enviar emails HTML formatados para cada titular do plano de saúde.

> **Nota:** Este sistema foi desenvolvido para processar relatórios no formato padrão da UNIMED. O arquivo Excel de entrada deve seguir a estrutura de relatórios fornecidos pela UNIMED.

## Índice

- [Características](#características)
- [Pré-requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Uso](#uso)
- [Estrutura do Projeto](#estrutura-do-projeto)
- [Estrutura dos Arquivos](#estrutura-dos-arquivos)
- [Como Funciona](#como-funciona)
- [Modo de Teste](#modo-de-teste)
- [Logs e Relatórios](#logs-e-relatorios)
- [Tratamento de Erros](#tratamento-de-erros)
- [Ferramentas](#ferramentas)
- [Suporte](#suporte)

## Características

- Processamento automático de planilhas Excel (.xls e .xlsx)
- Correspondência inteligente de nomes entre Excel e CSV
- Geração de emails HTML formatados e responsivos
- Agrupamento automático de consultas por titular e beneficiário
- Modo de teste para validação antes do envio em massa
- Logging detalhado para rastreamento e depuração
- Geração de relatório CSV de erros (emails não encontrados ou falhas de envio)
- Tratamento robusto de erros e encoding
- Suporte a múltiplos encodings de CSV

## Pré-requisitos

- Python 3.8 ou superior
- Conta Gmail com App Password configurada (ou outro servidor SMTP compatível)
- Arquivo Excel com consultas no formato especificado
- Arquivo CSV com emails dos titulares

## Instalação

### 1. Clone o repositório

```bash
git clone https://github.com/lipelou/envio-consultas
cd nome-do-repositorio
```

### 2. Crie um ambiente virtual (recomendado)

```bash
python3 -m venv venv
venv\Scripts\activate # Linux/Mac: source venv/bin/activate
```

### 3. Instale as dependências

```bash
pip install -r requirements.txt
```

## Configuração

### 1. Configurar Credenciais de Email

O sistema utiliza SMTP do Gmail por padrão. Para configurar as credenciais:

1. Crie um arquivo `.env` na raiz do projeto:
```bash
touch .env
```

2. Edite o arquivo `.env` e adicione sua senha:
```env
EMAIL_SENHA=sua_app_password_aqui
```

#### Como obter App Password do Gmail

1. Acesse [Google Account Security](https://myaccount.google.com/security)
2. Ative a verificação em duas etapas (se ainda não estiver ativa)
3. Gere uma "Senha de app" na seção "Senhas de app"
4. Utilize essa senha no arquivo `.env`

### 2. Configurar Email Remetente e Logo

Edite o arquivo `config/config.py` e configure o email remetente e a URL da logo:

```python
EMAIL_REMETENTE = "seu_email@gmail.com"
LOGO_URL = "https://exemplo.com/sua-logo.png"  # Deixe vazio "" para não usar logo
```

### 3. Configurar Mês do Relatório

Edite o arquivo `config/config.py` e configure o mês/ano que será exibido no relatório:

```python
MES_RELATORIO = "Mês de Ano"  # ALTERE AQUI: Configure o mês/ano do relatório
```

### 4. Configurar Arquivos de Dados

Coloque os arquivos necessários nos diretórios apropriados:

- `data/consultas.xls` ou `data/consultas.xlsx` - Arquivo Excel com as consultas
- `data/email.csv` - Arquivo CSV com os emails dos titulares

**Estrutura do CSV (`email.csv`):**
```csv
NOME;E-MAIL PRINCIPAL;OUTRO E-MAIL
JOAO SILVA;joao@example.com;joao.alt@example.com
MARIA SANTOS;maria@example.com;
```

**Observação:** O sistema prioriza "E-MAIL PRINCIPAL" e utiliza "OUTRO E-MAIL" como alternativa.

### 5. Personalizar Caminhos (Opcional)

Se necessário, ajuste os caminhos no arquivo `config/config.py`:

```python
DATA_DIR = BASE_DIR / "data"
CSV_EMAILS_PATH = DATA_DIR / "email.csv"
XLSX_CONSULTAS_PATH = DATA_DIR / "consultas.xls"
```

## Uso

### Uso Básico

Execute o script principal:

```bash
python main.py
```

O sistema realizará as seguintes operações:
1. Leitura do arquivo CSV de emails
2. Processamento do arquivo Excel de consultas
3. Geração e envio de emails HTML para cada titular

### Modo de Teste

Recomenda-se testar o sistema antes do envio em massa:

1. **Configure o email de teste** no arquivo `config/config_test.py`:
```python
EMAIL_TESTE = "seu_email_pessoal@gmail.com"
```

2. **Execute o script de teste**:
```bash
python test.py
```

O script processará apenas os **3 primeiros titulares** e enviará os emails para o endereço configurado em `EMAIL_TESTE`.

3. **Personalizar o limite de titulares**:
```bash
python test.py 5  # Processa 5 titulares
```

**Importante:** Todos os emails de teste serão enviados para `EMAIL_TESTE`, independentemente do email real do titular.

## Estrutura do Projeto

```
.
├── main.py                  # Script principal
├── test.py                  # Script de teste
├── requirements.txt          # Dependências Python
├── README.md               # Este arquivo
├── src/                    # Código fonte
│   ├── __init__.py
│   ├── email_utils.py      # Utilitários para leitura de emails
│   ├── xlsx_parser.py      # Parser do arquivo Excel
│   ├── html_generator.py   # Gerador de HTML para emails
│   └── email_sender.py     # Módulo de envio de emails
├── config/                 # Configurações
│   ├── __init__.py
│   ├── config.py          # Configurações do sistema
│   └── config_test.py     # Configurações de teste
├── data/                   # Arquivos de dados (não versionados)
│   ├── .gitkeep
│   ├── email.csv          # CSV com emails (adicionar aqui)
│   └── consultas.xls      # Excel com consultas (adicionar aqui)
└── logs/                   # Arquivos de log (não versionados)
    └── .gitkeep
```

## Estrutura dos Arquivos

### Arquivo Excel de Consultas

> **Importante:** Este sistema foi desenvolvido para processar relatórios no formato padrão da UNIMED. O arquivo Excel de entrada deve seguir a estrutura de relatórios fornecidos pela UNIMED.

O arquivo Excel (.xls ou .xlsx) apresenta a seguinte estrutura:

| Coluna A | Coluna B | Coluna C | Coluna D | Coluna E |
|----------|----------|----------|----------|----------|
| **Linha 1 (Evento)** | Código Evento | (Vazio) | Grupo Serviço | Doc. Finan. | Contrato |
| **Linha 2 (Detalhe)** | Qtde | Dt. Real. | (Vazio) | Total Cobr. | (Vazio) |

Exemplo visual:

```text
Cód Titular: 123456 - NOME TITULAR
Beneficiário: 123456 - NOME BENEFICIÁRIO (00)
8405240         |           | CONSULTAS | 59426.0     | 1.37...
1,0000          | 04/12/2025|           | 59.95       |
```

#### Padrões Reconhecidos

- **Titular**: `Cód Titular: [número] - NOME DO TITULAR`
- **Beneficiário**: `Beneficiário: [número] - NOME DO BENEFICIÁRIO ([código])`
- **Transação**: Bloco de duas linhas contendo código do evento, serviço, quantidade, data e valor.

### Arquivo CSV de Emails

```csv
NOME;E-MAIL PRINCIPAL;OUTRO E-MAIL
JOAO SILVA;joao@example.com;joao.alt@example.com
MARIA SANTOS;maria@example.com;
```

**Colunas obrigatórias:**
- `NOME`: Nome completo do titular
- `E-MAIL PRINCIPAL`: Email principal (prioritário)
- `OUTRO E-MAIL`: Email alternativo (opcional)

## Como Funciona

1. **Leitura do CSV**: O sistema lê o arquivo `email.csv` e cria um dicionário mapeando nomes normalizados para emails.

2. **Processamento do Excel**: O sistema processa o arquivo Excel varrendo as linhas:
   - Identifica titulares (linhas com "Cód Titular:")
   - Identifica beneficiários (linhas com "Beneficiário:")
   - Identifica transações médicas (blocos de duas linhas com detalhes do evento)
   - Agrupa consultas por titular e beneficiário

3. **Normalização de Nomes**: Os nomes são normalizados (uppercase, espaços removidos) para garantir correspondência consistente entre o arquivo Excel e o CSV.

4. **Geração de HTML**: Para cada titular, o sistema gera um email HTML formatado contendo:
   - Nome do titular
   - Lista de beneficiários e suas consultas
   - Detalhes de cada consulta (data, quantidade, valor)
   - Totais por beneficiário e total geral

5. **Envio de Emails**: O sistema envia o email HTML para o endereço do titular encontrado no CSV via SMTP.

## Modo de Teste

O modo de teste permite validar o sistema antes do envio em massa:

```bash
python test.py [limite]
```

**Características do modo de teste:**
- Processa apenas os primeiros N titulares (padrão: 3)
- Todos os emails são enviados para `EMAIL_TESTE` configurado
- Exibe no log o email real do titular (para verificação)
- Não envia emails reais para os titulares

## Logs e Relatórios

O sistema gera logs e relatórios em dois formatos:

### 1. Logs de Processamento (Texto)

- **Console**: Saída em tempo real do processamento
- **Arquivo**: `logs/envio_emails.log` (produção) ou `logs/test_emails.log` (teste)

**Exemplo de saída:**
```
============================================================
Iniciando processamento de consultas e envio de emails
============================================================
Lendo arquivo CSV de emails...
Carregados 150 emails do arquivo CSV
Processando arquivo Excel: data/consultas.xls
Processados 100 titulares do arquivo Excel

Processando 100 titulares...
------------------------------------------------------------

[1/100] Processando titular: JOAO SILVA
  - 2 beneficiário(s)
  - 3 consulta(s) total(is)
  - Email encontrado: joao@example.com
  - Enviando email...
  - Email enviado com sucesso!

============================================================
RESUMO DO PROCESSAMENTO
============================================================
Total de titulares processados: 100
Emails enviados com sucesso: 95
Emails falhados: 2
Titulares sem email: 3
============================================================
```

### 2. Relatório de Erros (CSV)

Sempre que houver falhas no envio ou titulares sem email encontrado, o sistema gera automaticamente um arquivo CSV na pasta `logs/`.

- **Formato do nome**: `nao_enviados_AAAA-MM-DD_HH-MM-SS.csv` (ou `nao_enviados_TESTE_...` no modo de teste)
- **Conteúdo do CSV**:
  - `titular`: Nome do titular
  - `motivo`: Motivo da falha (ex: "Email não encontrado", "Falha no envio SMTP")
  - `detalhe`: Informações adicionais (ex: nome normalizado buscado)
  - `email`: Email tentado (se houver)

Este relatório facilita a identificação e correção de cadastros faltantes ou incorretos.

## Tratamento de Erros

O sistema trata automaticamente os seguintes cenários:

- Titulares sem email no CSV (registra e continua o processamento)
- Emails inválidos ou falhas de envio
- Erros de conexão SMTP
- Arquivos não encontrados
- Linhas vazias ou com valores irrelevantes
- Múltiplos encodings de CSV (tenta automaticamente)
- Nomes não encontrados (com logging detalhado)

## Ferramentas

- [pandas](https://pandas.pydata.org/) - Processamento de dados
- [openpyxl](https://openpyxl.readthedocs.io/) - Leitura de arquivos Excel (.xlsx)
- [xlrd](https://xlrd.readthedocs.io/) - Leitura de arquivos Excel (.xls)
- [python-dotenv](https://github.com/theskumar/python-dotenv) - Gerenciamento de variáveis de ambiente

## Suporte

Em caso de problemas, verifique:

1. Se a senha do email está configurada corretamente
2. Se o arquivo Excel está no formato esperado
3. Se os nomes no Excel correspondem aos nomes no CSV
4. Os logs em `logs/envio_emails.log` para detalhes de erros

---

**Desenvolvido para o SINTUNIFEI** | Sistema de Envio de Relatórios - **Consultas Unimed**.