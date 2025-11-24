# Sistema de Envio de Emails de Consultas

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

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
- [Logs](#logs)
- [Tratamento de Erros](#tratamento-de-erros)
- [Contribuindo](#contribuindo)
- [Licença](#licença)

## Características

- Processamento automático de planilhas Excel (.xls e .xlsx)
- Correspondência inteligente de nomes entre Excel e CSV
- Geração de emails HTML formatados e responsivos
- Agrupamento automático de consultas por titular e beneficiário
- Modo de teste para validação antes do envio em massa
- Logging detalhado para rastreamento e depuração
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
git clone https://github.com/seu-usuario/nome-do-repositorio.git
cd nome-do-repositorio
```

### 2. Crie um ambiente virtual (recomendado)

```bash
python3 -m venv venv
source venv/bin/activate  # No Windows: venv\Scripts\activate
```

### 3. Instale as dependências

```bash
pip install -r requirements.txt
```

## Configuração

### 1. Configurar Credenciais de Email

O sistema utiliza SMTP do Gmail por padrão. Para configurar as credenciais:

1. Copie o arquivo de exemplo:
```bash
cp .env.example .env
```

2. Edite o arquivo `.env` e adicione suas credenciais:
```env
EMAIL_SENHA=sua_app_password_aqui
```

**Importante:** O arquivo `.env` não é versionado e contém informações sensíveis. Nunca compartilhe ou faça commit deste arquivo.

#### Como obter App Password do Gmail

1. Acesse [Google Account Security](https://myaccount.google.com/security)
2. Ative a verificação em duas etapas (se ainda não estiver ativa)
3. Gere uma "Senha de app" na seção "Senhas de app"
4. Utilize essa senha no arquivo `.env`

### 2. Configurar Email Remetente

Edite o arquivo `config/config.py` e altere o email remetente:

```python
EMAIL_REMETENTE = "seu_email@gmail.com"
```

### 3. Configurar Arquivos de Dados

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

### 4. Personalizar Caminhos (Opcional)

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
├── .env.example            # Exemplo de arquivo de configuração
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

O arquivo Excel (.xls ou .xlsx) deve apresentar a seguinte estrutura:

| Evento | Unnamed | Gp. Ap. | Doc. Finan. | Contrato Financeiro |
|--------|---------|---------|-------------|---------------------|
| Cód Titular: 123456 - NOME TITULAR | | | | |
| Beneficiário: 123456 - NOME BENEFICIÁRIO (00) | | | | |
| NOME PRESTADOR | 1 | 01/01/2024 | Consulta | 150.00 |

#### Padrões Reconhecidos

- **Titular**: `Cód Titular: [número] - NOME DO TITULAR`
- **Beneficiário**: `Beneficiário: [número] - NOME DO BENEFICIÁRIO ([código])`
- **Prestador**: Nome comum (sem prefixos especiais)

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

2. **Processamento do Excel**: O sistema processa o arquivo Excel linha por linha:
   - Identifica titulares (linhas com "Cód Titular:")
   - Identifica beneficiários (linhas com "Beneficiário:")
   - Identifica prestadores e suas consultas
   - Agrupa consultas por titular e beneficiário

3. **Normalização de Nomes**: Os nomes são normalizados (uppercase, espaços removidos) para garantir correspondência consistente entre o arquivo Excel e o CSV.

4. **Geração de HTML**: Para cada titular, o sistema gera um email HTML formatado contendo:
   - Nome do titular
   - Lista de beneficiários e suas consultas
   - Detalhes de cada consulta (prestador, data, tipo, quantidade, valor)
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

## Logs

O sistema gera logs em dois locais:

1. **Console**: Saída em tempo real do processamento
2. **Arquivo**: `logs/envio_emails.log` (produção) ou `logs/test_emails.log` (teste)

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

## Tratamento de Erros

O sistema trata automaticamente os seguintes cenários:

- Titulares sem email no CSV (registra e continua o processamento)
- Emails inválidos ou falhas de envio
- Erros de conexão SMTP
- Arquivos não encontrados
- Linhas vazias ou com valores irrelevantes
- Múltiplos encodings de CSV (tenta automaticamente)
- Nomes não encontrados (com logging detalhado)

## Contribuindo

Contribuições são bem-vindas. Para contribuir:

1. Faça um Fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

### Desenvolvimento

Para contribuir com o código:

```bash
# Clone o repositório
git clone https://github.com/seu-usuario/nome-do-repositorio.git

# Crie um ambiente virtual
python3 -m venv venv
source venv/bin/activate

# Instale as dependências
pip install -r requirements.txt

# Execute os testes
python test.py
```

## Licença

Este projeto está sob a licença MIT. Consulte o arquivo `LICENSE` para mais detalhes.

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
5. Abra uma [Issue](https://github.com/seu-usuario/nome-do-repositorio/issues) no GitHub

---
