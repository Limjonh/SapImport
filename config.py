# config.py
"""
Arquivo de configuração central para a automação SAP.
Altere os parâmetros aqui para ajustar o comportamento do robô sem mexer na lógica.
"""
import os

# --- CONFIGURAÇÕES GERAIS ---
# Caminho para a planilha de parâmetros.
# O 'os.path.dirname(__file__)' pega o diretório onde o script está rodando.
# Altere "sua_planilha.xlsx" para o nome real do seu arquivo Excel.
EXCEL_PATH = os.path.join(os.path.dirname(__file__), "sua_planilha.xlsx")

# Diretório onde os arquivos TXT serão salvos (geralmente a mesma pasta do projeto).
OUTPUT_DIR = os.path.dirname(__file__)

# --- NOMES DAS ABAS E TABELAS NO EXCEL ---
PARAMS_SHEET_NAME = "Parâmetros"
DATES_SHEET_NAME = "Planilha2" # Aba onde ficam as datas
ORDER_TABLE_NAME = "Ordem"
RC_TABLE_NAME = "NRC"
CENTER_TABLE_NAME = "CL"

# --- NOMES DOS ARQUIVOS DE SAÍDA (TXT) ---
FILENAMES = {
    "ZTMMQ123": "ZTMMQ123.XLSX",
    "IWBK": "IWBK.TXT",
    "IWBKe": "IWBKe.TXT",
    "ME5A": "ME5A.TXT",
    "ME5A0": "ME5A0.TXT",
    "IW39": "IW39.TXT",
    "IW29": "IW29.TXT",
    "EKPO": "EKPO.TXT",
    "NAST": "NAST.TXT",
    "TEMP": "temp.TXT"
}

# --- LAYOUTS (VARIANTES) UTILIZADOS NO SAP ---
LAYOUTS = {
    "ZTMMQ123": "/SESIEPM",
    "ME5A": "/sesiepm",
    "ME5A0": "/sesiepm",
    "IWBK": "/sesiepm",
    "IWBK_ORDEM": "/sesiepm2",
    "IW39": "/rjmg_acomp",
    "IW29": "/rjmg_acomp",
    "NAST": "/PRMN"
}

# --- PARÂMETROS FIXOS ---
# Status para a transação ZTMMQ123
ZTMMQ123_STATUS_LIST = ["N", "B", "A", "D", "E"]

# config.py

# --- CONFIGURAÇÃO PARA EXPORTAÇÃO FINAL EM EXCEL ---

# Nome do arquivo Excel que será criado com todos os dados.
DESTINATION_EXCEL_PATH = os.path.join(OUTPUT_DIR, "Relatorio_Final_SAP.xlsx")

# Mapeia cada arquivo TXT para o nome da aba na planilha final.
# Você pode customizar os nomes das abas como preferir.
SHEET_MAPPING = {
    "ZTMMQ123": "Dados ZTMMQ123",
    "ME5A": "Dados ME5A",
    "ME5A0": "Dados ME5A0",
    "IWBK": "Dados IWBK",
    "IW39": "Dados IW39",
    "IW29": "Dados IW29",
    "EKPO": "Dados EKPO",
    "NAST": "Dados NAST",
}