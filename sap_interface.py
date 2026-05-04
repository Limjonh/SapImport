# sap_interface.py
"""
Este módulo serve como a interface entre o front-end (GUI) e o back-end (lógica de automação).
"""
import sys
import os
import config
from sap_utils import connect_to_sap
from transactions import (
    run_ztmmq123, run_me5a, run_iwbk, run_sqvi, run_iw39, run_iw29
)

def executar_consultas_sap(pasta_destino, transacoes_selecionadas, log_callback, centros, start_date, end_date):
    """
    Função principal que orquestra toda a automação, chamada pela GUI.
    """
    log_callback("=============================================")
    log_callback("=      INICIANDO ROBÔ DE EXTRAÇÃO SAP       =")
    log_callback("=============================================")

    # 1. Constrói o dicionário de parâmetros diretamente a partir da GUI.
    params = {
        "start_date": start_date,
        "end_date": end_date,
        "centers": centros,
        "orders": [],
        "rcs": []
    }
    
    # 2. Conecta ao SAP
    log_callback("\nINFO: Conectando ao SAP...")
    session = connect_to_sap()
    if session is None:
        log_callback("ERRO: Falha ao conectar ao SAP. Verifique se o SAP Logon está aberto.")
        raise ConnectionError("Não foi possível conectar à sessão SAP.")
    log_callback("INFO: Sessão SAP encontrada e conectada.")

    # 3. Define o modo de execução como "Geral" (não por ordem)
    run_by_order = False
    log_callback("\nMODO DE EXECUÇÃO: Por Centro e Data (definido na interface)")

    # Atualiza o config com o diretório de saída escolhido na GUI
    config.OUTPUT_DIR = pasta_destino
    
    # 4. Mapeia o nome da transação para a sua função correspondente
    funcoes_transacoes = {
        "ZTMMQ123": run_ztmmq123,
        "ME5A": run_me5a,
        "IW39": run_iw39,
        "IWBK": run_iwbk,
        "IW29": run_iw29,
        "SQVI": run_sqvi,
    }

    # 5. Executa as transações selecionadas na GUI
    for transacao_nome in transacoes_selecionadas:
        if transacao_nome in funcoes_transacoes:
            funcao_a_executar = funcoes_transacoes[transacao_nome]
            try:
                # Primeiro, executa a transação e salva o arquivo
                funcao_a_executar(session, config, params, log_callback, by_order=run_by_order)
            except Exception as e:
                log_callback(f"ERRO ao executar a transação {transacao_nome}: {e}")
                raise e
        else:
            log_callback(f"AVISO: Nenhuma função de automação encontrada para '{transacao_nome}'.")

    log_callback("\n=============================================")
    log_callback("=      CONSULTAS SAP FINALIZADAS        =")
    log_callback("=============================================")