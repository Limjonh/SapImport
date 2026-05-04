# main.py
"""
Script principal para orquestrar a automação de extração de dados do SAP.
Este é o ponto de entrada do robô.
"""
import os
import sys
import pandas as pd

# Importa as configurações e funções dos outros módulos
import config
from sap_utils import connect_to_sap
from file_utils import get_params_from_excel, write_all_txt_to_excel
# Importa TODAS as funções de transação que serão usadas no fluxo
from transactions import run_ztmmq123, run_me5a, run_iwbk, run_sqvi, run_iw39, run_iw29

def main():
    """
    Função principal que executa todo o fluxo da automação.
    """
    print("=============================================")
    print("=      INICIANDO ROBÔ DE EXTRAÇÃO SAP       =")
    print("=============================================")

    # 1. Carregar parâmetros do Excel
    params = get_params_from_excel(
        config.EXCEL_PATH,
        config.PARAMS_SHEET_NAME,
        config.DATES_SHEET_NAME,
        config.ORDER_TABLE_NAME,
        config.RC_TABLE_NAME,
        config.CENTER_TABLE_NAME
    )
    if params is None:
        sys.exit("ERRO: Falha ao carregar parâmetros. Abortando execução.")

    # 2. Conectar ao SAP
    session = connect_to_sap()
    if session is None:
        sys.exit("ERRO: Falha ao conectar ao SAP. Abortando execução.")

    # 3. Determinar o modo de execução (por ordem ou por centro/data)
    # A lógica original verifica se há mais de uma ordem na tabela.
    run_by_order = len(params['orders']) > 0

    # 4. Executar a sequência de transações
    try:
        # A sequência exata de chamadas depende do seu fluxo de trabalho.
        # Este é um exemplo baseado no script VBA.
        
        if run_by_order:
            print("\nMODO DE EXECUÇÃO: Por Ordem")
            # run_iwbk(session, config, params, by_order=True) # Exemplo de como seria
            run_ztmmq123(session, config, params, by_order=True)
        else:
            print("\nMODO DE EXECUÇÃO: Por Centro e Data")
            run_ztmmq123(session, config, params, by_order=False)
        
        run_me5a(session, config, params)
        
        # À medida que as funções forem implementadas no arquivo 'transactions.py',
        # você pode descomentar as linhas abaixo para executá-las em sequência.
        run_sqvi(session, config, params)
        run_iw39(session, config, params)
        run_iwbk(session, config, params)
        run_iw29(session, config, params)
        
        # ================================================================= #
        # NOVA ETAPA: Escrever todos os resultados em um único arquivo Excel #
        # ================================================================= #
        write_all_txt_to_excel(
            config.OUTPUT_DIR,
            config.FILENAMES,
            config.SHEET_MAPPING,
            config.DESTINATION_EXCEL_PATH
        )
        
    except Exception as e:
        print("\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print(f"ERRO INESPERADO DURANTE A EXECUÇÃO: {e}")
        print("O robô encontrou um problema e precisou parar.")
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    finally:
        # 5. Limpeza final
        temp_file = os.path.join(config.OUTPUT_DIR, config.FILENAMES['TEMP'])
        if os.path.exists(temp_file):
            os.remove(temp_file)
            print(f"\nINFO: Arquivo temporário '{config.FILENAMES['TEMP']}' removido.")
            
        print("\n=============================================")
        print("=        EXECUÇÃO DO ROBÔ FINALIZADA        =")
        print("=============================================")

# Ponto de entrada padrão para scripts Python
if __name__ == "__main__":
    main()
