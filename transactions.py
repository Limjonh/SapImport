# transactions.py
import os
import pandas as pd
from datetime import datetime
from sap_utils import wait_for_element, apply_layout_filter, paste_from_clipboard
from file_utils import import_txt_column_to_sap, create_empty_file, merge_files, export_ztmmq123_excel

def run_ztmmq123(session, config, params, log_callback, by_order=False):
    log_callback("\n--- Iniciando Transação ZTMMQ123 ---")

    # ============================================================================================
    # MODO 1 – EXECUÇÃO POR ORDEM (usa o arquivo IWBK)
    # ============================================================================================
    if by_order:
        log_callback("INFO: Executando ZTMMQ123 por Ordem.")

        session.StartTransaction("ZTMMQ123")
        wait_for_element(session, "wnd[0]/usr/radPA_LOEK1", log_callback)
        session.FindById("wnd[0]/usr/radPA_LOEK1").Select()

        # Abre seleção múltipla
        session.FindById("wnd[0]/usr/btn%_SE_BANFN_%_APP_%-VALU_PUSH").Press()

        # Importa coluna 4 (ordens) do arquivo IWBK
        iwbk_path = os.path.join(config.OUTPUT_DIR, config.FILENAMES['IWBK'])
        import_txt_column_to_sap(session, iwbk_path, 4, config.OUTPUT_DIR, config.FILENAMES['TEMP'], log_callback)

        # Executar
        session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
        wait_for_element(session, "wnd[0]/usr/cntlGRID1/shellcont/shell", log_callback)

        # Caso apareça janela de “lista muito grande”
        try:
            if session.ActiveWindow.Name == "wnd[1]":
                session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
        except:
            pass

        apply_layout_filter(session, config.LAYOUTS['ZTMMQ123'], log_callback)

        # Exportar XLSX (versão completa que usa o diretório da GUI)
        export_ztmmq123_excel(session, config, log_callback)

        log_callback("--- Transação ZTMMQ123 Concluída ---")
        return

    # ============================================================================================
    # MODO 2 – EXECUÇÃO POR PERÍODO + INTERVALO DE CENTROS
    # (L001 → T001  e  L001 → T011)
    # ============================================================================================
    start_date_from_gui = params.get('start_date')
    end_date_from_gui = params.get('end_date')

    if not start_date_from_gui or not end_date_from_gui:
        log_callback("❌ ERRO: Datas inválidas. Verifique as datas de início/fim.")
        return

    log_callback(f"INFO: Período filtrado: {start_date_from_gui} até {end_date_from_gui}")

    # Lista que será consolidada no final
    temp_files = []

    # Sempre roda apenas 1 execução usando intervalo
    filename = "ZTMMQ123.txt"
    filepath = os.path.join(config.OUTPUT_DIR, filename)
    temp_files.append(filepath)

    session.StartTransaction("ZTMMQ123")
    wait_for_element(session, "wnd[0]/usr/radPA_LOEK1", log_callback)

    session.FindById("wnd[0]/usr/radPA_LOEK1").Select()

    # ======================================================================
    # DEFINIR INTERVALO DE CENTROS (L001 → T001  e  L001 → T011)
    # ======================================================================
    log_callback("INFO: Inserindo intervalo de Centros (L001 → T001) e (L001 → T011)")

    session.findById("wnd[0]/usr/btn%_SE_WERKS_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL").select()

    # Primeira linha LOW: L001 → T001
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/"
                     "tblSAPLALDBINTERVAL/ctxtRSCSEL_255-ILOW_I[1,0]").text = "L001"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/"
                     "tblSAPLALDBINTERVAL/ctxtRSCSEL_255-ILOW_I[1,1]").text = "T001"

    # Segunda linha HIGH: L001 → T011
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/"
                     "tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").text = "L001"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/"
                     "tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,1]").text = "T011"

    # Confirma
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # ======================================================================
    # Datas
    # ======================================================================
    session.FindById("wnd[0]/usr/ctxtSE_BADAT-LOW").text = start_date_from_gui
    session.FindById("wnd[0]/usr/ctxtSE_BADAT-HIGH").text = end_date_from_gui

    # ======================================================================
    # STATUS
    # ======================================================================
    log_callback("INFO: Carregando status pré-definidos...")
    session.FindById("wnd[0]/usr/btn%_SE_STATU_%_APP_%-VALU_PUSH").Press()

    wait_for_element(session, "wnd[1]/usr/tabsTAB_STRIP", log_callback)

    for i, status in enumerate(config.ZTMMQ123_STATUS_LIST):
        session.FindById(
            f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
            f"ssubSCREEN_HEADER:SAPLALDB:3010/"
            f"tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{i}]"
        ).text = status

    session.FindById("wnd[1]/tbar[0]/btn[8]").Press()

    # Executar
    session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    wait_for_element(session, "wnd[0]/usr/cntlGRID1/shellcont/shell", log_callback)

    # Layout
    apply_layout_filter(session, config.LAYOUTS['ZTMMQ123'], log_callback)

    # ======================================================================
    # Exportação XLSX usando o método atualizado
    # ======================================================================
    export_ztmmq123_excel(session, config, log_callback)

    # ======================================================================
    # Merge final (caso queira unir vários centros)
    # ======================================================================
    output_path = os.path.join(config.OUTPUT_DIR, config.FILENAMES['ZTMMQ123'])
    merge_files(temp_files, output_path)

    # Remover temporários
    for f in temp_files:
        if os.path.exists(f):
            os.remove(f)

    log_callback("--- Transação ZTMMQ123 Concluída ---")


def run_me5a(session, config, params, log_callback, by_order=False):
    log_callback("\n--- Iniciando Transação ME5A ---")
    session.StartTransaction("ME5A")
    wait_for_element(session, "wnd[0]/tbar[1]/btn[17]", log_callback)
    session.FindById("wnd[0]/tbar[1]/btn[17]").Press()
    wait_for_element(session, "wnd[1]/usr/txtV-LOW", log_callback)
    
    # --- INÍCIO DA CORREÇÃO ---
    # Adiciona .setFocus() antes de definir o texto para garantir que o campo está editável
    variant_field = session.FindById("wnd[1]/usr/txtV-LOW")
    variant_field.setFocus()
    variant_field.text = config.LAYOUTS['ME5A']
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
    
    wait_for_element(session, "wnd[0]/usr/btn%_BA_BANFN_%_APP_%-VALU_PUSH", log_callback)
    session.FindById("wnd[0]/usr/btn%_BA_BANFN_%_APP_%-VALU_PUSH").Press()
    
    ztmmq123_path = os.path.join(
    config.OUTPUT_DIR,
    config.FILENAMES['ZTMMQ123']  # agora .xlsx
)

    import_txt_column_to_sap(
        session,
        ztmmq123_path,
        2,  # coluna B
        config.OUTPUT_DIR,
        config.FILENAMES['TEMP'],
        log_callback
)

    
    field_lstub = session.FindById("wnd[0]/usr/ctxtP_LSTUB")
    field_lstub.setFocus()
    field_lstub.text = "ALV"

    field_knttp_low = session.FindById("wnd[0]/usr/ctxtS_KNTTP-LOW")
    field_knttp_low.setFocus()
    field_knttp_low.text = "A"

    field_knttp_high = session.FindById("wnd[0]/usr/ctxtS_KNTTP-HIGH")
    field_knttp_high.setFocus()
    field_knttp_high.text = "Z"
    # --- FIM DA CORREÇÃO ---
    
    session.FindById("wnd[0]/usr/chkP_ZUGBA").Selected = True
    session.FindById("wnd[0]/usr/chkP_ERLBA").Selected = True
    session.FindById("wnd[0]/usr/chkP_BSTBA").Selected = True
    session.FindById("wnd[0]/usr/chkP_SELGS").Selected = True
    session.FindById("wnd[0]/usr/chkP_SELPO").Selected = True
    
    df = pd.DataFrame(params['centers'])
    df.to_clipboard(index=False, header=False)
    session.FindById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").Press()
    paste_from_clipboard(session, log_callback)
    
    session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    wait_for_element(session, "wnd[0]/usr/cntlGRID1/shellcont/shell", log_callback)
    
    apply_layout_filter(session, config.LAYOUTS['ME5A0'], log_callback)
    session.FindById("wnd[0]/tbar[1]/btn[45]").Press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    wait_for_element(session, "wnd[1]/usr/ctxtDY_PATH", log_callback)
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = config.OUTPUT_DIR
    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = config.FILENAMES['ME5A0']
    session.FindById("wnd[1]/tbar[0]/btn[11]").Press()
    
    session.FindById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").Select()
    apply_layout_filter(session, config.LAYOUTS['ME5A'], log_callback)


    session.FindById("wnd[0]/tbar[1]/btn[45]").Press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    wait_for_element(session, "wnd[1]/usr/ctxtDY_PATH", log_callback)
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = config.OUTPUT_DIR
    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = config.FILENAMES['ME5A']
    session.FindById("wnd[1]/tbar[0]/btn[11]").Press()
    
    log_callback("--- Transação ME5A Concluída ---")


def run_iw39(session, config, params, log_callback, by_order=False):
    log_callback("\n--- Iniciando Transação IW39 ---")

    session.StartTransaction("IW39")
    wait_for_element(session, "wnd[0]/mbar", log_callback)

    session.findById("wnd[0]").maximize()

    # Marca as opções (mantive igual ao seu original)
    session.FindById("wnd[0]/usr/chkDY_OFN").Selected = True
    session.FindById("wnd[0]/usr/chkDY_IAR").Selected = True
    session.FindById("wnd[0]/usr/chkDY_MAB").Selected = True
    session.FindById("wnd[0]/usr/chkDY_HIS").Selected = True

    # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    # SUBSTITUI ESTE BLOCO PELOS PASSOS QUE VOCÊ ENVIOU
    # Agora filtrando por AUART (tipos ZADM, ZAPO, ZCOR, ZMEL, ZPRV)
    # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    session.findById("wnd[0]/usr/btn%_AUART_%_APP_%-VALU_PUSH").press()
    wait_for_element(session, "wnd[1]/usr/tabsTAB_STRIP", log_callback)

    tabela = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"

    tipos_ordem = ["zadm", "zapo", "zcor", "zmel", "zprv"]

    for i, tipo in enumerate(tipos_ordem):
        cell = f"{tabela}/ctxtRSCSEL_255-SLOW_I[1,{i}]"
        session.findById(cell).text = tipo

    # último item recebe foco (mantendo igual ao seu passo)
    last_cell = f"{tabela}/ctxtRSCSEL_255-SLOW_I[1,{len(tipos_ordem)-1}]"
    session.findById(last_cell).setFocus()
    session.findById(last_cell).caretPosition = len(tipos_ordem[-1])

    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    # A partir daqui, tudo segue igual ao seu script original:
    # IMPORTAÇÃO POR ORDENS OU ME5A/IWBK
    if by_order:
        log_callback("INFO: Filtrando IW39 por Ordem.")
        df = pd.DataFrame(params['orders'])
        df.to_clipboard(index=False, header=False)
        paste_from_clipboard(session, log_callback, window_id="wnd[1]")
    else:
        log_callback("INFO: Filtrando IW39 por dados de ME5A e IWBK.")
        me5a_path = os.path.join(config.OUTPUT_DIR, config.FILENAMES['ME5A'])
        import_txt_column_to_sap(session, me5a_path, 1, config.OUTPUT_DIR, config.FILENAMES['TEMP'], log_callback)

        iwbk_path = os.path.join(config.OUTPUT_DIR, config.FILENAMES['IWBKe'])
        import_txt_column_to_sap(session, iwbk_path, 1, config.OUTPUT_DIR, config.FILENAMES['TEMP'], log_callback)

    # CENTROS
    session.FindById("wnd[0]/usr/btn%_SWERK_%_APP_%-VALU_PUSH").Press()
    df = pd.DataFrame(params['centers'])
    df.to_clipboard(index=False, header=False)
    paste_from_clipboard(session, log_callback)

    # DATAS – MANTIDAS da sua GUI
    session.FindById("wnd[0]/usr/ctxtDATUV").text = params['start_date']
    session.FindById("wnd[0]/usr/ctxtDATUB").text = params['end_date']

    # LAYOUT / VARIANTE
    session.FindById("wnd[0]/usr/ctxtVARIANT").text = config.LAYOUTS['IW39']

    # EXECUTAR
    session.FindById("wnd[0]/tbar[1]/btn[8]").Press()

    # EXPORTAÇÃO
    session.FindById("wnd[0]/mbar/menu[0]/menu[11]/menu[2]").Select()
    wait_for_element(session, "wnd[1]/tbar[0]/btn[0]", log_callback)
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press()

    wait_for_element(session, "wnd[1]/usr/ctxtDY_PATH", log_callback)
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = config.OUTPUT_DIR
    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = config.FILENAMES['IW39']
    session.FindById("wnd[1]/tbar[0]/btn[11]").Press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

    log_callback("--- Transação IW39 Concluída ---")


def run_iwbk(session, config, log_callback, params, by_order=False):

    import re
    import pandas as pd
    import os

    print("\n--- Iniciando Transação IWBK ---")

    # ==============================================================
    # 1) LER ARQUIVO IW39.TXT E EXTRAIR APENAS ORDENS VÁLIDAS
    # ==============================================================
    iw39_path = os.path.join(config.OUTPUT_DIR, config.FILENAMES['IW39'])

    if not os.path.exists(iw39_path):
        raise FileNotFoundError(
            f"Arquivo {iw39_path} não encontrado. Execute a IW39 antes da IWBK."
        )

    print(f"Lendo arquivo: {iw39_path}")

    ordens_validas = []

    with open(iw39_path, "r", encoding="utf-8") as f:
        for linha in f:
            # Procurar padrões de ordem REAL: 8 dígitos começando com 5
            encontrados = re.findall(r"\b5\d{7}\b", linha)

            # Se encontrou → adicionar à lista
            ordens_validas.extend(encontrados)

    if not ordens_validas:
        raise ValueError(
            "Nenhuma ordem válida (8 dígitos iniciando com 5) encontrada no arquivo IW39."
        )

    print(f"Total de ordens encontradas: {len(ordens_validas)}")

    # Remover duplicatas, caso existam
    ordens_validas = list(dict.fromkeys(ordens_validas))

    print(f"Total de ordens após remover duplicadas: {len(ordens_validas)}")

    # ==============================================================
    # 2) COPIAR ORDENS PARA O CLIPBOARD
    # ==============================================================
    df = pd.DataFrame(ordens_validas)
    df.to_clipboard(index=False, header=False)
    print("Ordens copiadas para a área de transferência.")

    # ==============================================================
    # 3) ABRIR A TRANSAÇÃO IWBK
    # ==============================================================
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "iwbk"
    session.findById("wnd[0]").sendVKey(0)

    # Abrir popup de múltipla seleção
    session.findById("wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").press()

    # ==============================================================
    # 4) COLAR AS ORDENS NO POPUP
    # ==============================================================
    print("Colando ordens no popup da IWBK...")
    
    session.findById("wnd[1]").sendVKey(24)  # Ctrl+V

    session.findById("wnd[1]/tbar[0]/btn[8]").press()  # Confirmar

    # ==============================================================
    # 5) EXECUTAR CONSULTA
    # ==============================================================
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # Agora tentamos pressionar o botão da janela popup, SE ela existir

    try:
        # Descobre quantas janelas existem
        app = session
        for i in range(5):  # tenta até wnd[4]
            try:
                win = app.findById(f"wnd[{i}]")
                
                # tenta achar botões comuns de confirmação
                for btn_id in ["tbar[0]/btn[0]", "tbar[0]/btn[1]", "tbar[0]/btn[6]", "tbar[0]/btn[11]"]:
                    try:
                        btn = win.findById(btn_id)
                        btn.press()
                        print(f"Janela extra detectada em wnd[{i}] → botão {btn_id} pressionado.")
                        raise Exception("handled")  # sai dos loops
                    except:
                        pass
            except:
                pass
    except Exception as handled:
        pass

    # ==============================================================
    # 6) SALVAR O ARQUIVO IWBK.TXT
    # ==============================================================
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = config.OUTPUT_DIR
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = config.FILENAMES['IWBK']

    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    print(f"Arquivo salvo como: {config.FILENAMES['IWBK']}")
    print("--- Transação IWBK Concluída ---")




def run_iw29(session, config, params, log_callback, by_order=False):
    log_callback("\n--- Iniciando Transação IW29 ---")
    session.StartTransaction("IW29")
    wait_for_element(session, "wnd[0]/mbar", log_callback)

    session.FindById("wnd[0]/usr/chkDY_OFN").Selected = True
    session.FindById("wnd[0]/usr/chkDY_RST").Selected = True
    session.FindById("wnd[0]/usr/chkDY_IAR").Selected = True
    session.FindById("wnd[0]/usr/chkDY_MAB").Selected = True

    session.FindById("wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").Press()
    wait_for_element(session, "wnd[1]/usr/tabsTAB_STRIP", log_callback)
    session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "50000000"

    if by_order:
        log_callback("INFO: Filtrando IW29 por Ordem.")
        df = pd.DataFrame(params['orders'])
        df.to_clipboard(index=False, header=False)
        paste_from_clipboard(session, log_callback, window_id="wnd[1]")
    else:
        log_callback("INFO: Executando IW29 sem filtro de ordem adicional.")
        session.FindById("wnd[1]/tbar[0]/btn[8]").Press()

    session.FindById("wnd[0]/usr/btn%_SWERK_%_APP_%-VALU_PUSH").Press()
    df = pd.DataFrame(params['centers'])
    df.to_clipboard(index=False, header=False)
    paste_from_clipboard(session, log_callback)

    for field_id in ["DATUV", "DATUB", "MZEIT-LOW", "MZEIT-HIGH", "AUZTV-LOW", "AUZTV-HIGH", "AUZTB-LOW", "AUZTB-HIGH"]:
        session.FindById(f"wnd[0]/usr/ctxt{field_id}").text = ""
    session.FindById("wnd[0]/usr/ctxtVARIANT").text = config.LAYOUTS['IW29']
    
    session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    wait_for_element(session, "wnd[0]/usr/cntlALV_CONTAINER_1/shellcont/shell", log_callback)

    session.FindById("wnd[0]/mbar/menu[0]/menu[11]/menu[2]").Select()
    wait_for_element(session, "wnd[1]/tbar[0]/btn[0]", log_callback)
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
    wait_for_element(session, "wnd[1]/usr/ctxtDY_PATH", log_callback)
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = config.OUTPUT_DIR
    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = config.FILENAMES['IW29']
    session.FindById("wnd[1]/tbar[0]/btn[11]").Press()

    log_callback("--- Transação IW29 Concluída ---")

def run_sqvi(session, config, params, log_callback, by_order=False):
    log_callback("\n--- Iniciando Transação SQVI ---")
    
    log_callback("INFO: Processando query NAST.")
    session.StartTransaction("SQVI")
    wait_for_element(session, "wnd[0]/usr/ctxtRS38R-QNUM", log_callback)
    session.FindById("wnd[0]/usr/ctxtRS38R-QNUM").text = "NAST"
    session.FindById("wnd[0]/usr/btnP1").Press()

    if "não foi criado" in session.FindById("wnd[0]/sbar").text:
        log_callback("AVISO: Query NAST não encontrada. Tentando criar...")
        session.FindById("wnd[0]/usr/btnP7").Press()
        wait_for_element(session, "wnd[1]/usr/txtRS38R-HDTITLE", log_callback)
        session.FindById("wnd[1]/usr/txtRS38R-HDTITLE").text = "Mensagens de Compra"
        session.FindById("wnd[1]/usr/subSUBSOURCE:SAPMS38R:3110/ctxtRS38Q-DDNAME").text = "NAST"
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
        wait_for_element(session, "wnd[0]/usr/tabsTAB100/tabpSELE", log_callback)
        session.FindById("wnd[0]/usr/tabsTAB100/tabpSELE").select()
        session.FindById("wnd[0]/tbar[0]/btn[3]").Press()
        wait_for_element(session, "wnd[1]/usr/btnBUTTON_1", log_callback)
        session.FindById("wnd[1]/usr/btnBUTTON_1").Press()
        
        wait_for_element(session, "wnd[0]/usr/ctxtRS38R-QNUM", log_callback)
        session.FindById("wnd[0]/usr/ctxtRS38R-QNUM").text = "NAST"
        session.FindById("wnd[0]/usr/btnP1").Press()

    wait_for_element(session, "wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH", log_callback)
    session.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH").Press()
    ztmmq123_path = os.path.join(config.OUTPUT_DIR, config.FILENAMES['ZTMMQ123'])
    import_txt_column_to_sap(session, ztmmq123_path, 30, config.OUTPUT_DIR, config.FILENAMES['TEMP'], log_callback)

    session.FindById("wnd[0]/usr/ctxtSP$00002-LOW").text = "1"
    session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    wait_for_element(session, "wnd[0]/usr/cntlCONTAINER/shellcont/shell", log_callback)
    
    shell = session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell")
    shell.PressToolbarContextButton("&MB_VARIANT")
    shell.SelectContextMenuItem("&LOAD")
    apply_layout_filter(session, config.LAYOUTS['NAST'], log_callback)
    
    shell.PressToolbarContextButton("&MB_EXPORT")
    shell.SelectContextMenuItem("&PC")
    wait_for_element(session, "wnd[1]/tbar[0]/btn[0]", log_callback)
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
    wait_for_element(session, "wnd[1]/usr/ctxtDY_PATH", log_callback)
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = config.OUTPUT_DIR
    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = config.FILENAMES['NAST']
    session.FindById("wnd[1]/tbar[0]/btn[11]").Press()
    session.FindById("wnd[0]/tbar[0]/btn[3]").Press()
    session.FindById("wnd[0]/tbar[0]/btn[3]").Press()

    log_callback("INFO: Processando query EKPO.")
    wait_for_element(session, "wnd[0]/usr/ctxtRS38R-QNUM", log_callback)
    session.FindById("wnd[0]/usr/ctxtRS38R-QNUM").text = "EKPO"
    session.FindById("wnd[0]/usr/btnP1").Press()
    
    wait_for_element(session, "wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH", log_callback)
    session.FindById("wnd[0]/usr/btn%_SP$00001_%_APP_%-VALU_PUSH").Press()
    
    paste_from_clipboard(session, log_callback, window_id="wnd[1]")

    session.FindById("wnd[0]/tbar[1]/btn[8]").Press()
    wait_for_element(session, "wnd[0]/usr/cntlCONTAINER/shellcont/shell", log_callback)
    
    shell = session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell")
    shell.PressToolbarContextButton("&MB_EXPORT")
    shell.SelectContextMenuItem("&PC")
    wait_for_element(session, "wnd[1]/tbar[0]/btn[0]", log_callback)
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
    wait_for_element(session, "wnd[1]/usr/ctxtDY_PATH", log_callback)
    session.FindById("wnd[1]/usr/ctxtDY_PATH").text = config.OUTPUT_DIR
    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = config.FILENAMES['EKPO']
    session.FindById("wnd[1]/tbar[0]/btn[11]").Press()

    log_callback("--- Transação SQVI Concluída ---")
