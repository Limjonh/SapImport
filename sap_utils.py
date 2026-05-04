# sap_utils.py
import win32com.client
import time

def connect_to_sap():
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        if not hasattr(sap_gui_auto, 'GetScriptingEngine'):
            return None
        engine = sap_gui_auto.GetScriptingEngine
        for connection in engine.Children:
            for session in connection.Children:
                return session
        return None
    except Exception:
        return None

def wait_for_element(session, element_id, log_callback, timeout=20):
    log_callback(f"  -> Aguardando elemento: {element_id}")
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            session.FindById(element_id)
            return True
        except:
            time.sleep(0.5)
    raise TimeoutError(f"Elemento '{element_id}' não foi encontrado após {timeout} segundos.")

def apply_layout_filter(session, layout_name, log_callback):
    """
    Aplica um filtro em uma variante (layout) e seleciona o resultado
    usando o método correto para evitar erros de 'propriedade somente leitura'.
    """
    log_callback(f"INFO: Aplicando filtro para o layout '{layout_name}'...")
    
    # 1. Clica no botão para selecionar o layout
    log_callback("  -> Clicando no botão 'Selecionar Layout' (btn[33])...")
    session.FindById("wnd[0]/tbar[1]/btn[33]").Press()
    
    # 2. Espera a janela de seleção de layout aparecer
    shell_path = "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell"
    wait_for_element(session, shell_path, log_callback)
    shell = session.FindById(shell_path)
    log_callback("  -> Janela de seleção de layout encontrada.")
    
    # --- INÍCIO DA ALTERAÇÃO: Seção de filtro comentada para teste ---
    # 3. Filtra a lista para encontrar o layout desejado
    # log_callback("  -> Iniciando filtragem do layout...")
    # shell.SelectColumn("VARIANT")
    # shell.ContextMenu()
    # time.sleep(0.5)
    # shell.SelectContextMenuItem("&FILTER")
    
    # 4. Preenche o nome do layout na janela de filtro
    # filter_field_path = "wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW"
    # wait_for_element(session, filter_field_path, log_callback)
    # session.FindById(filter_field_path).text = layout_name
    # session.FindById("wnd[2]/tbar[0]/btn[0]").Press() # Confirma o filtro
    # --- FIM DA ALTERAÇÃO ---
    
    # 5. Seleciona a linha resultante usando o método correto
    wait_for_element(session, shell_path, log_callback)
    
    # Move o foco para a primeira linha (que é o nosso resultado do filtro)
    log_callback("  -> Selecionando o primeiro layout da lista (linha 0)...")
    shell.currentCellRow = 0
    
    # Simula um duplo clique, que é a ação correta para selecionar e aplicar
    shell.doubleClickCurrentCell()
    log_callback("INFO: Layout aplicado com sucesso.")


def paste_from_clipboard(session, log_callback, window_id="wnd[1]"):
    log_callback("INFO: Colando dados da área de transferência no SAP...")
    session.FindById(f"{window_id}/tbar[0]/btn[24]").Press()
    session.FindById(f"{window_id}/tbar[0]/btn[8]").Press()

