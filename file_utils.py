# file_utils.py
"""
Módulo com funções para manipulação de ficheiros (TXT e Excel).
"""
import os
import re
import pandas as pd
from sap_utils import wait_for_element

def get_params_from_excel(excel_path, params_sheet, dates_sheet, order_table, rc_table, center_table):
    """
    Lê todos os parâmetros necessários da folha de cálculo Excel usando Pandas.
    Agora é robusto contra folhas de cálculo ou colunas vazias.
    """
    print("INFO: A ler parâmetros do ficheiro Excel...")
    params = {
        "start_date": None,
        "end_date": None,
        "orders": [],
        "rcs": [],
        "centers": []
    }

    try:
        # --- Leitura de Datas (com verificação de segurança) ---
        try:
            df_dates = pd.read_excel(excel_path, sheet_name=dates_sheet, header=None)
            # Verifica se o DataFrame tem pelo menos 2 linhas e 3 colunas
            if not df_dates.empty and df_dates.shape[0] > 1 and df_dates.shape[1] > 2:
                params["start_date"] = pd.to_datetime(df_dates.iloc[1, 1]).strftime('%d.%m.%Y')
                params["end_date"] = pd.to_datetime(df_dates.iloc[1, 2]).strftime('%d.%m.%Y')
            else:
                print(f"AVISO: A folha de cálculo '{dates_sheet}' está vazia ou mal formatada. As datas serão ignoradas.")
        except Exception as e:
            print(f"AVISO: Não foi possível ler as datas da folha de cálculo '{dates_sheet}'. Erro: {e}")

        # --- Leitura das Listas de Parâmetros (com verificação de segurança) ---
        df_params = pd.read_excel(excel_path, sheet_name=params_sheet)
        
        # Verifica se cada coluna existe antes de tentar ler
        if order_table in df_params.columns:
            params["orders"] = df_params[order_table].dropna().astype(str).tolist()
        else:
            print(f"AVISO: Coluna '{order_table}' não encontrada na folha de cálculo '{params_sheet}'.")

        if rc_table in df_params.columns:
            params["rcs"] = df_params[rc_table].dropna().astype(str).tolist()
        else:
            print(f"AVISO: Coluna '{rc_table}' não encontrada na folha de cálculo '{params_sheet}'.")
            
        if center_table in df_params.columns:
            params["centers"] = df_params[center_table].dropna().astype(str).tolist()
        else:
            print(f"AVISO: Coluna '{center_table}' não encontrada na folha de cálculo '{params_sheet}'.")

        print("INFO: Parâmetros carregados com sucesso.")
        return params

    except FileNotFoundError:
        print(f"ERRO CRÍTICO: O ficheiro Excel de parâmetros não foi encontrado em '{excel_path}'")
        return None
    except Exception as e:
        print(f"ERRO CRÍTICO ao ler o ficheiro Excel '{excel_path}': {e}")
        return None

import os
import re
import pandas as pd

def import_txt_column_to_sap(
    session,
    file_path,
    column_identifier,
    temp_dir,
    temp_filename,
    log_callback
):
    """
    Lê uma coluna específica de um ficheiro TXT ou XLSX,
    remove duplicados, guarda num ficheiro temporário e importa-o no SAP.
    Apenas valores numéricos com 8 dígitos que comecem com '10' ou '52' são considerados.
    """
    try:
        ext = os.path.splitext(file_path)[1].lower()
        values = set()
        VALID_PREFIXES = ("00", "10", "52")
        pattern = re.compile(rf"^({'|'.join(VALID_PREFIXES)})\d+$")

        # ======================================================
        # ORIGEM XLSX (nova)
        # ======================================================
        if ext == ".xlsx":
            log_callback("INFO: Fonte XLSX detectada.")

            # coluna B = índice 1
            df = pd.read_excel(
                file_path,
                usecols=[1],
                dtype=str
            )

            for value in df.iloc[:, 0].dropna():
                value = value.strip()
                if pattern.match(value):
                    values.add(value)

        # ======================================================
        # ORIGEM TXT (legado)
        # ======================================================
        elif ext == ".txt":
            log_callback("INFO: Fonte TXT detectada.")

            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()

            if len(lines) < 7:
                log_callback(
                    f"AVISO: Ficheiro '{os.path.basename(file_path)}' está vazio ou inválido."
                )
                return

            header_line = lines[4]
            headers = [h.strip() for h in header_line.split('|')]

            column_index = -1
            if isinstance(column_identifier, str):
                try:
                    column_index = headers.index(column_identifier)
                    log_callback(
                        f"INFO: Coluna '{column_identifier}' encontrada no índice {column_index}."
                    )
                except ValueError:
                    log_callback(
                        f"❌ ERRO: Coluna '{column_identifier}' não encontrada. Cabeçalhos: {headers}"
                    )
                    raise
            elif isinstance(column_identifier, int):
                column_index = column_identifier
                log_callback(
                    f"INFO: A usar índice de coluna fixo {column_index}."
                )

            if column_index == -1:
                raise ValueError("Identificador de coluna inválido.")

            for line in lines[6:]:
                parts = line.split('|')
                if len(parts) > column_index:
                    value = parts[column_index].strip()
                    if not value or value == "*":
                        continue
                    if pattern.match(value):
                        values.add(value)

        else:
            raise ValueError(f"Formato não suportado: {ext}")

        # ======================================================
        # VALIDAÇÃO FINAL
        # ======================================================
        if not values:
            log_callback(
                "⚠️ Nenhum valor válido encontrado (começando com '00' '10' ou '52')."
            )
            return

        # ======================================================
        # GERA TXT TEMPORÁRIO (igual ao fluxo antigo)
        # ======================================================
        temp_file_path = os.path.join(temp_dir, temp_filename)
        with open(temp_file_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(sorted(values)))

        # ======================================================
        # IMPORTAÇÃO NO SAP (INALTERADA)
        # ======================================================
        session.FindById("wnd[1]/tbar[0]/btn[23]").Press()
        wait_for_element(session, "wnd[2]/usr/ctxtDY_PATH", log_callback)
        session.FindById("wnd[2]/usr/ctxtDY_PATH").text = temp_dir
        session.FindById("wnd[2]/usr/ctxtDY_FILENAME").text = temp_filename
        session.FindById("wnd[2]/tbar[0]/btn[0]").Press()
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        log_callback(
            f"INFO: Importação concluída com {len(values)} valores válidos."
        )

    except FileNotFoundError:
        log_callback(
            f"AVISO: Ficheiro '{os.path.basename(file_path)}' não encontrado."
        )
    except Exception as e:
        log_callback(f"ERRO ao importar ficheiro: {e}")
        raise



def merge_files(file_paths, output_path):
    """
    Junta o conteúdo de vários ficheiros de texto num único ficheiro.
    """
    print(f"INFO: A juntar {len(file_paths)} ficheiros em '{os.path.basename(output_path)}'...")
    with open(output_path, 'w', encoding='utf-8') as outfile:
        for i, fname in enumerate(file_paths):
            if not os.path.exists(fname):
                continue
            with open(fname, 'r', encoding='utf-8', errors='ignore') as infile:
                if i > 0:
                    for _ in range(5):
                        next(infile, None)
                outfile.write(infile.read())
    print("INFO: Junção de ficheiros concluída.")

def create_empty_file(file_path, header=""):
    """
    Cria um ficheiro de texto vazio, opcionalmente com um cabeçalho.
    """
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(header)
    print(f"INFO: Ficheiro vazio criado: '{os.path.basename(file_path)}'")

def export_ztmmq123_excel(session, config, log_callback):
    try:
        log_callback("INFO: Exportando XLSX usando menu SAP...")

        # Lista > Exportar > Planilha
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()

        # Nome do arquivo
        filename = "ZTMMQ123.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename

        # Abre diálogo F4 para mudar diretório
        session.findById("wnd[1]").sendVKey(4)

        # Escolhe o diretório vindo da GUI
        session.findById("wnd[2]/usr/ctxtDY_PATH").text = config.OUTPUT_DIR
        session.findById("wnd[2]/tbar[0]/btn[0]").press()

        # Salva
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        log_callback(f"✔ Arquivo exportado para: {os.path.join(config.OUTPUT_DIR, filename)}")

    except Exception as e:
        log_callback(f"❌ ERRO ao exportar XLSX: {str(e)}")

