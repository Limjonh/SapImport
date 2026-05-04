# excel_writer.py (Versão 2.0 - Mais Robusta)
"""
Módulo responsável por ler arquivos TXT com diferentes estruturas e
consolidá-los em um único arquivo Excel.
"""
import os
import pandas as pd
import config

def analisar_e_ler_txt(file_path):
    """
    Analisa a estrutura de um arquivo TXT para determinar como lê-lo
    corretamente com o pandas, especialmente para arquivos de largura fixa.
    Retorna um DataFrame limpo.
    """
    try:
        with open(file_path, 'r', encoding='latin1') as f:
            lines = f.readlines()
    except (FileNotFoundError, UnicodeDecodeError) as e:
        print(f"Erro ao ler o arquivo {file_path}: {e}")
        return pd.DataFrame()

    header_line_index = -1
    separator_line_index = -1
    footer_line_index = -1

    for i, line in enumerate(lines):
        # Encontra a linha de cabeçalho
        if '----' in line and separator_line_index == -1:
            separator_line_index = i
            if i > 0:
                header_line_index = i - 1
        
        # Encontra a linha do rodapé
        if 'Total de entradas selecionadas:' in line:
            footer_line_index = i
            break
            
    if header_line_index == -1 or separator_line_index == -1:
        print(f"AVISO: Cabeçalho ou separador não encontrado em {file_path}. Pulando.")
        return pd.DataFrame()

    # Determina a largura de cada coluna a partir da linha de separadores
    separator_line = lines[separator_line_index]
    col_starts = []
    
    # Encontra as posições dos separadores "|"
    for i, char in enumerate(separator_line):
        if char == '|':
            col_starts.append(i)

    # Se a última coluna for a última da linha, adicionamos o final
    if len(separator_line.strip()) > 0 and separator_line.strip()[-1] == '|':
        col_starts.append(len(separator_line) - 1)
        
    if not col_starts:
        print("AVISO: Linha de separadores inválida. Pulando.")
        return pd.DataFrame()

    col_widths = [col_starts[i+1] - col_starts[i] for i in range(len(col_starts)-1)]
    
    # Extrai os nomes das colunas
    header_line = lines[header_line_index]
    col_names = [header_line[col_starts[i]+1:col_starts[i+1]].strip() for i in range(len(col_starts)-1)]

    # Determina quantas linhas de dados ler
    if footer_line_index != -1:
        num_rows_to_read = footer_line_index - (separator_line_index + 1)
    else:
        num_rows_to_read = None # Lê até o final do arquivo

    df = pd.read_fwf(
        file_path,
        colspecs=[(col_starts[i], col_starts[i+1]) for i in range(len(col_starts)-1)],
        names=col_names,
        skiprows=separator_line_index + 1,
        nrows=num_rows_to_read,
        encoding='latin1'
    )

    # Remove a primeira linha de traços se ela for lida
    df = df[~df.iloc[:, 0].astype(str).str.strip().str.startswith('---')]

    # Remove colunas totalmente vazias, se houver
    df.dropna(axis=1, how='all', inplace=True)
    
    # Renomeia as colunas se houverem nomes repetidos
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols[cols == dup].index.values.tolist()] = [dup + '.' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
    df.columns = cols
    
    return df

def exportar_para_excel(pasta_dos_txts, caminho_excel_destino, *args, **kwargs):
    """
    Lê os arquivos TXT da pasta de destino, analisa suas estruturas
    individualmente e os salva em abas separadas do arquivo Excel.
    """
    try:
        with pd.ExcelWriter(caminho_excel_destino, engine='xlsxwriter') as writer:
            for key, sheet_name in config.SHEET_MAPPING.items():
                txt_filename = config.FILENAMES.get(key)
                if not txt_filename:
                    continue

                file_path = os.path.join(pasta_dos_txts, txt_filename)

                if not os.path.exists(file_path) or os.path.getsize(file_path) < 10:
                    print(f"AVISO: Arquivo '{txt_filename}' vazio ou não encontrado, pulando.")
                    continue

                print(f"Processando arquivo: '{txt_filename}'...")
                
                # Usa a nova função flexível para ler o arquivo
                df = analisar_e_ler_txt(file_path)

                if not df.empty:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f" -> Arquivo '{txt_filename}' salvo com sucesso na aba '{sheet_name}'.")
                else:
                    print(f" -> AVISO: Nenhum dado foi extraído de '{txt_filename}'. A aba '{sheet_name}' não será criada.")

        print(f"\nArquivo Excel '{caminho_excel_destino}' gerado com sucesso!")

    except Exception as e:
        print(f"ERRO CRÍTICO ao processar os arquivos e gerar o Excel: {e}")