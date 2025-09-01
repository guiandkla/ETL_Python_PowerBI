import pandas as pd
import os
import glob

# Caminho para leitura dos arquivos
folder_path = 'src\\data\\raw'

# Lista todos os arquivos de Excel
excel_files = glob.glob(os.path.join(folder_path, 'base_devoluções.xlsx'))

if not excel_files:
    print("Não foi possível encontrar um arquivo Excel (.xlsx).")
else:

    # Cria um data frame (uma tabela em memória para guardar o conteúdo dos arquivos):

    dfs = []

    for excel_file in excel_files:

        try: # Lê todo o conteúdo dos arquivos Excel:
            df_temp = pd.read_excel(excel_file)

            # Converte a coluna "data" para datetime (caso ainda não esteja):
            df_temp["Data"] = pd.to_datetime(df_temp["Data"], errors="coerce")

            # Formata a data para o padrão Brasileiro (dia/mês/ano)
            df_temp["Data"] = df_temp["Data"].dt.strftime("%d/%m/%Y")

            
            # Guarda os dados tratados dentro de um dataframe comum:
            dfs.append(df_temp)

            # Informa erro na leitura:
        except Exception as e:
            print(f"Erro ao ler o arquivo {excel_file} : {e}")


    if dfs:

        # Concatena todas as tabelas salvas DFS como uma tebela única:
        result = pd.concat(dfs, ignore_index=True)

        # Informa o cominho de saída:
        output_file = os.path.join('src', 'data', 'ready', 'base_devoluções_ready.xlsx')

        # Configuração do motor de escrita:
        writer = pd.ExcelWriter(output_file, engine='openpyxl')

        # Leva os dados do resultado a serem escritos no motor de Exel configurado:
        result.to_excel(writer, index=False)

        writer._save()
    else:
        print("Nenhum dado para ser salvo!")