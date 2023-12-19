import glob
import os
import win32api
import win32print
import time
import shutil

# Pasta onde os arquivos Word estão localizados
pasta = r"I:\ARQUIVO DIGITAL CPR\fabio jr\autodec\feitos"

# Pasta para onde os arquivos impressos serão movidos
pasta_impressos = r"i:\ARQUIVO DIGITAL CPR\fabio jr\autodec\feitos\impressos"

# Encontra todos os arquivos Word (.docx) na pasta
arquivos_word = glob.glob(os.path.join(pasta, '*.docx'))

# Verifica se há arquivos para imprimir
if not arquivos_word:
    print("Nenhum arquivo Word encontrado na pasta.")
else:
    # Configura a impressora padrão (você pode modificar para a impressora desejada)
    impressora_padrao = win32print.GetDefaultPrinter()

    # Loop pelos arquivos e imprime cada um deles com um atraso de 10 segundos entre as impressões
    for arquivo in arquivos_word:
        try:
            win32api.ShellExecute(
                0,  # hwnd
                'print',
                arquivo,
                f'"{impressora_padrao}"',  # nome da impressora
                '.',  # diretório de trabalho
                0  # ação (0 para imprimir)
            )
            print(f"Arquivo '{arquivo}' enviado para impressão.")
            
            # Aguarde 10 segundos antes de imprimir o próximo arquivo
            time.sleep(5)
            
            # Move o arquivo impresso para a pasta de impressos
            nome_arquivo = os.path.basename(arquivo)
            destino = os.path.join(pasta_impressos, nome_arquivo)
            shutil.move(arquivo, destino)
            print(f"Arquivo '{arquivo}' movido para '{destino}' após a impressão.")
        except Exception as e:
            print(f"Erro ao imprimir '{arquivo}': {str(e)}")
            # Adicione um tratamento de erro mais específico para mover o arquivo
            try:
                shutil.move(arquivo, destino)
            except Exception as move_error:
                print(f"Erro ao mover '{arquivo}' para '{destino}': {str(move_error)}")
