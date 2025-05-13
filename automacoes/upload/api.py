"""
LEIA!!!
Versão nova (03.04.2025).

atualizações: 
1- Renomea automaticamente os arquivos dentro das pastas GCPJ para o padrão aceito no sistema (*gcpj*_nome).
2- Resgata o caminho da origem dos GCPJs através do terminal.
"""

import os
import sys
import shutil
import string
import requests
from pathlib import Path
from win10toast import ToastNotifier

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))

# Configurações
TOKEN = "03de9d3089db8bbf4cc4d30740d3a2b6942a9467"
URL = "https://octopus.retake.com.br/lawsuits/api/attachment/link-lawsuit-by-gcpj/"

# Configurar cabeçalhos
headers = {
    "Authorization": f"Token {TOKEN}",
    "Accept": "application/json"
}

notificacao = ToastNotifier()

def renomear_subdir(pasta_anexos):
    # Convert to Path object for better path handling
    root_path = Path(pasta_anexos)
    
    #  Checa se o diretório existe
    if not root_path.exists():
        print(f"Error: Diretório '{pasta_anexos}' não existe.")
        return
    
    #  navegando pelos subdiretórios
    for subdir in root_path.iterdir():
        if subdir.is_dir():
            # Pegando o nome do subdiretório (GCPJ)
            subdir_name = subdir.name
            
            #  Processando todos os subdiretórios
            for file_path in subdir.glob('*'):
                if file_path.is_file():
                    # Resgata o nome riginal do subdiretório
                    original_name = file_path.name
                    
                    # Renomeando os diretórios com GCPJ no prefixo
                    translate_table = str.maketrans(' ','_',string.digits)
                    original_name = original_name.translate(translate_table)
                    new_name = f"{subdir_name}_{original_name}"
                    new_path = file_path.parent / new_name
                    
                    try:
                        # Renomenado os arquivos
                        file_path.replace(new_path)
                        print(f"Renamed: {original_name} -> {new_name}")
                    except Exception as e:
                        print(f"Error renaming {original_name}: {str(e)}")

def upload(pasta_anexos):
    print("Iniciando processo de upload...")

    for gcpj_pasta in os.listdir(pasta_anexos):
        caminho_pasta_gcpj = os.path.join(pasta_anexos, gcpj_pasta)

        if os.path.isdir(caminho_pasta_gcpj):
            print(f"\nProcessando pasta: {gcpj_pasta}")
            
            itens = os.listdir(caminho_pasta_gcpj)

            if not itens:
                #   Caso o diretório estiver vazio, ele será movido para a pasta VAZIOS
                print(f'Pasta {gcpj_pasta} está vazia')

                diretorios_vazios = os.path.join(pasta_anexos, 'VAZIOS')
                os.makedirs(diretorios_vazios, exist_ok=True)
                shutil.move(os.path.join(pasta_anexos, gcpj_pasta), diretorios_vazios)

            else:
                for arquivo in os.listdir(caminho_pasta_gcpj):
                    caminho_arquivo = os.path.join(caminho_pasta_gcpj, arquivo)
                    
                    try:
                        with open(caminho_arquivo, "rb") as f:
                            files = {
                                "document": (arquivo, f, "application/pdf")
                            }
                            response = requests.post(URL, files=files, headers=headers)
                        
                        print(f"\nEnviando {arquivo} para {gcpj_pasta}... Status: {response.status_code}")
                        
                        #   Caso o GCPJ não esteja no sistema, ele será movido para a pasta NÃO_ENCONTRADOS
                        if response.status_code == 400:
                            diretorios_nn_encotrados = os.path.join(pasta_anexos, 'NÃO_ENCONTRADOS')
                            os.makedirs(diretorios_nn_encotrados, exist_ok=True)
                            shutil.move(os.path.join(pasta_anexos, gcpj_pasta), diretorios_nn_encotrados)
                            break
                        
                        #   Caso ocorra um erro desconhecido, ele será movido para a pasta PARA_ANALISE
                        if response.status_code != 200 and response.status_code != 404:
                            print(f"Erro na resposta: {response.text}")
                            print(f"Cabeçalhos da resposta: {dict(response.headers)}")
                            para_analise = os.path.join(pasta_anexos, 'PARA_ANALISE')
                            os.makedirs(para_analise, exist_ok=True)
                            shutil.move(os.path.join(pasta_anexos, gcpj_pasta), para_analise)
                            break

                        else:
                            print(f"Upload realizado com sucesso para {arquivo}\n")
                        
                    except Exception as e:
                        print(f"\nErro ao processar arquivo {arquivo}: {str(e)}\n")
                        notificacao.show_toast("ERRO:", "Erro ao processar arquivo {arquivo}: {str(e)}", duration=2)
                # Move o diretório processado para o diretório "Anexos feitos" 
                # Cria o diretório "FEITOS" se não existir
                diretorio_feito = os.path.join(pasta_anexos, 'FEITOS')
                os.makedirs(diretorio_feito, exist_ok=True)
    print("\nTAREFA(S) CONCLUÍDA(S)!!")

def main(entrada):

    try:
        upload(entrada)
    except:
        print("Erro ao encontrar o Diretório")
        raise
    return