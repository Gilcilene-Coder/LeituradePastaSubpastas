import os
import pandas as pd
import time
import tkinter as tk
from tkinter import filedialog, messagebox

def listar_conteudo_detalhado(caminho_diretorio):
    data = []

    for root, _, files in os.walk(caminho_diretorio):
        rel_path = os.path.relpath(root, caminho_diretorio)
        split_path = rel_path.split(os.sep)

        for f in files:
            full_path = os.path.join(root, f)
            stats = os.stat(full_path)

            nome_arquivo = os.path.basename(f)
            extensao_arquivo = os.path.splitext(f)[1]
            data_acessada = time.strftime('%d/%m/%Y %H:%M', time.localtime(stats.st_atime))
            data_modificada = time.strftime('%d/%m/%Y %H:%M', time.localtime(stats.st_mtime))
            data_criada = time.strftime('%d/%m/%Y %H:%M', time.localtime(stats.st_ctime))

            entry = {
                'Pasta Raiz': os.path.basename(caminho_diretorio),
                'Extensão do Arquivo': extensao_arquivo,
                'Data acessada': data_acessada,
                'Data modificada': data_modificada,
                'Data Criada': data_criada
            }

            # Adiciona cada nível de subpasta como uma coluna e o arquivo correspondente
            for i, sp in enumerate(split_path, 1):
                entry[f'SubPasta_{i}'] = sp
                if i == len(split_path):
                    entry[f'Arquivos_{i}'] = nome_arquivo

            data.append(entry)

    return data

def exportar_para_excel(data, filename):
    df = pd.DataFrame(data)

    # Organizando a sequência das colunas
    subpasta_cols = [col for col in df.columns if 'SubPasta' in col]
    arquivo_cols = [col for col in df.columns if 'Arquivos' in col]
    columns_order = ['Pasta Raiz'] + subpasta_cols + arquivo_cols + ['Extensão do Arquivo', 'Data acessada', 'Data modificada', 'Data Criada']
    df = df[columns_order]

    df.to_excel(filename, sheet_name='Conteúdo Detalhado', index=False)

def escolher_diretorio_e_exportar():
    root = tk.Tk()
    root.withdraw()

    # Escolha do diretório
    caminho_diretorio = filedialog.askdirectory(title="Escolha a pasta de origem")
    if not caminho_diretorio:
        return  # O usuário cancelou a escolha do diretório

    data = listar_conteudo_detalhado(caminho_diretorio)

    # Escolha do local para salvar o arquivo
    caminho_arquivo = filedialog.asksaveasfilename(title="Salvar como", filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")), defaultextension=".xlsx")
    if not caminho_arquivo:
        return  # O usuário cancelou a escolha do local para salvar

    exportar_para_excel(data, caminho_arquivo)

    # Exibir mensagem de sucesso
    messagebox.showinfo("Concluído", "A exportação para Excel foi concluída com sucesso!")

# Execute a função
escolher_diretorio_e_exportar()
