# Passo 2 - Pegar o nome do arquivo que sera feito o update
import pandas as pd
import os
from arquivo import Arquivo

caminho_diretorio = "Y:/AGI/AGI/AGI PRONTOS HENRIQUE" # Diretorio que para pegar arquivos;
data_frame = {'Coluna1': [], 'Coluna2': []} # Criação do Data Frame de duas colunas
df = pd.DataFrame(data_frame) # Usando o objeto DataFrame da biblioteca pandas para colocar paramentros de um data frame;
df.columns = ['NOME FUNCIONARIO', 'TIPO DE ARQUIVO'] # Definindo os nomes das colunas 1 e 2 do Data Frame;

for arquivo in os.listdir(caminho_diretorio): # Acessando o diretorio usando a biblioteca os e lendo arquivo
    if arquivo.endswith('.pdf'):
        arquivoSuporte = Arquivo(arquivo)
        nome_func, tipo_arq = arquivoSuporte.fragmentName()
        df = df._append({'NOME FUNCIONARIO': nome_func, 'TIPO DE ARQUIVO': tipo_arq}, ignore_index=True)

df.to_excel('UpdateAGI.xlsx', index=False)