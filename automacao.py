"""
PASSO A PASSO PARA FAZER O UPDATE NOS ARQUIVOS AGI

Passo 1 - Acessar o sistema
Passo 2 - Pegar o nome do arquivo que sera feito o update
Passo 3 - Pesquisar no sistema o nome do arquivo
Passo 4 - Realizar o delete
Passo 5 - Acessar a aba para subir os arquivos que foram deletado


PONTO CHAVES
- Nome dos arquivos teram que ser feito uma analise e correcao
    - Passar todos os nomes para uma planilha
    - Pegar definir quais sao os nomes padroes de arquivos
    - Fazer testes para tipos de arquivos

ESCRITA DE TIPO DE ARQUIVO
    - ESCRITO NO ARQUIVO | TIPO DE ARQUIVO
    - SAUDE E SEGURANCA  | DOCUMENTOS SAUDE E SEGURANCA DO TRABALHO
    - MENSAL             | DOCUMENTOS MENSAIS
    - ADMISSIONAL        | DOCUMENTOS ADMISSIONAIS
    - PESSOAIS           | DOCUMENTOS PESSOAIS
    - DEMISSIONAL        | DOCUMENTOS DEMISSIONAIS
    - TREINAMENTOS E CERTIFICADOS | TREINAMENTOS E CERTIFICADOS

    CASO FUNCIONARIO TENHAS MAIS DE UMA MATRICULA

    EXEMPLO:
        - NOME FUNCIONARIO- TIPO DO ARQUIVO - TIPO DO DOCUMENTO.PDF

CASO DE GERENTES
    ESCRITA PARA ARQUIVOS
        - GERENTE NOME DO GERENTE - TIPO DO ARQUIVO - TIPO DE DOCUMENTO.PDF

    SISTEMA PEGA E FAZ COM QUE SEJA SELECIONADO O DEPARTAMENTO DE GERENCIA E APOS ISSO ELE EXCLUI A PARTE DO GERENTES
    FAZENDO A PESQUISA SOMENTO DO NOME DO GERENTE

NOMES COM PROBLEMAS
    - HUGO HENRIQUE LOPES DA SILVA - 70002 - 3233
    - JIDEONE SANTOS DA SILVA - JIDEONE SANTOS DA MOTA SEGGER

UPDATE PARA ANALISAR
    - LEITURA DE PDFS
    - AGILIDADE MAIOR COM O TEMPO DE UPLOAD
    - INCLUSÃO DE ARQUIVOS DIRETO NO ARQUIVO PRINCIPAL
    - ANALISE DE DADOS SOBRE O NOME DO FUNCIONARIO E O TIPO DE ARQUIVO, AONDE SE CASO O
 FUNCIONARIO REPITA O MESMO DOCUMENTO ELE NAO REPITA NA PLANILHA. (PROX ETAPA)

ANALISE FERNANDO
    - FERNANDO PASSOU QUE ACHA IMPORTANTE UTILIZARMOS O CTRL + C E O CTRL + V PARA ESCREVER O NOME DOS FUNCIONARIOS
PEGANDO O NOME DIRETAMENTE DO SERVIDO DA AGI
     - PADRONIZAR A ESCRITA DOS TIPOS DE ARQUIVO.
"""

import pandas as pd
import time
import pyautogui as pyag
import pygetwindow as gw
import os
from arquivo import Arquivo

# Passo 1 - Acessar o sistema

sistema = "https://sistema.qdoc.com.br/"  # Site do sistema;
pyag.PAUSE = 2  # Tempo de espera para realizar cada passo da automação;

pyag.press('win')  # Pressionando botão win do teclado;
pyag.write('chrome')  # Escrevendo chrome na pesquisa para abrir o google chrome;
pyag.press('enter')  # Pressionando a trecla enter do teclado;

time.sleep(3)  # Tempo de espera para carregar o chrome;

name_pag = 'Google Chrome'  # Nome da janela que deseja encontrar;

janela = gw.getWindowsWithTitle(name_pag)  # Utilizando a função getWindowsWithTitle da biblioteca pygetwindow para obter uma lista de janelas que correspodem ao nome fornecido;

if janela: # Verificando se a lista nao esta vazia;
    janela[0].maximize()  # Se uma janela for encontrada [0 (janela encontrada)], usa a função maximize para colocar a janela em fullscreem;

pyag.write(sistema)  # Escrevendo a variavel sistema na barra de pesquisa do chrome;
pyag.press("enter")  # Pressionando a trecla enter do teclado;

time.sleep(20)  # Tempo de espera para carregar o sistema

# Passo 3 - Pesquisar no sistema o nome do arquivo

# PAGINA DE LISTAR
pyag.click(x=356, y=215)
pyag.click(x=353, y=306)
pyag.click(x=223, y=390)
pyag.write('AGI')
pyag.press('enter')

# PAGINA DE CADASTRAR
pyag.click(x=356, y=215)
pyag.rightClick(x=362, y=271)
pyag.click(x=402, y=295)
pyag.click(x=383, y=15)
pyag.click(x=292, y=418)
pyag.write('AGI')
pyag.press('enter')

# LEITURA DA PLANILHA
planilha = pd.read_excel('UpdateAGI.xlsx')

for linha in planilha.index:
    # LEITURA DAS COLUNAS DA PLANILHA
    nome_funcionario = str(planilha.loc[linha, 'NOME FUNCIONARIO'])
    tipo_arquivo = str(planilha.loc[linha, 'TIPO DE ARQUIVO'])

    # PAGINA DE LISTAR
    pyag.click(x=172, y=9)

    if 'GERENTE' in nome_funcionario:
        nome_funcionario = nome_funcionario.replace('GERENTE',
                                                    '').strip()  # FAZENDO COM QUE APAGUE A PALAVRA GERENTE SE CASO HOUVER NO NOME
        pyag.click(x=785, y=383)  # CONTEUDO
        pyag.write('GERENCIA')  # ESCRITA CONTEUDO

        if tipo_arquivo == 'SAUDE E SEGURANCA':
            pyag.click(x=1212, y=387)  # DESCRICAO DE DOCUMENTO
            pyag.hotkey("ctrl", "a")  # SELECIONANDO TUDO ESCRITO NA DESCRICAO DE DOCUMENTOS
            pyag.press("backspace")  # APAGANDO OQUE ESTAVA ESCRITO
            nome_pesquisa = f'{nome_funcionario} - DOCUMENTOS DE SAUDE E SEGURANCA DO TRABALHO'  # NOME DE FUNCIONARIO E O TIPO DO ARQUIVO (TABELA NA DOCSTRING)
            pyag.write(nome_pesquisa.upper())  # ESCREVENDO NOME DO FUNCIONARIO TUDO EM MAIUSCULO
            pyag.press("enter")  # CONFIRMAR PESQUISA
            time.sleep(3)  # PAUSA CARREGAMENTO

            # DELETE DO ARQUIVO ANTIGO
            pyag.click(x=1811, y=652)  # CLICANDO EM FUNCOES NO ARQUIVO
            pyag.click(x=1772, y=757)  # CLICANDO EM APAGAR
            pyag.click(x=1031, y=185)  # CONFIRMANDO DELETE

            # UPLOAD ARQUIVO NOVO
            pyag.click(x=383, y=15)  # SELECIONANDO ABA PARA UPLOAD
            pyag.click(x=797, y=411)  # SELECIONANDO CONTEUDO
            pyag.write('GERENCIA')  # CONTEUDO DESEJADO
            pyag.click(x=954, y=523)  # CLICANDO NA TELA PARA ABRIR A SELECAO DE ARQUIVOS

            # ABA DE ESCOLHER ARQUIVOS
            pyag.click(x=78, y=133)  # (TELA DO EXPLORADOR DE ARQUIVOS) APERTANDO NO INICIO
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('Clientes')  # DIRETORIO PESQUISADO
            pyag.doubleClick(x=843, y=146)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('AGI')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('PRONTUARIOS')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('- GERENTES')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA

            pyag.click(x=770, y=966)  # BARRA DE PESQUISA
            pyag.write(nome_funcionario)  # NOME DO FUNCIONARIO PARA PROCURAR SUB-DIRETORIO
            pyag.press('enter')  # CONFIRMANDO SUB-DIRETORIO
            pyag.write(f"{nome_pesquisa}.pdf")  # ESCOLHENDO ARQUIVO QUE DESEJA FAZER O UPLOAD
            pyag.press('enter')  # CONFIRMANDO ARQUIVO

            time.sleep(15)  # TEMPO DE CARREGAMENTO

            pyag.click(x=172, y=9)  # ABA DE LISTAR
            pyag.hotkey("fn", "f5")  # ATUALIZANDO ABA

            time.sleep(2)  # TEMPO DE ESPERA

        elif tipo_arquivo == 'MENSAL' or tipo_arquivo == 'MENSAIS':
            pyag.click(x=1212, y=387)  # DESCRICAO DE DOCUMENTO
            pyag.hotkey("ctrl", "a")  # SELECIONANDO TUDO ESCRITO NA DESCRICAO DE DOCUMENTOS
            pyag.press("backspace")  # APAGANDO OQUE ESTAVA ESCRITO
            nome_pesquisa = f'{nome_funcionario} - DOCUMENTOS MENSAIS'  # NOME DE FUNCIONARIO E O TIPO DO ARQUIVO (TABELA NA DOCSTRING)
            pyag.write(nome_pesquisa.upper())  # ESCREVENDO NOME DO FUNCIONARIO TUDO EM MAIUSCULO
            pyag.press("enter")  # CONFIRMAR PESQUISA
            time.sleep(3)  # PAUSA CARREGAMENTO

            # DELETE DO ARQUIVO ANTIGO
            pyag.click(x=1811, y=652)  # CLICANDO EM FUNCOES NO ARQUIVO
            pyag.click(x=1772, y=757)  # CLICANDO EM APAGAR
            pyag.click(x=1031, y=185)  # CONFIRMANDO DELETE

            # UPLOAD ARQUIVO NOVO
            pyag.click(x=383, y=15)  # SELECIONANDO ABA PARA UPLOAD
            pyag.click(x=797, y=411)  # SELECIONANDO CONTEUDO
            pyag.write('GERENCIA')  # CONTEUDO DESEJADO
            pyag.click(x=954, y=523)  # CLICANDO NA TELA PARA ABRIR A SELECAO DE ARQUIVOS

            # ABA DE ESCOLHER ARQUIVOS
            pyag.click(x=78, y=133)  # (TELA DO EXPLORADOR DE ARQUIVOS) APERTANDO NO INICIO
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('Clientes')  # DIRETORIO PESQUISADO
            pyag.doubleClick(x=843, y=146)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('AGI')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('PRONTUARIOS')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
            pyag.doubleClick(x=362, y=219)

            pyag.click(x=770, y=966)  # BARRA DE PESQUISA
            pyag.write(nome_funcionario)  # NOME DO FUNCIONARIO PARA PROCURAR SUB-DIRETORIO
            pyag.press('enter')  # CONFIRMANDO SUB-DIRETORIO
            pyag.write(f"{nome_pesquisa}.pdf")  # ESCOLHENDO ARQUIVO QUE DESEJA FAZER O UPLOAD
            pyag.press('enter')  # CONFIRMANDO ARQUIVO

            time.sleep(15)  # TEMPO DE CARREGAMENTO

            pyag.click(x=172, y=9)  # ABA DE LISTAR
            pyag.hotkey("fn", "f5")  # ATUALIZANDO ABA

            time.sleep(2)  # TEMPO DE ESPERA

        elif tipo_arquivo == 'ADMISSIONAL' or tipo_arquivo == 'DOCUMENTOS ADMISSIONAIS':
            pyag.click(x=1212, y=387)  # DESCRICAO DE DOCUMENTO
            pyag.hotkey("ctrl", "a")  # SELECIONANDO TUDO ESCRITO NA DESCRICAO DE DOCUMENTOS
            pyag.press("backspace")  # APAGANDO OQUE ESTAVA ESCRITO
            nome_pesquisa = f'{nome_funcionario} - DOCUMENTOS ADMISSIONAIS'  # NOME DE FUNCIONARIO E O TIPO DO ARQUIVO (TABELA NA DOCSTRING)
            pyag.write(nome_pesquisa.upper())  # ESCREVENDO NOME DO FUNCIONARIO TUDO EM MAIUSCULO
            pyag.press("enter")  # CONFIRMAR PESQUISA
            time.sleep(3)  # PAUSA CARREGAMENTO

            # DELETE DO ARQUIVO ANTIGO
            pyag.click(x=1811, y=652)  # CLICANDO EM FUNCOES NO ARQUIVO
            pyag.click(x=1772, y=757)  # CLICANDO EM APAGAR
            pyag.click(x=1031, y=185)  # CONFIRMANDO DELETE

            # UPLOAD ARQUIVO NOVO
            pyag.click(x=383, y=15)  # SELECIONANDO ABA PARA UPLOAD
            pyag.click(x=797, y=411)  # SELECIONANDO CONTEUDO
            pyag.write('GERENCIA')  # CONTEUDO DESEJADO
            pyag.click(x=954, y=523)  # CLICANDO NA TELA PARA ABRIR A SELECAO DE ARQUIVOS

            # ABA DE ESCOLHER ARQUIVOS
            pyag.click(x=78, y=133)  # (TELA DO EXPLORADOR DE ARQUIVOS) APERTANDO NO INICIO
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('Clientes')  # DIRETORIO PESQUISADO
            pyag.doubleClick(x=843, y=146)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('AGI')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('PRONTUARIOS')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('- GERENTES')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA

            pyag.click(x=770, y=966)  # BARRA DE PESQUISA
            pyag.write(nome_funcionario)  # NOME DO FUNCIONARIO PARA PROCURAR SUB-DIRETORIO
            pyag.press('enter')  # CONFIRMANDO SUB-DIRETORIO
            pyag.write(f"{nome_pesquisa}.pdf")  # ESCOLHENDO ARQUIVO QUE DESEJA FAZER O UPLOAD
            pyag.press('enter')  # CONFIRMANDO ARQUIVO

            time.sleep(15)  # TEMPO DE CARREGAMENTO

            pyag.click(x=172, y=9)  # ABA DE LISTAR
            pyag.hotkey("fn", "f5")  # ATUALIZANDO ABA

            time.sleep(2)  # TEMPO DE ESPERA

        elif tipo_arquivo == 'PESSOAIS' or tipo_arquivo == 'DOCUMENTOS PESSOAIS':
            pyag.click(x=1212, y=387)  # DESCRICAO DE DOCUMENTO
            pyag.hotkey("ctrl", "a")  # SELECIONANDO TUDO ESCRITO NA DESCRICAO DE DOCUMENTOS
            pyag.press("backspace")  # APAGANDO OQUE ESTAVA ESCRITO
            nome_pesquisa = f'{nome_funcionario} - DOCUMENTOS PESSOAIS'  # NOME DE FUNCIONARIO E O TIPO DO ARQUIVO (TABELA NA DOCSTRING)
            pyag.write(nome_pesquisa.upper())  # ESCREVENDO NOME DO FUNCIONARIO TUDO EM MAIUSCULO
            pyag.press("enter")  # CONFIRMAR PESQUISA
            time.sleep(3)  # PAUSA CARREGAMENTO

            # DELETE DO ARQUIVO ANTIGO
            pyag.click(x=1811, y=652)  # CLICANDO EM FUNCOES NO ARQUIVO
            pyag.click(x=1772, y=757)  # CLICANDO EM APAGAR
            pyag.click(x=1031, y=185)  # CONFIRMANDO DELETE

            # UPLOAD ARQUIVO NOVO
            pyag.click(x=383, y=15)  # SELECIONANDO ABA PARA UPLOAD
            pyag.click(x=797, y=411)  # SELECIONANDO CONTEUDO
            pyag.write('GERENCIA')  # CONTEUDO DESEJADO
            pyag.click(x=954, y=523)  # CLICANDO NA TELA PARA ABRIR A SELECAO DE ARQUIVOS

            # ABA DE ESCOLHER ARQUIVOS
            pyag.click(x=78, y=133)  # (TELA DO EXPLORADOR DE ARQUIVOS) APERTANDO NO INICIO
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('Clientes')  # DIRETORIO PESQUISADO
            pyag.doubleClick(x=843, y=146)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('AGI')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('PRONTUARIOS')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('- GERENTES')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA

            pyag.click(x=770, y=966)  # BARRA DE PESQUISA
            pyag.write(nome_funcionario)  # NOME DO FUNCIONARIO PARA PROCURAR SUB-DIRETORIO
            pyag.press('enter')  # CONFIRMANDO SUB-DIRETORIO
            pyag.write(f"{nome_pesquisa}.pdf")  # ESCOLHENDO ARQUIVO QUE DESEJA FAZER O UPLOAD
            pyag.press('enter')  # CONFIRMANDO ARQUIVO

            time.sleep(15)  # TEMPO DE CARREGAMENTO

            pyag.click(x=172, y=9)  # ABA DE LISTAR
            pyag.hotkey("fn", "f5")  # ATUALIZANDO ABA

            time.sleep(2)  # TEMPO DE ESPERA

        elif tipo_arquivo == 'DEMISSIONAL'or tipo_arquivo == 'DOCUMENTOS DEMISSIONAIS':
            pyag.click(x=1212, y=387)  # DESCRICAO DE DOCUMENTO
            pyag.hotkey("ctrl", "a")  # SELECIONANDO TUDO ESCRITO NA DESCRICAO DE DOCUMENTOS
            pyag.press("backspace")  # APAGANDO OQUE ESTAVA ESCRITO
            nome_pesquisa = f'{nome_funcionario} - DOCUMENTOS DEMISSIONAL'  # NOME DE FUNCIONARIO E O TIPO DO ARQUIVO (TABELA NA DOCSTRING)
            pyag.write(nome_pesquisa.upper())  # ESCREVENDO NOME DO FUNCIONARIO TUDO EM MAIUSCULO
            pyag.press("enter")  # CONFIRMAR PESQUISA
            time.sleep(3)  # PAUSA CARREGAMENTO

            # DELETE DO ARQUIVO ANTIGO
            pyag.click(x=1811, y=652)  # CLICANDO EM FUNCOES NO ARQUIVO
            pyag.click(x=1772, y=757)  # CLICANDO EM APAGAR
            pyag.click(x=1031, y=185)  # CONFIRMANDO DELETE

            # UPLOAD ARQUIVO NOVO
            pyag.click(x=383, y=15)  # SELECIONANDO ABA PARA UPLOAD
            pyag.click(x=797, y=411)  # SELECIONANDO CONTEUDO
            pyag.write('GERENCIA')  # CONTEUDO DESEJADO
            pyag.click(x=954, y=523)  # CLICANDO NA TELA PARA ABRIR A SELECAO DE ARQUIVOS

            # ABA DE ESCOLHER ARQUIVOS
            pyag.click(x=78, y=133)  # (TELA DO EXPLORADOR DE ARQUIVOS) APERTANDO NO INICIO
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('Clientes')  # DIRETORIO PESQUISADO
            pyag.doubleClick(x=843, y=146)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('AGI')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('PRONTUARIOS')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('- GERENTES')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA

            pyag.click(x=770, y=966)  # BARRA DE PESQUISA
            pyag.write(nome_funcionario)  # NOME DO FUNCIONARIO PARA PROCURAR SUB-DIRETORIO
            pyag.press('enter')  # CONFIRMANDO SUB-DIRETORIO
            pyag.write(f"{nome_pesquisa}.pdf")  # ESCOLHENDO ARQUIVO QUE DESEJA FAZER O UPLOAD
            pyag.press('Page Down')
            pyag.press('enter')  # CONFIRMANDO ARQUIVO

            time.sleep(15)  # TEMPO DE CARREGAMENTO

            pyag.click(x=172, y=9)  # ABA DE LISTAR
            pyag.hotkey("fn", "f5")  # ATUALIZANDO ABA

            time.sleep(2)  # TEMPO DE ESPERA

        elif tipo_arquivo == 'TREINAMENTOS E CERTIFICADOS' or tipo_arquivo == 'TREINAMENTO E CERTIFICADO' \
                or tipo_arquivo == 'TREINAMENTO E CERTIFICADOS' or tipo_arquivo == 'TREINAMENTOS E CERTIFICADO':

            pyag.click(x=1212, y=387)  # DESCRICAO DE DOCUMENTO
            pyag.hotkey("ctrl", "a")  # SELECIONANDO TUDO ESCRITO NA DESCRICAO DE DOCUMENTOS
            pyag.press("backspace")  # APAGANDO OQUE ESTAVA ESCRITO
            nome_pesquisa = f'{nome_funcionario} - TREINAMENTOS E CERTIFICADOS'  # NOME DE FUNCIONARIO E O TIPO DO ARQUIVO (TABELA NA DOCSTRING)
            pyag.write(nome_pesquisa.upper())  # ESCREVENDO NOME DO FUNCIONARIO TUDO EM MAIUSCULO
            pyag.press("enter")  # CONFIRMAR PESQUISA
            time.sleep(3)  # PAUSA CARREGAMENTO

            # DELETE DO ARQUIVO ANTIGO
            pyag.click(x=1811, y=652)  # CLICANDO EM FUNCOES NO ARQUIVO
            pyag.click(x=1772, y=757)  # CLICANDO EM APAGAR
            pyag.click(x=1031, y=185)  # CONFIRMANDO DELETE

            # UPLOAD ARQUIVO NOVO
            pyag.click(x=383, y=15)  # SELECIONANDO ABA PARA UPLOAD
            pyag.click(x=797, y=411)  # SELECIONANDO CONTEUDO
            pyag.write('GERENCIA')  # CONTEUDO DESEJADO
            pyag.click(x=954, y=523)  # CLICANDO NA TELA PARA ABRIR A SELECAO DE ARQUIVOS

            # ABA DE ESCOLHER ARQUIVOS
            pyag.click(x=78, y=133)  # (TELA DO EXPLORADOR DE ARQUIVOS) APERTANDO NO INICIO
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('Clientes')  # DIRETORIO PESQUISADO
            pyag.doubleClick(x=843, y=146)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('AGI')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('PRONTUARIOS')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('- GERENTES')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA

            pyag.click(x=770, y=966)  # BARRA DE PESQUISA
            pyag.write(nome_funcionario)  # NOME DO FUNCIONARIO PARA PROCURAR SUB-DIRETORIO
            pyag.press('enter')  # CONFIRMANDO SUB-DIRETORIO
            pyag.write(f"{nome_pesquisa}.pdf")  # ESCOLHENDO ARQUIVO QUE DESEJA FAZER O UPLOAD
            pyag.press('enter')  # CONFIRMANDO ARQUIVO

            time.sleep(15)  # TEMPO DE CARREGAMENTO

            pyag.click(x=172, y=9)  # ABA DE LISTAR
            pyag.hotkey("fn", "f5")  # ATUALIZANDO ABA

            time.sleep(2)  # TEMPO DE ESPERA

    else:

        if tipo_arquivo == 'SAUDE E SEGURANCA':
            pyag.click(x=785, y=383)  # CONTEUDO
            pyag.write('SEGURANÇA DO TRABALHO')  # ESCRITA CONTEUDO

            pyag.click(x=1212, y=387)  # DESCRICAO DE DOCUMENTO
            pyag.hotkey("ctrl", "a")  # SELECIONANDO TUDO ESCRITO NA DESCRICAO DE DOCUMENTOS
            pyag.press("backspace")  # APAGANDO OQUE ESTAVA ESCRITO
            nome_pesquisa = f'{nome_funcionario} - DOCUMENTOS DE SAUDE E SEGURANCA DO TRABALHO'  # NOME DE FUNCIONARIO E O TIPO DO ARQUIVO (TABELA NA DOCSTRING)
            pyag.write(nome_pesquisa.upper())  # ESCREVENDO NOME DO FUNCIONARIO TUDO EM MAIUSCULO
            pyag.press("enter")  # CONFIRMAR PESQUISA
            time.sleep(3)  # PAUSA CARREGAMENTO

            # DELETE DO ARQUIVO ANTIGO
            pyag.click(x=1811, y=652)  # CLICANDO EM FUNCOES NO ARQUIVO
            pyag.click(x=1772, y=757)  # CLICANDO EM APAGAR
            pyag.click(x=1031, y=185)  # CONFIRMANDO DELETE

            # UPLOAD ARQUIVO NOVO
            pyag.click(x=383, y=15)  # SELECIONANDO ABA PARA UPLOAD
            pyag.click(x=797, y=411)  # SELECIONANDO CONTEUDO
            pyag.write('SEGURAÇA DO TRABALHO')  # CONTEUDO DESEJADO
            pyag.click(x=954, y=523)  # CLICANDO NA TELA PARA ABRIR A SELECAO DE ARQUIVOS

            # ABA DE ESCOLHER ARQUIVOS
            pyag.click(x=78, y=133)  # (TELA DO EXPLORADOR DE ARQUIVOS) APERTANDO NO INICIO
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('Clientes')  # DIRETORIO PESQUISADO
            pyag.doubleClick(x=843, y=146)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('AGI')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
            pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
            pyag.write('PRONTUARIOS')  # SUB-DIRETORIO PESQUISADO
            pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA

            pyag.click(x=770, y=966)  # BARRA DE PESQUISA
            pyag.write(nome_funcionario)  # NOME DO FUNCIONARIO PARA PROCURAR SUB-DIRETORIO
            pyag.press('enter')  # CONFIRMANDO SUB-DIRETORIO
            pyag.write(f"{nome_pesquisa}.pdf")  # ESCOLHENDO ARQUIVO QUE DESEJA FAZER O UPLOAD
            pyag.press('enter')  # CONFIRMANDO ARQUIVO

            time.sleep(15)  # TEMPO DE CARREGAMENTO

            pyag.click(x=172, y=9)  # ABA DE LISTAR
            pyag.hotkey("fn", "f5")  # ATUALIZANDO ABA

            time.sleep(2)  # TEMPO DE ESPERA

        else:
            pyag.click(x=785, y=383)  # CONTEUDO
            pyag.write('RECURSOS HUMANOS')  # ESCRITA CONTEUDO

            pyag.click(x=1212, y=387)  # DESCRICAO DE DOCUMENTO
            pyag.hotkey("ctrl", "a")  # SELECIONANDO TUDO ESCRITO NA DESCRICAO DE DOCUMENTOS
            pyag.press("backspace")  # APAGANDO OQUE ESTAVA ESCRITO

            if tipo_arquivo == 'MENSAL' or tipo_arquivo == 'MENSAIS' or tipo_arquivo == 'DOCUMENTOS MENSAIS':

                nome_pesquisa = f'{nome_funcionario} - DOCUMENTOS MENSAIS'  # NOME DE FUNCIONARIO E O TIPO DO ARQUIVO (TABELA NA DOCSTRING)
                pyag.write(nome_pesquisa.upper())  # ESCREVENDO NOME DO FUNCIONARIO TUDO EM MAIUSCULO
                pyag.press("enter")  # CONFIRMAR PESQUISA
                time.sleep(3)  # PAUSA CARREGAMENTO

                # DELETE DO ARQUIVO ANTIGO
                pyag.click(x=1811, y=652)  # CLICANDO EM FUNCOES NO ARQUIVO
                pyag.click(x=1772, y=757)  # CLICANDO EM APAGAR
                pyag.click(x=1031, y=185)  # CONFIRMANDO DELETE

                # UPLOAD ARQUIVO NOVO
                pyag.click(x=383, y=15)  # SELECIONANDO ABA PARA UPLOAD
                pyag.click(x=797, y=411)  # SELECIONANDO CONTEUDO
                pyag.write('RECURSOS HUMANOS')  # CONTEUDO DESEJADO
                pyag.click(x=954, y=523)  # CLICANDO NA TELA PARA ABRIR A SELECAO DE ARQUIVOS

                # ABA DE ESCOLHER ARQUIVOS
                pyag.click(x=78, y=133)  # (TELA DO EXPLORADOR DE ARQUIVOS) APERTANDO NO INICIO
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('Clientes')  # DIRETORIO PESQUISADO
                pyag.doubleClick(x=843, y=146)  # SELECIONANDO PESQUISA
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('AGI')  # SUB-DIRETORIO PESQUISADO
                pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('PRONTUARIOS')  # SUB-DIRETORIO PESQUISADO
                pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA

                pyag.click(x=770, y=966)  # BARRA DE PESQUISA
                pyag.write(nome_funcionario)  # NOME DO FUNCIONARIO PARA PROCURAR SUB-DIRETORIO
                pyag.press('enter')  # CONFIRMANDO SUB-DIRETORIO
                pyag.write(f"{nome_pesquisa}.pdf")  # ESCOLHENDO ARQUIVO QUE DESEJA FAZER O UPLOAD
                pyag.press('enter')  # CONFIRMANDO ARQUIVO

                time.sleep(15)  # TEMPO DE CARREGAMENTO

                pyag.click(x=172, y=9)  # ABA DE LISTAR
                pyag.hotkey("fn", "f5")  # ATUALIZANDO ABA

                time.sleep(2)  # TEMPO DE ESPERA

            elif tipo_arquivo == 'ADMISSIONAL' or tipo_arquivo == 'DOCUMENTOS ADMISSIONAIS':

                nome_pesquisa = f'{nome_funcionario} - DOCUMENTOS ADMISSIONAIS'  # NOME DE FUNCIONARIO E O TIPO DO ARQUIVO (TABELA NA DOCSTRING)
                pyag.write(nome_pesquisa.upper())  # ESCREVENDO NOME DO FUNCIONARIO TUDO EM MAIUSCULO
                pyag.press("enter")  # CONFIRMAR PESQUISA
                time.sleep(3)  # PAUSA CARREGAMENTO

                # DELETE DO ARQUIVO ANTIGO
                pyag.click(x=1811, y=652)  # CLICANDO EM FUNCOES NO ARQUIVO
                pyag.click(x=1772, y=757)  # CLICANDO EM APAGAR
                pyag.click(x=1031, y=185)  # CONFIRMANDO DELETE

                # UPLOAD ARQUIVO NOVO
                pyag.click(x=383, y=15)  # SELECIONANDO ABA PARA UPLOAD
                pyag.click(x=797, y=411)  # SELECIONANDO CONTEUDO
                pyag.write('RECURSOS HUMANOS')  # CONTEUDO DESEJADO
                pyag.click(x=954, y=523)  # CLICANDO NA TELA PARA ABRIR A SELECAO DE ARQUIVOS

                # ABA DE ESCOLHER ARQUIVOS
                pyag.click(x=78, y=133)  # (TELA DO EXPLORADOR DE ARQUIVOS) APERTANDO NO INICIO
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('Clientes')  # DIRETORIO PESQUISADO
                pyag.doubleClick(x=843, y=146)  # SELECIONANDO PESQUISA
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('AGI')  # SUB-DIRETORIO PESQUISADO
                pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('PRONTUARIOS')  # SUB-DIRETORIO PESQUISADO
                pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA

                pyag.click(x=770, y=966)  # BARRA DE PESQUISA
                pyag.write(nome_funcionario)  # NOME DO FUNCIONARIO PARA PROCURAR SUB-DIRETORIO
                pyag.press('enter')  # CONFIRMANDO SUB-DIRETORIO
                pyag.write(f"{nome_pesquisa}.pdf")  # ESCOLHENDO ARQUIVO QUE DESEJA FAZER O UPLOAD
                pyag.press('enter')  # CONFIRMANDO ARQUIVO

                time.sleep(15)  # TEMPO DE CARREGAMENTO

                pyag.click(x=172, y=9)  # ABA DE LISTAR
                pyag.hotkey("fn", "f5")  # ATUALIZANDO ABA

                time.sleep(2)  # TEMPO DE ESPERA

            elif tipo_arquivo == 'PESSOAIS' or tipo_arquivo == 'DOCUMENTOS PESSOAIS':

                nome_pesquisa = f'{nome_funcionario} - DOCUMENTOS PESSOAIS'  # NOME DE FUNCIONARIO E O TIPO DO ARQUIVO (TABELA NA DOCSTRING)
                pyag.write(nome_pesquisa.upper())  # ESCREVENDO NOME DO FUNCIONARIO TUDO EM MAIUSCULO
                pyag.press("enter")  # CONFIRMAR PESQUISA
                time.sleep(3)  # PAUSA CARREGAMENTO

                # DELETE DO ARQUIVO ANTIGO
                pyag.click(x=1811, y=652)  # CLICANDO EM FUNCOES NO ARQUIVO
                pyag.click(x=1772, y=757)  # CLICANDO EM APAGAR
                pyag.click(x=1031, y=185)  # CONFIRMANDO DELETE

                # UPLOAD ARQUIVO NOVO
                pyag.click(x=383, y=15)  # SELECIONANDO ABA PARA UPLOAD
                pyag.click(x=797, y=411)  # SELECIONANDO CONTEUDO
                pyag.write('RECURSOS HUMANOS')  # CONTEUDO DESEJADO
                pyag.click(x=954, y=523)  # CLICANDO NA TELA PARA ABRIR A SELECAO DE ARQUIVOS

                # ABA DE ESCOLHER ARQUIVOS
                pyag.click(x=78, y=133)  # (TELA DO EXPLORADOR DE ARQUIVOS) APERTANDO NO INICIO
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('Clientes')  # DIRETORIO PESQUISADO
                pyag.doubleClick(x=843, y=146)  # SELECIONANDO PESQUISA
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('AGI')  # SUB-DIRETORIO PESQUISADO
                pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('PRONTUARIOS')  # SUB-DIRETORIO PESQUISADO
                pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA

                pyag.click(x=770, y=966)  # BARRA DE PESQUISA
                pyag.write(nome_funcionario)  # NOME DO FUNCIONARIO PARA PROCURAR SUB-DIRETORIO
                pyag.press('enter')  # CONFIRMANDO SUB-DIRETORIO
                pyag.write(f"{nome_pesquisa}.pdf")  # ESCOLHENDO ARQUIVO QUE DESEJA FAZER O UPLOAD
                pyag.press('enter')  # CONFIRMANDO ARQUIVO

                time.sleep(15)  # TEMPO DE CARREGAMENTO

                pyag.click(x=172, y=9)  # ABA DE LISTAR
                pyag.hotkey("fn", "f5")  # ATUALIZANDO ABA

                time.sleep(2)  # TEMPO DE ESPERA

            elif tipo_arquivo == 'DEMISSIONAL' or tipo_arquivo == 'DOCUMENTOS DEMISSIONAIS':

                nome_pesquisa = f'{nome_funcionario} - DOCUMENTOS DEMISSIONAIS'  # NOME DE FUNCIONARIO E O TIPO DO ARQUIVO (TABELA NA DOCSTRING)
                pyag.write(nome_pesquisa.upper())  # ESCREVENDO NOME DO FUNCIONARIO TUDO EM MAIUSCULO
                pyag.press("enter")  # CONFIRMAR PESQUISA
                time.sleep(3)  # PAUSA CARREGAMENTO

                # DELETE DO ARQUIVO ANTIGO
                pyag.click(x=1811, y=652)  # CLICANDO EM FUNCOES NO ARQUIVO
                pyag.click(x=1772, y=757)  # CLICANDO EM APAGAR
                pyag.click(x=1031, y=185)  # CONFIRMANDO DELETE

                # UPLOAD ARQUIVO NOVO
                pyag.click(x=383, y=15)  # SELECIONANDO ABA PARA UPLOAD
                pyag.click(x=797, y=411)  # SELECIONANDO CONTEUDO
                pyag.write('RECURSOS HUMANOS')  # CONTEUDO DESEJADO
                pyag.click(x=954, y=523)  # CLICANDO NA TELA PARA ABRIR A SELECAO DE ARQUIVOS

                # ABA DE ESCOLHER ARQUIVOS
                pyag.click(x=78, y=133)  # (TELA DO EXPLORADOR DE ARQUIVOS) APERTANDO NO INICIO
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('Clientes')  # DIRETORIO PESQUISADO
                pyag.doubleClick(x=843, y=146)  # SELECIONANDO PESQUISA
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('AGI')  # SUB-DIRETORIO PESQUISADO
                pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('PRONTUARIOS')  # SUB-DIRETORIO PESQUISADO
                pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA

                pyag.click(x=770, y=966)  # BARRA DE PESQUISA
                pyag.write(nome_funcionario)  # NOME DO FUNCIONARIO PARA PROCURAR SUB-DIRETORIO
                pyag.press('enter')  # CONFIRMANDO SUB-DIRETORIO
                pyag.write(f"{nome_pesquisa}.pdf")  # ESCOLHENDO ARQUIVO QUE DESEJA FAZER O UPLOAD
                pyag.press('enter')  # CONFIRMANDO ARQUIVO

                time.sleep(15)  # TEMPO DE CARREGAMENTO

                pyag.click(x=172, y=9)  # ABA DE LISTAR
                pyag.hotkey("fn", "f5")  # ATUALIZANDO ABA

                time.sleep(2)  # TEMPO DE ESPERA

            elif tipo_arquivo == 'TREINAMENTOS E CERTIFICADOS' or tipo_arquivo == 'TREINAMENTO E CERTIFICADO' \
                    or tipo_arquivo == 'TREINAMENTO E CERTIFICADOS' or tipo_arquivo == 'TREINAMENTOS E CERTIFICADO':

                nome_pesquisa = f'{nome_funcionario} - TREINAMENTOS E CERTIFICADOS'  # NOME DE FUNCIONARIO E O TIPO DO ARQUIVO (TABELA NA DOCSTRING)
                pyag.write(nome_pesquisa.upper())  # ESCREVENDO NOME DO FUNCIONARIO TUDO EM MAIUSCULO
                pyag.press("enter")  # CONFIRMAR PESQUISA
                time.sleep(3)  # PAUSA CARREGAMENTO

                # DELETE DO ARQUIVO ANTIGO
                pyag.click(x=1811, y=652)  # CLICANDO EM FUNCOES NO ARQUIVO
                pyag.click(x=1772, y=757)  # CLICANDO EM APAGAR
                pyag.click(x=1031, y=185)  # CONFIRMANDO DELETE

                # UPLOAD ARQUIVO NOVO
                pyag.click(x=383, y=15)  # SELECIONANDO ABA PARA UPLOAD
                pyag.click(x=797, y=411)  # SELECIONANDO CONTEUDO
                pyag.write('RECURSOS HUMANOS')  # CONTEUDO DESEJADO
                pyag.click(x=954, y=523)  # CLICANDO NA TELA PARA ABRIR A SELECAO DE ARQUIVOS

                # ABA DE ESCOLHER ARQUIVOS
                pyag.click(x=78, y=133)  # (TELA DO EXPLORADOR DE ARQUIVOS) APERTANDO NO INICIO
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('Clientes')  # DIRETORIO PESQUISADO
                pyag.doubleClick(x=843, y=146)  # SELECIONANDO PESQUISA
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('AGI')  # SUB-DIRETORIO PESQUISADO
                pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA
                pyag.click(x=1796, y=60)  # BARRA DE PESQUISA
                pyag.write('PRONTUARIOS')  # SUB-DIRETORIO PESQUISADO
                pyag.doubleClick(x=905, y=142)  # SELECIONANDO PESQUISA

                pyag.click(x=770, y=966)  # BARRA DE PESQUISA
                pyag.write(nome_funcionario)  # NOME DO FUNCIONARIO PARA PROCURAR SUB-DIRETORIO
                pyag.press('enter')  # CONFIRMANDO SUB-DIRETORIO
                pyag.write(f"{nome_pesquisa}.pdf")  # ESCOLHENDO ARQUIVO QUE DESEJA FAZER O UPLOAD
                pyag.press('enter')  # CONFIRMANDO ARQUIVO

                time.sleep(15)  # TEMPO DE CARREGAMENTO

                pyag.click(x=172, y=9)  # ABA DE LISTAR
                pyag.hotkey("fn", "f5")  # ATUALIZANDO ABA

                time.sleep(2)  # TEMPO DE ESPERA
