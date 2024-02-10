import pandas as pd

class PlanilhaExcel:

    def __init__(
            self,
            nome_arquivo
    ):
        self.nome_arquivo = nome_arquivo

    def criar_data_frame (self, caminho_diretorio):

        caminho_diretorio = caminho_diretorio
        data_frame = {
            'Coluna1': [],
            'Coluna2': []
        }
        df = pd.DataFrame(data_frame)

