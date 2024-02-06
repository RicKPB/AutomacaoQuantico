class Arquivo:

    # Parametros de um Arquivo;
    def __init__(self, name_file):
        self.name_file = name_file

    # Função para fragmentar o nome do arquivo em duas partes (nome_funcionario e tipo de arquivo);
    def fragmentName(self):

        # Fragementando o nome do arquivo atraves da caracter '-';
        fragment_file = self.name_file.split('-')

        # Deletando os espaços em branco do nome fracionado;
        fragment_file = [file.strip() for file in fragment_file]

        # Teste para conferir se o nome fracionado possui 2 ou mais partes;
        if len(fragment_file) >= 2:
            # Dividindo as partes em 2 tipos (nome_funcionario e tipo_arquivo) pegando elas da variavel fracionada do,
            # nome do arquivo;
            name_employee = fragment_file[0]
            file_type = fragment_file[1]

            # Realizando um teste para que se ouver mais de duas partes a 3 parte em diante ficar junta a 2 parte;
            if len(fragment_file) > 2:
                pass

            return name_employee, file_type

        else:
            print('Formato de nome invalido para o sistema.')
            return None, None