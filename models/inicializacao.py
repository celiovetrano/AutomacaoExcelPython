import os

class Inicializacao:
    # Função para limpar a tela
    @staticmethod
    def limpar_tela():
        os.system('cls' if os.name == 'nt' else 'clear')

    @staticmethod
    def exibir_nome_aplicaçao():
        print("Controle de ponto Servidores")
    
    @staticmethod
    # Função principal do menu
    def menu():
        print("\nMenu:")
        print("1. Cadastrar funcionário")
        print("2. Consultar funcionários")
        print("3. Alterar dados do funcionário")
        print("4. Excluir funcionário")
        print("5. Gerar planilhas de horas")
        print("6. Sair\n")
        
