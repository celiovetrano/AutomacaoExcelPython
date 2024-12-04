import os
import calendar
from models.inicializacao import Inicializacao
from models.gerar_planilha import GerarPlanilha
from models.cadastro_funcionario import CadastroFuncionario

class SalvarPlanilhasFuncionarios:
    # Função para gerar planilhas de horas de cada funcionario da lista
    def salvar_planilhas():
        cadastro = CadastroFuncionario()  # Criar uma instância da classe
        funcionarios = cadastro.carregar_funcionarios()
        if not funcionarios:
            print("Nenhum funcionário cadastrado.")
            return

        # Garante que o diretório 'planilhas' exista
        diretorio = "planilhas"
        if not os.path.exists(diretorio):
            os.makedirs(diretorio)

        # Obter a diretoria regional de ensino que a escola está inserida:
        dre = input("Informe a DRE da unidade (Exemplo Guaianases): ")
        escola = input("Informe o nome da Escola: ")

        planilha = GerarPlanilha()  # Criar uma instância da classe GerarPlanilha
        mes = planilha.obter_inteiro_valido("Por favor digite o Mês atual com 2 dígitos (MM): ", 1, 12)
        ano = planilha.obter_inteiro_valido("Por favor digite o Ano atual com 4 dígitos (AAAA): ", 1900, 2100)
        recesso = planilha.obter_resposta_sim_nao("Há recesso neste mês? (s/n): ") == 's'
        nome_nova_aba = f"{planilha.obter_nome_mes(mes)}_{ano}"

        if recesso:
            recesso_inicio = planilha.obter_inteiro_valido("Por favor digite o dia de início do recesso (DD): ", 1, calendar.monthrange(ano, mes)[1])
            recesso_fim = planilha.obter_inteiro_valido("Por favor digite o dia de fim do recesso (DD): ", recesso_inicio, calendar.monthrange(ano, mes)[1])
        else:
            recesso_inicio, recesso_fim = None, None

        for rf, dados in funcionarios.items():
            nome, qpe, inicio_exercicio, horarios_jeif, horarios_regencia, serie_regencia, cargo = dados
            primeiro_nome = nome.split()[0]
            rf_formatado = rf # Mantém o formato original do RF

            # Caminho do arquivo de origem
            arquivo_origem = "exemplo.xlsm"

            # Gera a Planilha para cada servidor
            planilha.editar_aba(
                dre, escola, arquivo_origem, nome_nova_aba, mes, ano, nome, rf_formatado, 
                qpe, inicio_exercicio, horarios_jeif, horarios_regencia, serie_regencia,
                cargo, diretorio, recesso_inicio, recesso_fim
            )

        if input("Pressione 'Espaço' para retornar ao menu: ") == ' ':
            Inicializacao.limpar_tela()
            Inicializacao.exibir_nome_aplicaçao()
            Inicializacao.menu()
            return
