if __name__ == "__main__":
    from utils import copiar_para_planilha, selecionar_arquivo
    from models import AutomacaoFusion
    from getpass import getpass
    import os

    print('-----------Automação COB---------\n')

    tempo_fusion = input("O Site Fusion está mais lento que o normal?? Responda com S ou N: ")  # Validação para o código rodar sem travar devido a lentidão do site!
    if tempo_fusion.lower() == 's':
        tempo_carregar = 1.5

    usuario = input('Insira o usuario do Fusion: ')
    senha = getpass('Insira a senha do Fusion: ')

    caminho = selecionar_arquivo()

    destino = os.path.dirname(caminho)  # Pega o caminho da pasta
    planilha_destino = destino + '/Historico.xlsx'  # Caminho do Historico
    local_destino = 'C:/Temp/Historico.xlsx'

    while True:

        print("\nMenu de Escolhas:")
        print("1 - Medição Variável")
        print("2 - Criar Novos COB's")
        print("0 - Sair")
        escolha = input("Escolha uma opção: ")

        try:
            escolha = int(escolha)

            if escolha == 1:
                global_instance = AutomacaoFusion(caminho, None, None, planilha_destino, local_destino, "medicao_vr", cod_filial='01MG0014', cod_uo='10310', tempo_espera=0.5)
                navegador, planilha = global_instance.inicializacao(usuario, senha, "Medição")
                global_instance.navegador = navegador
                global_instance.planilha = planilha
                global_instance.medicao_vr()


            elif escolha == 2:
                global_instance = AutomacaoFusion(caminho, None, None, planilha_destino, local_destino, "cob_nv", cod_filial='01MG0014', cod_uo='10310', tempo_espera=0.5)
                navegador, planilha = global_instance.inicializacao(usuario, senha, "Novo")
                global_instance.navegador = navegador
                global_instance.planilha = planilha
                global_instance.cob_nv()

            elif escolha == 0:
                print("Saindo...")
                break

            else:
                print("Escolha inválida. Tente novamente.")
        except ValueError as e:
            print(f"Ocorreu um erro: {e}")
            copiar_para_planilha(planilha_destino, local_destino)
            input('Chame a T.I')
        except Exception as e:
            print(f"Ocorreu um erro: {e}")
            copiar_para_planilha(planilha_destino, local_destino)
            input('Chame a T.I')
