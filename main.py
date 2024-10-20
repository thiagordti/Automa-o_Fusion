if __name__ == "__main__":
    from utils import copiar_para_planilha, selecionar_arquivo,inicializacao
    from cobnv_class import CobNV
    from getpass import getpass
    from medicao_class import MedicaoVR
    import os
    import globals
    import traceback 

    print('-----------Automação COB-----------\n')

    tempo_fusion = input("O Site Fusion está mais lento que o normal?? Responda com S ou N: ")  # Validação para o código rodar sem travar devido a lentidão do site!
    if tempo_fusion.lower() == 's':
        tempo_espera = 1.5

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
                navegador, planilha = inicializacao(caminho, "Medição", local_destino, planilha_destino, usuario, senha, tempo_espera=0.5)
                globals.global_instance = MedicaoVR(caminho, navegador, planilha, planilha_destino, local_destino, cod_filial='01MG0014', cod_uo='10310', tempo_espera=0.5)
                globals.global_instance.medicao_vr()  # Chama o método medicao_vr da instância global_medicao_instance

            elif escolha == 2:
                navegador, planilha = inicializacao(caminho,"Novo",local_destino,planilha_destino,usuario, senha, tempo_espera=0.5)
                globals.global_instance = CobNV(caminho,navegador, planilha,planilha_destino,local_destino, cod_filial='01MG0014', cod_uo='10310', tempo_espera=0.5)
                globals.global_instance.cob_nv() # Chama o método cob_nv da instância global_cobnv_instance

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
            print(traceback.format_exc())  # Imprime o stack trace completo
            copiar_para_planilha(planilha_destino, local_destino)
            input('Chame a T.I')
