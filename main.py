if __name__ == "__main__":
    from utils import copiar_para_planilha, selecionar_arquivo
    from models import AutomacaoFusion
    import os

    print('-----------Automação COB---------\n')

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
                global_instance = AutomacaoFusion(caminho, None, None, None,  planilha_destino, local_destino, "medicao_vr", cod_filial='01MG0014', cod_uo='10310')
                navegador, chrome_proc, planilha = global_instance.inicializacao("Medição")
                global_instance.navegador = navegador
                global_instance.chrome_proc = chrome_proc
                global_instance.planilha = planilha
                global_instance.medicao_vr()

            elif escolha == 2:
                global_instance = AutomacaoFusion(caminho, None, None, None, planilha_destino, local_destino, "cob_nv", cod_filial='01MG0014', cod_uo='10310')
                navegador, chrome_proc, planilha = global_instance.inicializacao("Novo")
                global_instance.navegador = navegador
                global_instance.chrome_proc = chrome_proc
                global_instance.planilha = planilha
                global_instance.cob_nv()

            elif escolha == 0:
                print("Saindo...")
                break

            else:
                print("Escolha inválida. Tente novamente.")
        except Exception as e:
            if global_instance:
                tipo = "Medição" if escolha == 1 else "Novo"
                global_instance.tratar_erro_critico(e, tipo)
            else:
                print(f"Ocorreu um erro: {e}")
                copiar_para_planilha(planilha_destino, local_destino)
                input('Chame a T.I')