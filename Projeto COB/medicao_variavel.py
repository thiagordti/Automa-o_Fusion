from utils import *

def medicao_vr(caminho,usuario,senha,cod_filial = '01MG0014',cod_uo = '10310',tempo_espera = 0.5):
    planilha = pd.read_excel(caminho,'Medição').applymap(lambda x: str(x).replace('\xa0', '') if isinstance(x, str) else x) # Carrega a Planilha
    destino = os.path.dirname(caminho)# Pega o caminho da pasta
    planilha_destino = destino + r'/Historico.xlsx' # Caminho do Historico
    local_destino = r'C:\Temp\Historico.xlsx'
    copiar_para_planilha(local_destino, planilha_destino)
    navegador = iniciar_navegador()
    enviarkey_elemento(navegador,'user',By.ID,usuario)# Login
    enviarkey_elemento(navegador,'pass',By.ID,senha)# Senha
    clicar_elemento(navegador,'btnLogin',By.ID) # Clica no botão de Login
    acessar_iframe_default(navegador,tempo_espera)# Acessa o Iframe
    for linha in range(len(planilha)):
        enviarkey_elemento(navegador,'searchBarProcessQuery',By.ID,planilha.iloc[linha]['COB'])#Envio do COB
        esperar_elementos_carregar(navegador)
        clicar_elemento_rustico(navegador,'//*[@id="page-content-wrapper"]/div/div/div[1]/div[1]/nav/div/form/div/div/span/button',By.XPATH) # Clica no botão de pesquisa inicial
        aba_original = navegador.window_handles[0] # Identifica Aba Primaria
        clicar_elemento_rustico(navegador, 'header', By.CLASS_NAME) # Clica no COB pesquisado
        WebDriverWait(navegador, 10).until(lambda d: len(d.window_handles) > 1)
        nova_aba = navegador.window_handles[1]# Identifica nova aba apos iniciar Cobrança
        navegador.switch_to.window(nova_aba) # Troca para nova Aba
        nome_cob = texto_elemento(navegador,'headerTitle',By.ID)
        data = planilha.iloc[linha]['DATA_DESCRIÇÃO'] # Pega data de Descrição
        date = datetime.strptime(data.strftime('%d/%m/%Y'), '%d/%m/%Y') # Transforma data em string
        primeiro_dia, ultimo_dia = primeiro_e_ultimo_dia_do_mes(date.year, date.month) # Pega o mês e dia
        if pd.isna(planilha.iloc[linha]['DATA_DE_VENCIMENTO']):
            data_venc = date
        else:
            data_venc = planilha.iloc[linha]['DATA_DE_VENCIMENTO'] # Pega data de Vencimento
        # ---------------------- Esta Parte se refere ao COB sem Rateio ------------------------
        for sem_rateio in range(2):
            if pd.isna(planilha.iloc[linha][f'CR-SR{sem_rateio+1}']): # Verifica se o campo está vazio
                pass
            else:
                clicar_elemento(navegador,'createitem',By.ID)# Clica para criar novo Item
                acessar_iframe(navegador,tempo_espera)# Acessa o Iframe
                enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(planilha.iloc[linha]['CNPJ'])) # Envia CNPJ
                clicar_elemento_dinamico(navegador) # Clica no CNPJ informado
                if sem_rateio == 0: # Difere o primeiro produto do segundo
                    enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA SESI VIVA+: AEP,PGR,PCMSO,LTCAT \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
                else:
                    enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
                enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Não')# Envia não ao campo de rateio
                clicar_elemento(navegador,'id_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDeCobranca__UOCRProtheus___anchor',By.ID)# Clica na pesquisa de produto
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                clicar_elemento(navegador,'vfilter',By.ID) # Clica no Filtro
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe do Filtro
                enviarkey_elemento(navegador,'var_codclvlr__',By.NAME,str(int(planilha.iloc[linha]['CLASSE DE VALOR']))) # Envia Classe de valor Cliente
                enviarkey_elemento(navegador,'var_codfilialprotheus__',By.NAME,cod_filial) # Envia COD FILIAL - PADRÃO
                enviarkey_elemento(navegador,'var_coduo__',By.NAME,cod_uo) # Envia COD UO - PADRÃO
                enviarkey_elemento(navegador,'var_codccusto__',By.NAME,str(int(planilha.iloc[linha][f'CR-SR{sem_rateio+1}']))) # Envia COD PRODUTO
                clicar_elemento(navegador,'searchbutton',By.ID) # Clica na Pesquisa
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                clicar_elemento(navegador,'tooltip0',By.ID)
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario
                clicar_elemento(navegador,'createitem',By.ID) # Clica para adicionar Valor
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe de valor
                opcoes_pagamento(navegador,'//*[@id="mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos_ori"]/option[1]','move_this_right_mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos')#Loop para selecionar as opções de pagamento
                enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__data__',By.NAME,data_venc.strftime('%d/%m/%Y')) # Envia data da cobrança
                enviarkey_java(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__valor__',planilha.iloc[linha][f'VALORSR{sem_rateio+1}']) # Envia Valor
                clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario
                enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__numeroContratoProtheus__',By.ID,str(int(planilha.iloc[linha]['NUMERO DO CONTRATO']))) # Envia o numero de contrato
                clicar_elemento(navegador,'//*[@id="ui-id-10"]/li',By.XPATH) # Clica no numero de contrato
                clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
                navegador.switch_to.default_content()#Volta para o inicio

        # ---------------------- Esta Parte se refere ao COB com Rateio ------------------------

        contador = 0 # Contador utilizado para clicar nos rateios no processo Final!
        contador_1 = 0 # Contador utilizado para clicar nos rateios no processo Final!
        contador_2 = 0 # Contador utilizado para clicar nos rateios no processo Final!
        if pd.isna(planilha.iloc[linha]['CRR1']): # Verifica se o primeiro item está vazio, se o mesmo estiver vazio, todo o loop é pulado!
            pass

        elif pd.isna(planilha.iloc[linha]['QTD RATEIO']) or int(planilha.iloc[linha]['QTD RATEIO']) == 1 : # Caso não esteja vazio é iniciado o processo de Rateio e a QTD seja um executa todos os rateios em um unico processo
            clicar_elemento(navegador,'createitem',By.ID)# Clica para criar novo Item
            acessar_iframe(navegador,tempo_espera)# Acessa o Iframe
            enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(planilha.iloc[linha]['CNPJ'])) # Envia CNPJ
            clicar_elemento_dinamico(navegador) # Clica no CNPJ informado
            enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
            enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Sim')# Envia sim ao campo de rateio
            #Loop para a quantidade de Itens
            for com_rateio in range(4): # Loop para verificar todos os itens (Total 4) com rateio na planilha!!
                if pd.isna(planilha.iloc[linha][f'CRR{com_rateio+1}']): # Loop para verificar se o Item está vazio!!
                    pass # Pula o item vazio
                else:
                    dados_rateio(navegador,linha,cod_filial,cod_uo,planilha, tempo_espera)
                    enviarkey_elemento(navegador,'var_codccusto__',By.NAME,str(int(planilha.iloc[linha][f'CRR{com_rateio+1}']))) # Envia COD PRODUTO
                    clicar_elemento(navegador,'searchbutton',By.ID) # Clica na Pesquisa
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                    clicar_elemento(navegador,'tooltip0',By.ID) # Clica no item filtrado
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario3
                    enviarkey_java(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__',planilha.iloc[linha][f'VALOR{com_rateio+1}']) # Envia Valor
                    clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario
                    contador += 1 # Soma 1 a quantidade de contador, será utiizado para clicar no loop Contador!
            clicar_porcentagem(navegador,contador,linha,planilha, tempo_espera) # Baseado na soma do Contador clica nos itens

        elif int(planilha.iloc[linha]['QTD RATEIO']) == 2: # Ira rodar o processo de sem rateio duas vezes uma para a coluna CRR1 e 2 e ou para CRR3 e 4
            # Processo para coluna 1 e 2
            clicar_elemento(navegador,'createitem',By.ID)# Clica para criar novo Item
            acessar_iframe(navegador,tempo_espera)# Acessa o Iframe
            enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(planilha.iloc[linha]['CNPJ'])) # Envia CNPJ
            clicar_elemento_dinamico(navegador) # Clica no CNPJ informado
            enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
            enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Sim')# Envia sim ao campo de rateio
            #Loop para a quantidade de Itens
            for com_rateio in range(2): # Loop para verificar todos os itens (Total 4) com rateio na planilha!!
                if pd.isna(planilha.iloc[linha][f'CRR{com_rateio+1}']): # Loop para verificar se o Item está vazio!!
                    pass # Pula o item vazio
                else:
                    dados_rateio(navegador,linha,cod_filial,cod_uo,planilha, tempo_espera)
                    enviarkey_elemento(navegador,'var_codccusto__',By.NAME,str(int(planilha.iloc[linha][f'CRR{com_rateio+1}']))) # Envia COD PRODUTO
                    clicar_elemento(navegador,'searchbutton',By.ID) # Clica na Pesquisa
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                    clicar_elemento(navegador,'tooltip0',By.ID) # Clica no item filtrado
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario3
                    enviarkey_java(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__',planilha.iloc[linha][f'VALOR{com_rateio+1}']) # Envia Valor
                    clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario
                    contador_1 += 1 # Soma 1 a quantidade de contador, será utiizado para clicar no loop Contador!
            clicar_porcentagem(navegador,contador_1,linha,planilha, tempo_espera) # Baseado na soma do Contador clica nos itens

            # Processo para coluna 3 e 4
            clicar_elemento(navegador,'createitem',By.ID)# Clica para criar novo Item
            acessar_iframe(navegador,tempo_espera)# Acessa o Iframe
            enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(planilha.iloc[linha]['CNPJ'])) # Envia CNPJ
            clicar_elemento_dinamico(navegador) # Clica no CNPJ informado
            enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'IN LOCO COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
            enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Sim')# Envia sim ao campo de rateio
            #Loop para a quantidade de Itens
            for com_rateio in range(2): # Loop para verificar todos os itens (Total 4) com rateio na planilha!!
                if pd.isna(planilha.iloc[linha][f'CRR{com_rateio+3}']): # Loop para verificar se o Item está vazio!!
                    pass # Pula o item vazio
                else:
                    dados_rateio(navegador,linha,cod_filial,cod_uo,planilha, tempo_espera)
                    enviarkey_elemento(navegador,'var_codccusto__',By.NAME,str(int(planilha.iloc[linha][f'CRR{com_rateio+3}']))) # Envia COD PRODUTO
                    clicar_elemento(navegador,'searchbutton',By.ID) # Clica na Pesquisa
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                    clicar_elemento(navegador,'tooltip0',By.ID) # Clica no item filtrado
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario3
                    enviarkey_java(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__',planilha.iloc[linha][f'VALOR{com_rateio+3}']) # Envia Valor
                    clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario
                    contador_2 += 1 # Soma 1 a quantidade de contador, será utiizado para clicar no loop Contador!
            clicar_porcentagem(navegador,contador_2,linha,planilha, tempo_espera) # Baseado na soma do Contador clica nos itens

        # ---------------------- Esta Parte se refere aos Anexos ------------------------
        enviar_anexo(navegador,linha,'//*[@id="menu_bar_genericoHistoricoAtendimento"]/li[1]','var_dadosDaCobranca__historico__anexo__','//*[@id="progress-complete-var_dadosDaCobranca__historico__anexo__"]/span','var_dadosDaCobranca__historico__registro__',planilha,caminho,tempo_espera) # Envia Anexos
        if len(navegador.find_elements(By.ID, 'id_dadosDaCobranca__acao__')) >= 1: # Verifica se o campo existe
            enviarkey_elemento(navegador,'id_dadosDaCobranca__acao__',By.ID,'Solicitar Nova Medição')
        input('Confirma o lançamento!!!')
        clicar_elemento(navegador,'action.send',By.NAME)
        esperar_alerta(navegador,nome_cob, aba_original,planilha,local_destino,'Medição',linha)
        time.sleep(1)
        acessar_iframe_default(navegador,tempo_espera)
        clicar_elemento_rustico(navegador,'clear-input-filter',By.CLASS_NAME)#Limpa o campo de Pesquisa
    copiar_para_planilha(planilha_destino,local_destino)
