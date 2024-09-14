from utils import *

def cob_nv(caminho,usuario,senha,cod_filial = '01MG0014',cod_uo = '10310',tempo_espera=0.5):
    planilha = pd.read_excel(caminho,'Novo') # Carrega a Planilha
    destino = os.path.dirname(caminho) # Pega o caminho da pasta
    planilha_destino = destino + r'/Historico.xlsx' # Caminho do Historico
    local_destino = r'C:\Temp\Historico.xlsx'
    copiar_para_planilha(local_destino, planilha_destino)
    navegador = iniciar_navegador() #Inicia o Navegador
    enviarkey_elemento(navegador,'user',By.ID,usuario)# Login
    enviarkey_elemento(navegador,'pass',By.ID,senha)# Senha
    clicar_elemento(navegador,'btnLogin',By.ID) # Clica no botão de Login
    acessar_iframe_default(navegador,tempo_espera)# Acessa o Iframe
    aba_original = navegador.window_handles[0] # Identifica Aba Primaria

    for linha in range(len(planilha)):
        clicar_elemento(navegador,'btnStartProcess',By.ID) # Iniciar novo processo
        clicar_elemento(navegador,'//*[@id="page-content-wrapper"]/div/div/div[1]/div[1]/nav/div/div/div/ul/li[3]/ul/li[5]/a/div/span[1]',By.XPATH) # Iniciar nova Cobrança
        WebDriverWait(navegador, 10).until(lambda d: len(d.window_handles) > 1)
        nova_aba = navegador.window_handles[1]# Identifica nova aba apos iniciar nova Cobrança
        navegador.switch_to.window(nova_aba) # Troca para nova Aba
        enviarkey_elemento(navegador,'id_informeNucleo__',By.ID,'Núcleo de Faturamento')# Envia nucleo - Padrão
        enviarkey_elemento(navegador,'id_tipoSolicitacao__',By.ID,'Solicitação de cobrança (FG-176)')# Solicitação de cobrança - Padrão
        enviarkey_elemento(navegador,'id_plataformaGestaoDaVenda__',By.ID,'Protheus')# Plataforma - Padrão

        if planilha.iloc[linha]['TIPO'].lower() == "variavel" and planilha.iloc[linha]['RATEIO'].lower() == "sim":

            data = planilha.iloc[linha]['DESCRICAO']
            date = datetime.strptime(data.strftime('%d/%m/%Y'), '%d/%m/%Y') # Transforma data em string
            primeiro_dia, ultimo_dia = primeiro_e_ultimo_dia_do_mes(date.year, date.month) # Pega o mês e dia
            data_str = planilha.iloc[linha]['DATA']
            data_obj = datetime.strptime(data_str.strftime('%d/%m/%Y'), '%d/%m/%Y')
            nome_cob = texto_elemento(navegador,'headerTitle',By.ID)
            variavel_novo(navegador,linha,planilha,primeiro_dia,ultimo_dia) # Carrega dados de preenchimento
            if pd.isna(planilha.iloc[linha][f'CR1']): # Loop para verificar se o Item está vazio!!
                pass
            else:
                clicar_elemento(navegador,'//*[@id="createitem"]',By.XPATH) # Clica no Novo Item
                acessar_iframe(navegador,tempo_espera)# Acessa o Iframe
                enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(planilha.iloc[linha]['CNPJ'])) # Envia CNPJ
                clicar_elemento(navegador,'ui-id-11',By.ID) # Clica no CNPJ informado
                enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA SESI VIVA+: AEP,PGR,PCMSO,LTCAT \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
                enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Não')# Envia não ao campo de rateio
                clicar_elemento(navegador,'id_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDeCobranca__UOCRProtheus___anchor',By.ID) # Clica na pesquisa
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                clicar_elemento(navegador,'vfilter',By.ID) # Clica no Filtro
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe do Filtro
                enviarkey_elemento(navegador,'var_codclvlr__',By.NAME,str(int(planilha.iloc[linha]['CLASSE DE VALOR']))) # Envia Classa de valor Cliente
                enviarkey_elemento(navegador,'var_codfilialprotheus__',By.NAME,cod_filial) # Envia COD FILIAL - PADRÃO
                enviarkey_elemento(navegador,'var_coduo__',By.NAME,cod_uo) # Envia COD UO - PADRÃO
                enviarkey_elemento(navegador,'var_codccusto__',By.NAME,int(planilha.iloc[linha]['CR1'])) # Envia COD PRODUTO
                clicar_elemento(navegador,'searchbutton',By.ID) # Clica na Pesquisa
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                clicar_elemento(navegador,'tooltip0',By.ID)
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario
                clicar_elemento(navegador,'createitem',By.ID) # Clica para adicionar Valor
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe de valor
                opcoes_pagamento(navegador,'//*[@id="mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos_ori"]/option[1]','move_this_right_mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos')#Loop para selecionar as opções de pagamento     
                enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__data__',By.ID,data_obj.strftime('%d/%m/%Y')) # Envia data da cobrança
                enviarkey_java(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__valor__',planilha.iloc[linha]['VALOR1'])
                clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario
            enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__numeroContratoProtheus__',By.ID,int(planilha.iloc[linha]['NUMERO DO CONTRATO'])) # Envia o numero de contrato
            clicar_elemento(navegador,'//*[@id="ui-id-10"]/li',By.XPATH) # Clica no numero de contrato
            clicar_elemento(navegador,'//*[@id="dibButtons"]/input[1]',By.XPATH) # Clica no numero de contrato
            navegador.switch_to.default_content()#Volta para o inicio
            #------------------------Rateio--------------------------------------------------
            clicar_elemento(navegador,'//*[@id="createitem"]',By.XPATH) # Clica no Novo
            acessar_iframe(navegador,tempo_espera)# Acessa o Iframe
            enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(planilha.iloc[linha]['CNPJ'])) # Envia CNPJ
            clicar_elemento(navegador,'ui-id-11',By.ID) # Clica no CNPJ informado
            enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
            enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'SIM')# Envia Sim ao campo de rateio

            contador = 0 # Contador utilizado para clicar nos rateios no processo Final!
            #Loop para a quantidade de Itens
            for com_rateio in range(5): # Loop para verificar todos os itens (Total 5) com rateio na planilha!!
                if pd.isna(planilha.iloc[linha][f'CR{com_rateio+2}']): # Loop para verificar se o Item está vazio!!
                    pass # Pula o item vazio
                else:
                    dados_rateio(navegador,linha,cod_filial,cod_uo,planilha)
                    enviarkey_elemento(navegador,'var_codccusto__',By.NAME,str(int(planilha.iloc[linha][f'CR{com_rateio+2}']))) # Envia COD PRODUTO
                    clicar_elemento(navegador,'searchbutton',By.ID) # Clica na Pesquisa
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                    clicar_elemento(navegador,'tooltip0',By.ID) # Clica no item filtrado
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario3
                    elemento2 = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.NAME, 'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__')))
                    script_valor_cr = f"document.getElementsByName('var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__')[0].value='{planilha.iloc[linha][f'VALOR{com_rateio+2}']}';"
                    navegador.execute_script(script_valor_cr)
                    clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario
                    contador += 1 # Soma 1 a quantidade de contador, será utiizado para clicar no loop Contador!

            for i in range(contador): # Baseado na soma do Contador clica nos itens
                clicar_elemento_rustico(navegador,f'//*[@id="{i}"]/td[2]',By.XPATH) # Clica no Item baseado nos indices (No fusion o indice 0 conta!)!!
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario
                clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
                acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario

            enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__numeroContratoProtheus__',By.ID,str(int(planilha.iloc[linha]['NUMERO DO CONTRATO']))) # Envia o numero de contrato
            clicar_elemento(navegador,'//*[@id="ui-id-10"]/li',By.XPATH) # Clica no numero de contrato
            clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
            navegador.switch_to.default_content()#Volta para o inicio
            
            enviar_anexo(navegador,linha,'//*[@id="menu_bar_genericoHistoricoAtendimento"]/li[1]','var_dadosDaCobranca__historico__anexo__','//*[@id="progress-complete-var_dadosDaCobranca__historico__anexo__"]/span','var_dadosDaCobranca__historico__registro__',planilha)#Envia Anexos
            enviar_emails(navegador,linha,"//li[@onclick=\"activeDeactiveObjMenu2(this);javascript: ellist_emailClienteFP__.addNewItem('CreateItens', true);\"]//a[@id='createitens']",'var_emailClienteFP__Email__',planilha) # Envia e-mails
            input('Confirma o lançamento!!!')
            navegador.switch_to.default_content()
            clicar_elemento(navegador,'action.send',By.NAME)
            esperar_alerta(navegador,nome_cob, aba_original,planilha,local_destino,'Novo',linha,nome_cob)
            time.sleep(1)
            acessar_iframe_default(navegador,tempo_espera)

        elif planilha.iloc[linha]['TIPO'].lower() == "variavel" and planilha.iloc[linha]['RATEIO'].lower() == "não":

            data = planilha.iloc[linha]['DESCRICAO']
            date = datetime.strptime(data.strftime('%d/%m/%Y'), '%d/%m/%Y') # Transforma data em string
            primeiro_dia, ultimo_dia = primeiro_e_ultimo_dia_do_mes(date.year, date.month) # Pega o mês e dia
            data_str = planilha.iloc[linha]['DATA']
            data_obj = datetime.strptime(data_str.strftime('%d/%m/%Y'), '%d/%m/%Y')
            nome_cob = texto_elemento(navegador,'headerTitle',By.ID)
            variavel_novo(navegador,linha,planilha,primeiro_dia,ultimo_dia) # Carrega dados de preenchimento
            clicar_elemento(navegador,'//*[@id="createitem"]',By.XPATH) #  Clica no Novo Item
            acessar_iframe(navegador,tempo_espera)# Acessa o Iframe
            enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(planilha.iloc[linha]['CNPJ'])) # Envia CNPJ
            clicar_elemento(navegador,'ui-id-11',By.ID) # Clica no CNPJ informado
            enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA SESI VIVA+: AEP,PGR,PCMSO,LTCAT \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
            enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Não')# Envia não ao campo de rateio
            for sem_rateio in range(2):
                if pd.isna(planilha.iloc[linha][f'CR{sem_rateio+1}']):
                    pass
                else:
                    clicar_elemento(navegador,'id_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDeCobranca__UOCRProtheus___anchor',By.ID)# Clica novo
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                    clicar_elemento(navegador,'vfilter',By.ID) # Clica no Filtro
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe do Filtro
                    enviarkey_elemento(navegador,'var_codclvlr__',By.NAME,str(int(planilha.iloc[linha]['CLASSE DE VALOR']))) # Envia Classa de valor Cliente
                    enviarkey_elemento(navegador,'var_codfilialprotheus__',By.NAME,cod_filial) # Envia COD FILIAL - PADRÃO
                    enviarkey_elemento(navegador,'var_coduo__',By.NAME,cod_uo) # Envia COD UO - PADRÃO
                    enviarkey_elemento(navegador,'var_codccusto__',By.NAME,int(planilha.iloc[linha][f'CR{sem_rateio+1}'])) # Envia COD PRODUTO
                    clicar_elemento(navegador,'searchbutton',By.ID) # Clica na Pesquisa
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                    clicar_elemento(navegador,'tooltip0',By.ID)
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario
                    clicar_elemento(navegador,'createitem',By.ID) # Clica para adicionar Valor
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe de valor
                    opcoes_pagamento(navegador,'//*[@id="mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos_ori"]/option[1]','move_this_right_mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos')#Loop para selecionar as opções de pagamento 
                    enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__data__',By.ID,data_obj.strftime('%d/%m/%Y')) # Envia data da cobrança
                    enviarkey_java(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__valor__',planilha.iloc[linha][f'VALOR{sem_rateio+1}'])
                    clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe primario
            enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__numeroContratoProtheus__',By.ID,int(planilha.iloc[linha]['NUMERO DO CONTRATO'])) # Envia o numero de contrato
            clicar_elemento(navegador,'//*[@id="ui-id-10"]/li',By.XPATH) # Clica no numero de contrato
            clicar_elemento(navegador,'//*[@id="dibButtons"]/input[1]',By.XPATH) # Clica no numero de contrato
            navegador.switch_to.default_content()#Volta para o inicio

            # ---------------------- Esta Parte se refere aos Anexos ------------------------
            enviar_anexo(navegador,linha,'//*[@id="menu_bar_genericoHistoricoAtendimento"]/li[1]','var_dadosDaCobranca__historico__anexo__','//*[@id="progress-complete-var_dadosDaCobranca__historico__anexo__"]/span','var_dadosDaCobranca__historico__registro__')#Envia Anexos
            enviar_emails(navegador,linha,"//li[@onclick=\"activeDeactiveObjMenu2(this);javascript: ellist_emailClienteFP__.addNewItem('CreateItens', true);\"]//a[@id='createitens']",'var_emailClienteFP__Email__',planilha)
            input('Confirma o lançamento!!!')
            navegador.switch_to.default_content()
            clicar_elemento(navegador,'action.send',By.NAME)
            esperar_alerta(navegador,nome_cob, aba_original,planilha,local_destino,'Novo',linha,nome_cob)
            time.sleep(1)
            acessar_iframe_default(navegador,tempo_espera)

        elif planilha.iloc[linha]['TIPO'].lower() == "fixo" and planilha.iloc[linha]['RATEIO'].lower() == "não":
            enviarkey_elemento(navegador,'id_tipoDeMedicao__',By.ID,'Fixa')# Tipo de medição
            clicar_elemento(navegador,'//*[@id="tab_bar_"]/li[2]/a',By.XPATH) # Clica na guia Dados de cobrança
            nome_cob = texto_elemento(navegador,'headerTitle',By.ID)
            enviarkey_elemento(navegador,'id_txt_dadosCobranca__DadosDoCliente__',By.ID,tratar_cnpj(planilha.iloc[linha]['CNPJ']))# Envia CNPJ
            clicar_elemento(navegador,'ui-menu-item',By.CLASS_NAME) # Clica no CNPJ informado
            enviarkey_elemento(navegador,'var_dadosCobranca__descricaoServico__',By.ID,planilha.iloc[linha]['DESCRICAO'])# Envia Descrição
            enviarkey_elemento(navegador,'var_dadosCobranca__dadosParaHistorico2__HouvePrestacaoDeServicos__',By.ID,'Sim')# Prestação de Serviço - Padrão
            enviarkey_elemento(navegador,'id_txt_dadosCobranca__dadosParaHistorico2__numeroContratoProtheus__',By.ID,str(int(planilha.iloc[linha]['NUMERO DO CONTRATO'])))# Numero de contrato
            clicar_elemento(navegador,'//*[@id="ui-id-6"]/li',By.XPATH) # Clica no contrato informado
            enviarkey_elemento(navegador,'id_dadosCobranca__dadosParaHistorico2__diaLimiteNFCliente__',By.ID,str(int(planilha.iloc[linha]['DIA LIMITE'])))# Data Limite
            opcoes_pagamento(navegador,'//*[@id="mul_dadosCobranca__formaDeEntradaDosRecursosRevisado_ori"]/option[1]','move_this_right_mul_dadosCobranca__formaDeEntradaDosRecursosRevisado') # Loop para opções de pagamento
            enviarkey_elemento(navegador,'//*[@id="var_dadosCobranca__rateio__"]',By.XPATH,'Não')# Tipo de Rateio

            clicar_elemento(navegador,'id_dadosCobranca__filaisSemRateio__UOCRProtheus___anchor',By.ID) # Pequisa
            acessar_iframe_default(navegador,tempo_espera)# Acessa o Iframe

            clicar_elemento(navegador,'vfilter',By.ID) # Clica no Filtro
            acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe do Filtro
            enviarkey_elemento(navegador,'var_codclvlr__',By.NAME,str(int(planilha.iloc[linha]['CLASSE DE VALOR']))) # Envia Classa de valor Cliente
            enviarkey_elemento(navegador,'var_codfilialprotheus__',By.NAME,cod_filial) # Envia COD FILIAL - PADRÃO
            enviarkey_elemento(navegador,'var_coduo__',By.NAME,cod_uo) # Envia COD UO - PADRÃO
            enviarkey_elemento(navegador,'var_codccusto__',By.NAME,str(int(planilha.iloc[linha]['CR1']))) # Envia COD PRODUTO
            clicar_elemento(navegador,'searchbutton',By.ID) # Clica na Pesquisa
            acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
            clicar_elemento(navegador,'tooltip0',By.ID)
            navegador.switch_to.default_content()

            valor = float(planilha.iloc[linha]['VALOR1'])
            parcelas = int(planilha.iloc[linha]['PARCELA'])
            valor_parcela = valor / parcelas
            date = planilha.iloc[linha]['DATA']

            clicar_elemento(navegador,'//*[@id="menu_bar_FinCobFFDatasVencimentos"]/li[2]',By.XPATH) # Itens novos

            for i in range(int(planilha.iloc[linha]['PARCELA'])):
                    new_data = date + relativedelta(months=i)
                    new_date_str = new_data.strftime('%d/%m/%Y')
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe dos itens novos
                    enviarkey_elemento(navegador,'var_dadosCobranca__cobrancas__data__',By.ID,new_date_str) # Envia Data
                    enviarkey_java(navegador,'var_dadosCobranca__cobrancas__valor__',valor_parcela) # Envia Valor
                    clicar_elemento(navegador,'//*[@id="dibButtons"]/input[1]',By.XPATH) # Botão Ok
            acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe dos itens novos
            clicar_elemento(navegador,'cancelButtonModal',By.ID) # Botão Cancelar
            navegador.switch_to.default_content()
            enviarkey_elemento(navegador,'var_dadosCobranca__Observacao__',By.ID,planilha.iloc[linha]['OBSERVACAO']) # Envia Observação
            enviar_emails(navegador,linha,"//li[@onclick=\"activeDeactiveObjMenu2(this);javascript: ellist_EmailDeContatoDosClientes__.addNewItem('CreateItens', true);\"]/a[@id='createitens']",'var_EmailDeContatoDosClientes__Email__',planilha)
            input('Confirma o lançamento!!!')
            navegador.switch_to.default_content()
            clicar_elemento(navegador,'action.send',By.NAME)
            esperar_alerta(navegador,nome_cob, aba_original,planilha,local_destino,'Novo',linha,nome_cob)
            time.sleep(1)
            acessar_iframe_default(navegador,tempo_espera)

        elif planilha.iloc[linha]['TIPO'].lower() == "fixo" and planilha.iloc[linha]['RATEIO'].lower() == "sim":
            enviarkey_elemento(navegador,'id_tipoDeMedicao__',By.ID,'Fixa')# Tipo de medição
            clicar_elemento(navegador,'//*[@id="tab_bar_"]/li[2]/a',By.XPATH) # Clica na guia Dados de cobrança
            nome_cob = texto_elemento(navegador,'headerTitle',By.ID)
            enviarkey_elemento(navegador,'id_txt_dadosCobranca__DadosDoCliente__',By.ID,tratar_cnpj(planilha.iloc[linha]['CNPJ']))# Envia CNPJ
            clicar_elemento(navegador,'ui-menu-item',By.CLASS_NAME) # Clica no CNPJ informado
            enviarkey_elemento(navegador,'var_dadosCobranca__descricaoServico__',By.ID,planilha.iloc[linha]['DESCRICAO'])# Envia Descrição
            enviarkey_elemento(navegador,'var_dadosCobranca__dadosParaHistorico2__HouvePrestacaoDeServicos__',By.ID,'Sim')# Prestação de Serviço - Padrão
            enviarkey_elemento(navegador,'id_txt_dadosCobranca__dadosParaHistorico2__numeroContratoProtheus__',By.ID,str(int(planilha.iloc[linha]['NUMERO DO CONTRATO'])))# Numero de contrato
            clicar_elemento(navegador,'//*[@id="ui-id-6"]/li',By.XPATH) # Clica no contrato informado
            enviarkey_elemento(navegador,'id_dadosCobranca__dadosParaHistorico2__diaLimiteNFCliente__',By.ID,str(int(planilha.iloc[linha]['DIA LIMITE'])))# Data Limite
            opcoes_pagamento(navegador,'//*[@id="mul_dadosCobranca__formaDeEntradaDosRecursosRevisado_ori"]/option[1]','move_this_right_mul_dadosCobranca__formaDeEntradaDosRecursosRevisado')# Loop para opções de pagamento
            enviarkey_elemento(navegador,'//*[@id="var_dadosCobranca__rateio__"]',By.XPATH,'Sim')# Tipo de Rateio
            parcemlamento = 0
            for i in range(6):
                if not pd.isna(planilha.iloc[linha][f'CR{i+1}']):
                    parcemlamento += 1
            date = planilha.iloc[0]['DATA']
            for i in range(int(planilha.iloc[linha]['PARCELA'])):
                clicar_elemento(navegador,'//*[@id="menu_bar_finCobDataXFilialXCcusto"]/li[1]',By.XPATH) # Clica produtos novo
                acessar_iframe_default(navegador,tempo_espera)# Acessa o Iframe
                new_data = date + relativedelta(months=i)
                new_date_str = new_data.strftime('%d/%m/%Y')
                enviarkey_elemento(navegador,'var_dadosCobranca__cobRateio__dataCobranca__',By.ID,new_date_str) # Envia Data
                contador = 0
                for i in range(parcemlamento):
                    clicar_elemento(navegador,'//*[@id="menu_bar_finCobFFCentroValor"]/li[1]',By.XPATH) # Pequisa
                    acessar_iframe_default(navegador,tempo_espera)# Acessa o Iframe
                    enviarkey_elemento(navegador,'id_txt_dadosCobranca__cobRateio__filialCustos__FilialProtheus__',By.ID,cod_filial) # Envia COD FILIAL - PADRÃO
                    clicar_elemento(navegador,'ui-menu-item',By.CLASS_NAME)
                    enviarkey_elemento(navegador,'id_txt_dadosCobranca__cobRateio__filialCustos__UO__',By.ID,cod_uo) # Envia COD FILIAL - PADRÃO
                    clicar_elemento(navegador,'//*[@id="ui-id-4"]/li',By.XPATH)
                    clicar_elemento(navegador,'id_dadosCobranca__cobRateio__filialCustos__UOXCRProtheus___anchor',By.ID)

                    acessar_iframe_default(navegador,tempo_espera)# Acessa o Iframe
                    clicar_elemento(navegador,'vfilter',By.ID) # Clica no Filtro
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe do Filtro
                    enviarkey_elemento(navegador,'var_codclvlr__',By.NAME,str(int(planilha.iloc[linha]['CLASSE DE VALOR']))) # Envia Classa de valor Cliente
                    enviarkey_elemento(navegador,'var_codfilialprotheus__',By.NAME,cod_filial) # Envia COD FILIAL - PADRÃO
                    enviarkey_elemento(navegador,'var_coduo__',By.NAME,cod_uo) # Envia COD UO - PADRÃO
                    enviarkey_elemento(navegador,'var_codccusto__',By.NAME,str(int(planilha.iloc[linha][f'CR{i+1}']))) # Envia COD PRODUTO
                    clicar_elemento(navegador,'searchbutton',By.ID) # Clica na Pesquisa
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                    clicar_elemento(navegador,'tooltip0',By.ID)
                    navegador.switch_to.default_content()
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                    enviarkey_java(navegador,'var_dadosCobranca__cobRateio__filialCustos__valor__',planilha.iloc[linha][f'VALOR{i+1}']/planilha.iloc[linha]['PARCELA']) # Envia Valor
                    clicar_elemento(navegador,'//*[@id="dibButtons"]/input[1]',By.XPATH) #botão ok
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                    contador += 1
                for i in range(contador):
                    clicar_elemento_rustico(navegador,f'//*[@id="{i}"]/td[2]',By.XPATH) #Para acertar as porcentagens 
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                    clicar_elemento_rustico(navegador,'//*[@id="dibButtons"]/input[1]',By.XPATH) #botão ok
                    acessar_iframe_default(navegador,tempo_espera) # Acessa Iframe da Pesquisa
                clicar_elemento(navegador,'//*[@id="dibButtons"]/input[1]',By.XPATH) #botão ok
                navegador.switch_to.default_content()

            enviarkey_elemento(navegador,'var_dadosCobranca__Observacao__',By.ID,planilha.iloc[linha]['OBSERVACAO']) # Envia Observação
            enviar_emails(navegador,linha,"//li[@onclick=\"activeDeactiveObjMenu2(this);javascript: ellist_EmailDeContatoDosClientes__.addNewItem('CreateItens', true);\"]/a[@id='createitens']",'var_EmailDeContatoDosClientes__Email__',planilha)
            input('Confirma o lançamento!!!')
            navegador.switch_to.default_content()
            clicar_elemento(navegador,'action.send',By.NAME)
            esperar_alerta(navegador,nome_cob, aba_original,planilha,local_destino,'Novo',linha,nome_cob)
            time.sleep(1)
            acessar_iframe_default(navegador,tempo_espera)

    copiar_para_planilha(planilha_destino,local_destino)