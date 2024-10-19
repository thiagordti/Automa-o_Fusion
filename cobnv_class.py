from utils import *

class CobNV:
    def __init__(self, caminho, navegador, planilha, planilha_destino, local_destino, cod_filial='01MG0014', cod_uo='10310', tempo_espera=0.5):
        self.caminho = caminho
        self.navegador = navegador
        self.planilha = planilha
        self.planilha_destino = planilha_destino
        self.local_destino = local_destino
        self.cod_filial = cod_filial
        self.cod_uo = cod_uo
        self.tempo_espera = tempo_espera
        self.linha_atual = 0
    
    def cob_nv(self):
        aba_original = self.navegador.window_handles[0] # Identifica Aba Primaria
        for linha in range(len(self.planilha)):
            clicar_elemento(self.navegador,'btnStartProcess',By.ID) # Iniciar novo processo
            clicar_elemento(self.navegador,'//*[@id="page-content-wrapper"]/div/div/div[1]/div[1]/nav/div/div/div/ul/li[3]/ul/li[5]/a/div/span[1]',By.XPATH) # Iniciar nova Cobrança
            WebDriverWait(self.navegador, 10).until(lambda d: len(d.window_handles) > 1)
            nova_aba = self.navegador.window_handles[1]# Identifica nova aba apos iniciar nova Cobrança
            self.navegador.switch_to.window(nova_aba) # Troca para nova Aba
            enviarkey_elemento(self.navegador,'id_informeNucleo__',By.ID,'Núcleo de Faturamento')# Envia nucleo - Padrão
            enviarkey_elemento(self.navegador,'id_tipoSolicitacao__',By.ID,'Solicitação de cobrança (FG-176)')# Solicitação de cobrança - Padrão
            enviarkey_elemento(self.navegador,'id_plataformaGestaoDaVenda__',By.ID,'Protheus')# Plataforma - Padrão

            if self.planilha.iloc[linha]['TIPO'].lower() == "variavel" and self.planilha.iloc[linha]['RATEIO'].lower() == "sim":
                enviarkey_elemento(self.navegador,'id_tipoDeMedicao__',By.ID,'Variavel')# Tipo de medição
                data = self.planilha.iloc[linha]['DESCRICAO']
                date = datetime.strptime(data.strftime('%d/%m/%Y'), '%d/%m/%Y') # Transforma data em string
                primeiro_dia, ultimo_dia = primeiro_e_ultimo_dia_do_mes(date.year, date.month) # Pega o mês e dia
                data_str = self.planilha.iloc[linha]['DATA']
                data_obj = datetime.strptime(data_str.strftime('%d/%m/%Y'), '%d/%m/%Y')
                nome_cob = texto_elemento(self.navegador,'headerTitle',By.ID)
                variavel_novo(self.navegador,linha,self.planilha,primeiro_dia,ultimo_dia) # Carrega dados de preenchimento
                if pd.isna(self.planilha.iloc[linha][f'CR1']): # Loop para verificar se o Item está vazio!!
                    pass
                else:
                    clicar_elemento(self.navegador,'//*[@id="createitem"]',By.XPATH) # Clica no Novo Item
                    acessar_iframe(self.navegador,self.tempo_espera)# Acessa o Iframe
                    enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(self.planilha.iloc[linha]['CNPJ'])) # Envia CNPJ
                    clicar_elemento_dinamico(self.navegador) # Clica no CNPJ informado
                    enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA SESI VIVA+: AEP,PGR,PCMSO,LTCAT \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
                    enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Não')# Envia não ao campo de rateio
                    clicar_elemento(self.navegador,'id_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDeCobranca__UOCRProtheus___anchor',By.ID) # Clica na pesquisa
                    acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe da Pesquisa
                    clicar_elemento(self.navegador,'vfilter',By.ID) # Clica no Filtro
                    acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe do Filtro
                    enviarkey_elemento(self.navegador,'var_codclvlr__',By.NAME,str(int(self.planilha.iloc[linha]['CLASSE DE VALOR']))) # Envia Classa de valor Cliente
                    enviarkey_elemento(self.navegador,'var_codfilialprotheus__',By.NAME,self.cod_filial) # Envia COD FILIAL - PADRÃO
                    enviarkey_elemento(self.navegador,'var_coduo__',By.NAME,self.cod_uo) # Envia COD UO - PADRÃO
                    enviarkey_elemento(self.navegador,'var_codccusto__',By.NAME,int(self.planilha.iloc[linha]['CR1'])) # Envia COD PRODUTO
                    clicar_elemento(self.navegador,'searchbutton',By.ID) # Clica na Pesquisa
                    acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe da Pesquisa
                    clicar_elemento(self.navegador,'tooltip0',By.ID)
                    acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe primario
                    clicar_elemento(self.navegador,'createitem',By.ID) # Clica para adicionar Valor
                    acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe de valor
                    opcoes_pagamento(self.navegador,'//*[@id="mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos_ori"]/option[1]','move_this_right_mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos')#Loop para selecionar as opções de pagamento     
                    enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__data__',By.ID,data_obj.strftime('%d/%m/%Y')) # Envia data da cobrança
                    enviarkey_java(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__valor__',self.planilha.iloc[linha]['VALOR1'])
                    clicar_elemento(self.navegador,'action.save',By.NAME) # Clica para salvar.
                    acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe primario
                enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__numeroContratoProtheus__',By.ID,int(self.planilha.iloc[linha]['NUMERO DO CONTRATO'])) # Envia o numero de contrato
                clicar_elemento(self.navegador,'//*[@id="ui-id-10"]/li',By.XPATH) # Clica no numero de contrato
                clicar_elemento(self.navegador,'//*[@id="dibButtons"]/input[1]',By.XPATH) # Clica no numero de contrato
                self.navegador.switch_to.default_content()#Volta para o inicio
                #------------------------Rateio--------------------------------------------------
                clicar_elemento(self.navegador,'//*[@id="createitem"]',By.XPATH) # Clica no Novo
                acessar_iframe(self.navegador,self.tempo_espera)# Acessa o Iframe
                enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(self.planilha.iloc[linha]['CNPJ'])) # Envia CNPJ
                clicar_elemento_dinamico(self.navegador) # Clica no CNPJ informado
                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'SIM')# Envia Sim ao campo de rateio

                contador = 0 # Contador utilizado para clicar nos rateios no processo Final!
                #Loop para a quantidade de Itens
                for com_rateio in range(5): # Loop para verificar todos os itens (Total 5) com rateio na self.planilha!!
                    if pd.isna(self.planilha.iloc[linha][f'CR{com_rateio+2}']): # Loop para verificar se o Item está vazio!!
                        pass # Pula o item vazio
                    else:
                        dados_rateio(self.navegador,linha,self.cod_filial,self.cod_uo,self.planilha, self.tempo_espera)
                        enviarkey_elemento(self.navegador,'var_codccusto__',By.NAME,str(int(self.planilha.iloc[linha][f'CR{com_rateio+2}']))) # Envia COD PRODUTO
                        clicar_elemento(self.navegador,'searchbutton',By.ID) # Clica na Pesquisa
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe da Pesquisa
                        clicar_elemento(self.navegador,'tooltip0',By.ID) # Clica no item filtrado
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe primario3
                        elemento2 = WebDriverWait(self.navegador, 10).until(EC.presence_of_element_located((By.NAME, 'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__')))
                        script_valor_cr = f"document.getElementsByName('var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__')[0].value='{self.planilha.iloc[linha][f'VALOR{com_rateio+2}']}';"
                        self.navegador.execute_script(script_valor_cr)
                        clicar_elemento(self.navegador,'action.save',By.NAME) # Clica para salvar.
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe primario
                        contador += 1 # Soma 1 a quantidade de contador, será utiizado para clicar no loop Contador!

                for i in range(contador): # Baseado na soma do Contador clica nos itens
                    clicar_elemento_rustico(self.navegador,f'//*[@id="{i}"]/td[2]',By.XPATH) # Clica no Item baseado nos indices (No fusion o indice 0 conta!)!!
                    acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe primario
                    clicar_elemento(self.navegador,'action.save',By.NAME) # Clica para salvar.
                    acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe primario

                enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__numeroContratoProtheus__',By.ID,str(int(self.planilha.iloc[linha]['NUMERO DO CONTRATO']))) # Envia o numero de contrato
                clicar_elemento(self.navegador,'//*[@id="ui-id-10"]/li',By.XPATH) # Clica no numero de contrato
                clicar_elemento(self.navegador,'action.save',By.NAME) # Clica para salvar.
                self.navegador.switch_to.default_content()#Volta para o inicio
                
                enviar_anexo(self.navegador,linha,'//*[@id="menu_bar_genericoHistoricoAtendimento"]/li[1]','var_dadosDaCobranca__historico__anexo__','//*[@id="progress-complete-var_dadosDaCobranca__historico__anexo__"]/span','var_dadosDaCobranca__historico__registro__',self.planilha,self.caminho,self.tempo_espera)#Envia Anexos
                enviar_emails(self.navegador,linha,"//li[@onclick=\"activeDeactiveObjMenu2(this);javascript: ellist_emailClienteFP__.addNewItem('CreateItens', true);\"]//a[@id='createitens']",'var_emailClienteFP__Email__',self.planilha, self.tempo_espera) # Envia e-mails
                input('Confirma o lançamento!!!')
                self.navegador.switch_to.default_content()
                clicar_elemento(self.navegador,'action.send',By.NAME)
                esperar_alerta(self.navegador,nome_cob, aba_original,self.planilha,self.local_destino,'Novo',linha,nome_cob)
                time.sleep(1)
                acessar_iframe_default(self.navegador,self.tempo_espera)

            elif self.planilha.iloc[linha]['TIPO'].lower() == "variavel" and self.planilha.iloc[linha]['RATEIO'].lower() == "não":
                enviarkey_elemento(self.navegador,'id_tipoDeMedicao__',By.ID,'Variavel')# Tipo de medição
                data = self.planilha.iloc[linha]['DESCRICAO']
                date = datetime.strptime(data.strftime('%d/%m/%Y'), '%d/%m/%Y') # Transforma data em string
                primeiro_dia, ultimo_dia = primeiro_e_ultimo_dia_do_mes(date.year, date.month) # Pega o mês e dia
                data_str = self.planilha.iloc[linha]['DATA']
                data_obj = datetime.strptime(data_str.strftime('%d/%m/%Y'), '%d/%m/%Y')
                nome_cob = texto_elemento(self.navegador,'headerTitle',By.ID)
                variavel_novo(self.navegador,linha,self.planilha,primeiro_dia,ultimo_dia) # Carrega dados de preenchimento
                clicar_elemento(self.navegador,'//*[@id="createitem"]',By.XPATH) #  Clica no Novo Item
                acessar_iframe(self.navegador,self.tempo_espera)# Acessa o Iframe
                enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(self.planilha.iloc[linha]['CNPJ'])) # Envia CNPJ
                clicar_elemento_dinamico(self.navegador) # Clica no CNPJ informado
                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA SESI VIVA+: AEP,PGR,PCMSO,LTCAT \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Não')# Envia não ao campo de rateio
                for sem_rateio in range(2):
                    if pd.isna(self.planilha.iloc[linha][f'CR{sem_rateio+1}']):
                        pass
                    else:
                        clicar_elemento(self.navegador,'id_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDeCobranca__UOCRProtheus___anchor',By.ID)# Clica novo
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe da Pesquisa
                        clicar_elemento(self.navegador,'vfilter',By.ID) # Clica no Filtro
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe do Filtro
                        enviarkey_elemento(self.navegador,'var_codclvlr__',By.NAME,str(int(self.planilha.iloc[linha]['CLASSE DE VALOR']))) # Envia Classa de valor Cliente
                        enviarkey_elemento(self.navegador,'var_codfilialprotheus__',By.NAME,self.cod_filial) # Envia COD FILIAL - PADRÃO
                        enviarkey_elemento(self.navegador,'var_coduo__',By.NAME,self.cod_uo) # Envia COD UO - PADRÃO
                        enviarkey_elemento(self.navegador,'var_codccusto__',By.NAME,int(self.planilha.iloc[linha][f'CR{sem_rateio+1}'])) # Envia COD PRODUTO
                        clicar_elemento(self.navegador,'searchbutton',By.ID) # Clica na Pesquisa
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe da Pesquisa
                        clicar_elemento(self.navegador,'tooltip0',By.ID)
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe primario
                        clicar_elemento(self.navegador,'createitem',By.ID) # Clica para adicionar Valor
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe de valor
                        opcoes_pagamento(self.navegador,'//*[@id="mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos_ori"]/option[1]','move_this_right_mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos')#Loop para selecionar as opções de pagamento 
                        enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__data__',By.ID,data_obj.strftime('%d/%m/%Y')) # Envia data da cobrança
                        enviarkey_java(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__valor__',self.planilha.iloc[linha][f'VALOR{sem_rateio+1}'])
                        clicar_elemento(self.navegador,'action.save',By.NAME) # Clica para salvar.
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe primario
                enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__numeroContratoProtheus__',By.ID,int(self.planilha.iloc[linha]['NUMERO DO CONTRATO'])) # Envia o numero de contrato
                clicar_elemento(self.navegador,'//*[@id="ui-id-10"]/li',By.XPATH) # Clica no numero de contrato
                clicar_elemento(self.navegador,'//*[@id="dibButtons"]/input[1]',By.XPATH) # Clica no numero de contrato
                self.navegador.switch_to.default_content()#Volta para o inicio

                # ---------------------- Esta Parte se refere aos Anexos ------------------------
                enviar_anexo(self.navegador,linha,'//*[@id="menu_bar_genericoHistoricoAtendimento"]/li[1]','var_dadosDaCobranca__historico__anexo__','//*[@id="progress-complete-var_dadosDaCobranca__historico__anexo__"]/span','var_dadosDaCobranca__historico__registro__',self.planilha,self.caminho,self.tempo_espera)#Envia Anexos
                enviar_emails(self.navegador,linha,"//li[@onclick=\"activeDeactiveObjMenu2(this);javascript: ellist_emailClienteFP__.addNewItem('CreateItens', true);\"]//a[@id='createitens']",'var_emailClienteFP__Email__',self.planilha, self.tempo_espera)
                input('Confirma o lançamento!!!')
                self.navegador.switch_to.default_content()
                clicar_elemento(self.navegador,'action.send',By.NAME)
                esperar_alerta(self.navegador,nome_cob, aba_original,self.planilha,self.local_destino,'Novo',linha,nome_cob)
                time.sleep(1)
                acessar_iframe_default(self.navegador,self.tempo_espera)

            elif self.planilha.iloc[linha]['TIPO'].lower() == "fixo" and self.planilha.iloc[linha]['RATEIO'].lower() == "não":
                enviarkey_elemento(self.navegador,'id_tipoDeMedicao__',By.ID,'Fixa')# Tipo de medição
                clicar_elemento(self.navegador,'//*[@id="tab_bar_"]/li[2]/a',By.XPATH) # Clica na guia Dados de cobrança
                nome_cob = texto_elemento(self.navegador,'headerTitle',By.ID)
                enviarkey_elemento(self.navegador,'id_txt_dadosCobranca__DadosDoCliente__',By.ID,tratar_cnpj(self.planilha.iloc[linha]['CNPJ']))# Envia CNPJ
                clicar_elemento(self.navegador,'ui-menu-item',By.CLASS_NAME) # Clica no CNPJ informado
                enviarkey_elemento(self.navegador,'var_dadosCobranca__descricaoServico__',By.ID,self.planilha.iloc[linha]['DESCRICAO'])# Envia Descrição
                enviarkey_elemento(self.navegador,'var_dadosCobranca__dadosParaHistorico2__HouvePrestacaoDeServicos__',By.ID,'Sim')# Prestação de Serviço - Padrão
                enviarkey_elemento(self.navegador,'id_txt_dadosCobranca__dadosParaHistorico2__numeroContratoProtheus__',By.ID,str(int(self.planilha.iloc[linha]['NUMERO DO CONTRATO'])))# Numero de contrato
                clicar_elemento(self.navegador,'//*[@id="ui-id-6"]/li',By.XPATH) # Clica no contrato informado
                enviarkey_elemento(self.navegador,'id_dadosCobranca__dadosParaHistorico2__diaLimiteNFCliente__',By.ID,str(int(self.planilha.iloc[linha]['DIA LIMITE'])))# Data Limite
                opcoes_pagamento(self.navegador,'//*[@id="mul_dadosCobranca__formaDeEntradaDosRecursosRevisado_ori"]/option[1]','move_this_right_mul_dadosCobranca__formaDeEntradaDosRecursosRevisado') # Loop para opções de pagamento
                enviarkey_elemento(self.navegador,'//*[@id="var_dadosCobranca__rateio__"]',By.XPATH,'Não')# Tipo de Rateio

                clicar_elemento(self.navegador,'id_dadosCobranca__filaisSemRateio__UOCRProtheus___anchor',By.ID) # Pequisa
                acessar_iframe_default(self.navegador,self.tempo_espera)# Acessa o Iframe

                clicar_elemento(self.navegador,'vfilter',By.ID) # Clica no Filtro
                acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe do Filtro
                enviarkey_elemento(self.navegador,'var_codclvlr__',By.NAME,str(int(self.planilha.iloc[linha]['CLASSE DE VALOR']))) # Envia Classa de valor Cliente
                enviarkey_elemento(self.navegador,'var_codfilialprotheus__',By.NAME,self.cod_filial) # Envia COD FILIAL - PADRÃO
                enviarkey_elemento(self.navegador,'var_coduo__',By.NAME,self.cod_uo) # Envia COD UO - PADRÃO
                enviarkey_elemento(self.navegador,'var_codccusto__',By.NAME,str(int(self.planilha.iloc[linha]['CR1']))) # Envia COD PRODUTO
                clicar_elemento(self.navegador,'searchbutton',By.ID) # Clica na Pesquisa
                acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe da Pesquisa
                clicar_elemento(self.navegador,'tooltip0',By.ID)
                self.navegador.switch_to.default_content()

                valor = float(self.planilha.iloc[linha]['VALOR1'])
                parcelas = int(self.planilha.iloc[linha]['PARCELA'])
                valor_parcela = valor / parcelas
                date = self.planilha.iloc[linha]['DATA']

                clicar_elemento(self.navegador,'//*[@id="menu_bar_FinCobFFDatasVencimentos"]/li[2]',By.XPATH) # Itens novos

                for i in range(int(self.planilha.iloc[linha]['PARCELA'])):
                        new_data = date + relativedelta(months=i)
                        new_date_str = new_data.strftime('%d/%m/%Y')
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe dos itens novos
                        enviarkey_elemento(self.navegador,'var_dadosCobranca__cobrancas__data__',By.ID,new_date_str) # Envia Data
                        enviarkey_java(self.navegador,'var_dadosCobranca__cobrancas__valor__',valor_parcela) # Envia Valor
                        clicar_elemento(self.navegador,'//*[@id="dibButtons"]/input[1]',By.XPATH) # Botão Ok
                acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe dos itens novos
                clicar_elemento(self.navegador,'cancelButtonModal',By.ID) # Botão Cancelar
                self.navegador.switch_to.default_content()
                enviarkey_elemento(self.navegador,'var_dadosCobranca__Observacao__',By.ID,self.planilha.iloc[linha]['OBSERVACAO']) # Envia Observação
                enviar_emails(self.navegador,linha,"//li[@onclick=\"activeDeactiveObjMenu2(this);javascript: ellist_EmailDeContatoDosClientes__.addNewItem('CreateItens', true);\"]/a[@id='createitens']",'var_EmailDeContatoDosClientes__Email__',self.planilha, self.tempo_espera)
                input('Confirma o lançamento!!!')
                self.navegador.switch_to.default_content()
                clicar_elemento(self.navegador,'action.send',By.NAME)
                esperar_alerta(self.navegador,nome_cob, aba_original,self.planilha,self.local_destino,'Novo',linha,nome_cob)
                time.sleep(1)
                acessar_iframe_default(self.navegador,self.tempo_espera)

            elif self.planilha.iloc[linha]['TIPO'].lower() == "fixo" and self.planilha.iloc[linha]['RATEIO'].lower() == "sim":
                enviarkey_elemento(self.navegador,'id_tipoDeMedicao__',By.ID,'Fixa')# Tipo de medição
                clicar_elemento(self.navegador,'//*[@id="tab_bar_"]/li[2]/a',By.XPATH) # Clica na guia Dados de cobrança
                nome_cob = texto_elemento(self.navegador,'headerTitle',By.ID)
                enviarkey_elemento(self.navegador,'id_txt_dadosCobranca__DadosDoCliente__',By.ID,tratar_cnpj(self.planilha.iloc[linha]['CNPJ']))# Envia CNPJ
                clicar_elemento(self.navegador,'ui-menu-item',By.CLASS_NAME) # Clica no CNPJ informado
                enviarkey_elemento(self.navegador,'var_dadosCobranca__descricaoServico__',By.ID,self.planilha.iloc[linha]['DESCRICAO'])# Envia Descrição
                enviarkey_elemento(self.navegador,'var_dadosCobranca__dadosParaHistorico2__HouvePrestacaoDeServicos__',By.ID,'Sim')# Prestação de Serviço - Padrão
                enviarkey_elemento(self.navegador,'id_txt_dadosCobranca__dadosParaHistorico2__numeroContratoProtheus__',By.ID,str(int(self.planilha.iloc[linha]['NUMERO DO CONTRATO'])))# Numero de contrato
                clicar_elemento(self.navegador,'//*[@id="ui-id-6"]/li',By.XPATH) # Clica no contrato informado
                enviarkey_elemento(self.navegador,'id_dadosCobranca__dadosParaHistorico2__diaLimiteNFCliente__',By.ID,str(int(self.planilha.iloc[linha]['DIA LIMITE'])))# Data Limite
                opcoes_pagamento(self.navegador,'//*[@id="mul_dadosCobranca__formaDeEntradaDosRecursosRevisado_ori"]/option[1]','move_this_right_mul_dadosCobranca__formaDeEntradaDosRecursosRevisado')# Loop para opções de pagamento
                enviarkey_elemento(self.navegador,'//*[@id="var_dadosCobranca__rateio__"]',By.XPATH,'Sim')# Tipo de Rateio
                parcemlamento = 0
                for i in range(6):
                    if not pd.isna(self.planilha.iloc[linha][f'CR{i+1}']):
                        parcemlamento += 1
                date = self.planilha.iloc[0]['DATA']
                for i in range(int(self.planilha.iloc[linha]['PARCELA'])):
                    clicar_elemento(self.navegador,'//*[@id="menu_bar_finCobDataXFilialXCcusto"]/li[1]',By.XPATH) # Clica produtos novo
                    acessar_iframe_default(self.navegador,self.tempo_espera)# Acessa o Iframe
                    new_data = date + relativedelta(months=i)
                    new_date_str = new_data.strftime('%d/%m/%Y')
                    enviarkey_elemento(self.navegador,'var_dadosCobranca__cobRateio__dataCobranca__',By.ID,new_date_str) # Envia Data
                    contador = 0
                    for i in range(parcemlamento):
                        clicar_elemento(self.navegador,'//*[@id="menu_bar_finCobFFCentroValor"]/li[1]',By.XPATH) # Pequisa
                        acessar_iframe_default(self.navegador,self.tempo_espera)# Acessa o Iframe
                        enviarkey_elemento(self.navegador,'id_txt_dadosCobranca__cobRateio__filialCustos__FilialProtheus__',By.ID,self.cod_filial) # Envia COD FILIAL - PADRÃO
                        clicar_elemento(self.navegador,'ui-menu-item',By.CLASS_NAME)
                        enviarkey_elemento(self.navegador,'id_txt_dadosCobranca__cobRateio__filialCustos__UO__',By.ID,self.cod_uo) # Envia COD FILIAL - PADRÃO
                        clicar_elemento_dinamico(self.navegador)
                        clicar_elemento(self.navegador,'id_dadosCobranca__cobRateio__filialCustos__UOXCRProtheus___anchor',By.ID)

                        acessar_iframe_default(self.navegador,self.tempo_espera)# Acessa o Iframe
                        clicar_elemento(self.navegador,'vfilter',By.ID) # Clica no Filtro
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe do Filtro
                        enviarkey_elemento(self.navegador,'var_codclvlr__',By.NAME,str(int(self.planilha.iloc[linha]['CLASSE DE VALOR']))) # Envia Classa de valor Cliente
                        enviarkey_elemento(self.navegador,'var_codfilialprotheus__',By.NAME,self.cod_filial) # Envia COD FILIAL - PADRÃO
                        enviarkey_elemento(self.navegador,'var_coduo__',By.NAME,self.cod_uo) # Envia COD UO - PADRÃO
                        enviarkey_elemento(self.navegador,'var_codccusto__',By.NAME,str(int(self.planilha.iloc[linha][f'CR{i+1}']))) # Envia COD PRODUTO
                        clicar_elemento(self.navegador,'searchbutton',By.ID) # Clica na Pesquisa
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe da Pesquisa
                        clicar_elemento(self.navegador,'tooltip0',By.ID)
                        self.navegador.switch_to.default_content()
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe da Pesquisa
                        enviarkey_java(self.navegador,'var_dadosCobranca__cobRateio__filialCustos__valor__',self.planilha.iloc[linha][f'VALOR{i+1}']/self.planilha.iloc[linha]['PARCELA']) # Envia Valor
                        clicar_elemento(self.navegador,'//*[@id="dibButtons"]/input[1]',By.XPATH) #botão ok
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe da Pesquisa
                        contador += 1
                    for i in range(contador):
                        clicar_elemento_rustico(self.navegador,f'//*[@id="{i}"]/td[2]',By.XPATH) #Para acertar as porcentagens 
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe da Pesquisa
                        clicar_elemento_rustico(self.navegador,'//*[@id="dibButtons"]/input[1]',By.XPATH) #botão ok
                        acessar_iframe_default(self.navegador,self.tempo_espera) # Acessa Iframe da Pesquisa
                    clicar_elemento(self.navegador,'//*[@id="dibButtons"]/input[1]',By.XPATH) #botão ok
                    self.navegador.switch_to.default_content()

                enviarkey_elemento(self.navegador,'var_dadosCobranca__Observacao__',By.ID,self.planilha.iloc[linha]['OBSERVACAO']) # Envia Observação
                enviar_emails(self.navegador,linha,"//li[@onclick=\"activeDeactiveObjMenu2(this);javascript: ellist_EmailDeContatoDosClientes__.addNewItem('CreateItens', true);\"]/a[@id='createitens']",'var_EmailDeContatoDosClientes__Email__',self.planilha, self.tempo_espera)
                input('Confirma o lançamento!!!')
                self.navegador.switch_to.default_content()
                clicar_elemento(self.navegador,'action.send',By.NAME)
                esperar_alerta(self.navegador,nome_cob, aba_original,self.planilha,self.local_destino,'Novo',linha,nome_cob)
                time.sleep(1)
                acessar_iframe_default(self.navegador,self.tempo_espera)

        copiar_para_planilha(self.planilha_destino,self.local_destino)
