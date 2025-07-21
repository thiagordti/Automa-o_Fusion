from utils import *
import sys
import tkinter as tk
from tkinter import messagebox

class AutomacaoFusion:

    def __init__(self, caminho, navegador, chrome_proc, planilha, planilha_destino, local_destino, metodo, cod_filial='01MG0014', cod_uo='10310'):
        self.caminho = caminho
        self.navegador = navegador
        self.chrome_proc = chrome_proc
        self.planilha = planilha
        self.planilha_destino = planilha_destino
        self.local_destino = local_destino
        self.cod_filial = cod_filial
        self.cod_uo = cod_uo
        self.tempo_espera = 0.5 # Tempo de espera para carregar os elementos
        self.linha_atual = 0
        self.metodo = metodo
    
    def medicao_vr(self):
        for linha in range(self.linha_atual, len(self.planilha)):
            self.linha_atual = linha  # Atualiza a linha atual
            acessar_iframe(self.navegador,self.tempo_espera, self)# Acessa o Iframe
            enviarkey_elemento(self.navegador,'searchBarProcessQuery',By.ID,self.planilha.iloc[linha]['COB'], self)#Envio do COB
            esperar_elementos_carregar(self.navegador)
            clicar_elemento_rustico(self.navegador,'//*[@id="page-content-wrapper"]/div/div/div[1]/div[1]/nav/div/form/div/div/span/button',By.XPATH, self) # Clica no botão de pesquisa inicial
            aba_original = self.navegador.window_handles[0] # Identifica Aba Primaria
            clicar_elemento_rustico(self.navegador, 'header', By.CLASS_NAME, self) # Clica no COB pesquisado
            WebDriverWait(self.navegador, 10).until(lambda d: len(d.window_handles) > 1)
            nova_aba = self.navegador.window_handles[1]# Identifica nova aba apos iniciar Cobrança
            self.navegador.switch_to.window(nova_aba) # Troca para nova Aba
            nome_cob = texto_elemento(self.navegador,'headerTitle',By.ID)
            data = self.planilha.iloc[linha]['DATA_DESCRIÇÃO'] # Pega data de Descrição
            date = datetime.strptime(data.strftime('%d/%m/%Y'), '%d/%m/%Y') # Transforma data em string
            primeiro_dia, ultimo_dia = primeiro_e_ultimo_dia_do_mes(date.year, date.month) # Pega o mês e dia
            if pd.isna(self.planilha.iloc[linha]['DATA_DE_VENCIMENTO']):
                data_venc = date
            else:
                data_venc = self.planilha.iloc[linha]['DATA_DE_VENCIMENTO'] # Pega data de Vencimento
            # ---------------------- Esta Parte se refere ao COB sem Rateio ------------------------
            for sem_rateio in range(2):
                if pd.isna(self.planilha.iloc[linha][f'CR-SR{sem_rateio+1}']): # Verifica se o campo está vazio
                    pass
                else:
                    clicar_elemento(self.navegador,'createitem',By.ID, self)# Clica para criar novo Item
                    acessar_iframe(self.navegador,self.tempo_espera, self)# Acessa o Iframe
                    enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(self.planilha.iloc[linha]['CNPJ']), self) # Envia CNPJ
                    clicar_elemento_dinamico(self.navegador, self) # Clica no CNPJ informado
                    enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__numeroContratoProtheus__',By.ID,str(int(self.planilha.iloc[linha]['NUMERO DO CONTRATO'])), self) # Envia Numero de Contrato
                    clicar_elemento(self.navegador,'//*[@id="ui-id-10"]/li',By.XPATH,self) # Clica no numero de contrato
                    if sem_rateio == 0: # Difere o primeiro produto do segundo
                        if not pd.isna(self.planilha.iloc[linha]['TEXTO1']): # Verifica se o campo TEXTO1 é maior que 3, se sim Envia o TEXTO1
                            enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,self.planilha.iloc[linha]['TEXTO1'], self) # Envia Descrição da coluna TEXTO1
                        else: #Se não envia a descrição padrão
                            enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA SESI VIVA+: AEP,PGR,PCMSO,LTCAT \nPERÍODO: {primeiro_dia} a {ultimo_dia}.', self) # Envia Descrição Padrão
                    else:
                        if not pd.isna(self.planilha.iloc[linha]['TEXTO2']): # Verifica se o campo TEXTO2 é maior que 3, se sim Envia o TEXTO2
                            enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,self.planilha.iloc[linha]['TEXTO2'], self) # Envia Descrição da coluna TEXTO2
                        else:#Se não envia a descrição padrão
                            enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.', self) # Envia Descrição Padrão
                    enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Não', self)# Envia não ao campo de rateio
                    clicar_elemento(self.navegador,'id_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDeCobranca__UOCRProtheus___anchor',By.ID, self)# Clica na pesquisa de produto
                    acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe da Pesquisa
                    clicar_elemento(self.navegador,'vfilter',By.ID, self) # Clica no Filtro
                    acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe do Filtro
                    enviarkey_elemento(self.navegador,'var_codclvlr__',By.NAME,str(int(self.planilha.iloc[linha]['CLASSE DE VALOR'])), self) # Envia Classe de valor Cliente
                    enviarkey_elemento(self.navegador,'var_codfilialprotheus__',By.NAME,self.cod_filial, self) # Envia COD FILIAL - PADRÃO
                    enviarkey_elemento(self.navegador,'var_coduo__',By.NAME,self.cod_uo, self) # Envia COD UO - PADRÃO
                    enviarkey_elemento(self.navegador,'var_codccusto__',By.NAME,str(int(self.planilha.iloc[linha][f'CR-SR{sem_rateio+1}'])), self) # Envia COD PRODUTO
                    clicar_elemento(self.navegador,'searchbutton',By.ID, self) # Clica na Pesquisa
                    acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe da Pesquisa
                    clicar_elemento(self.navegador,'tooltip0',By.ID, self)
                    acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario
                    clicar_elemento(self.navegador,'createitem',By.ID, self) # Clica para adicionar Valor
                    acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe de valor
                    opcoes_pagamento(self.navegador,'//*[@id="mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos_ori"]/option[1]','move_this_right_mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos',self)#Loop para selecionar as opções de pagamento
                    enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__data__',By.NAME,data_venc.strftime('%d/%m/%Y'), self) # Envia data da cobrança
                    enviarkey_java(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__valor__',self.planilha.iloc[linha][f'VALORSR{sem_rateio+1}'], self) # Envia Valor
                    clicar_elemento(self.navegador,'action.save',By.NAME, self) # Clica para salvar.
                    acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario
                    clicar_elemento(self.navegador,'action.save',By.NAME, self) # Clica para salvar.
                    self.navegador.switch_to.default_content()#Volta para o inicio

            # ---------------------- Esta Parte se refere ao COB com Rateio ------------------------

            contador = 0 # Contador utilizado para clicar nos rateios no processo Final!
            contador_1 = 0 # Contador utilizado para clicar nos rateios no processo Final!
            contador_2 = 0 # Contador utilizado para clicar nos rateios no processo Final!
            if pd.isna(self.planilha.iloc[linha]['CRR1']): # Verifica se o primeiro item está vazio, se o mesmo estiver vazio, todo o loop é pulado!
                pass

            elif pd.isna(self.planilha.iloc[linha]['QTD RATEIO']) or int(self.planilha.iloc[linha]['QTD RATEIO']) == 1 : # Caso não esteja vazio é iniciado o processo de Rateio e a QTD seja um executa todos os rateios em um unico processo
                clicar_elemento(self.navegador,'createitem',By.ID, self)# Clica para criar novo Item
                acessar_iframe(self.navegador,self.tempo_espera, self)# Acessa o Iframe
                enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(self.planilha.iloc[linha]['CNPJ']), self) # Envia CNPJ
                clicar_elemento_dinamico(self.navegador, self) # Clica no CNPJ informado
                enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__numeroContratoProtheus__',By.ID,str(int(self.planilha.iloc[linha]['NUMERO DO CONTRATO'])), self) # Envia Numero de Contrato
                clicar_elemento(self.navegador,'//*[@id="ui-id-10"]/li',By.XPATH,self) # Clica no numero de contrato
                if not pd.isna(self.planilha.iloc[linha]['TEXTO3']): # Verifica se o campo TEXTO3 é maior que 3, se sim Envia o TEXTO3
                            enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,self.planilha.iloc[linha]['TEXTO3'], self) # Envia Descrição da coluna TEXTO3
                else:#Se não envia a descrição padrão
                    enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.', self) # Envia Descrição
                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Sim', self)# Envia sim ao campo de rateio
                #Loop para a quantidade de Itens
                for com_rateio in range(4): # Loop para verificar todos os itens (Total 4) com rateio na self.planilha!!
                    if pd.isna(self.planilha.iloc[linha][f'CRR{com_rateio+1}']): # Loop para verificar se o Item está vazio!!
                        pass # Pula o item vazio
                    else:
                        dados_rateio(self.navegador,linha,self.cod_filial,self.cod_uo,self.planilha, self.tempo_espera, self)
                        enviarkey_elemento(self.navegador,'var_codccusto__',By.NAME,str(int(self.planilha.iloc[linha][f'CRR{com_rateio+1}'])), self) # Envia COD PRODUTO
                        clicar_elemento(self.navegador,'searchbutton',By.ID, self) # Clica na Pesquisa
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe da Pesquisa
                        clicar_elemento(self.navegador,'tooltip0',By.ID, self) # Clica no item filtrado
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario3
                        enviarkey_java(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__',self.planilha.iloc[linha][f'VALOR{com_rateio+1}'], self) # Envia Valor
                        clicar_elemento(self.navegador,'action.save',By.NAME, self) # Clica para salvar.
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario
                        contador += 1 # Soma 1 a quantidade de contador, será utiizado para clicar no loop Contador!
                clicar_porcentagem(self.navegador,contador,self.tempo_espera, self) # Baseado na soma do Contador clica nos itens

            elif int(self.planilha.iloc[linha]['QTD RATEIO']) == 2: # Ira rodar o processo de sem rateio duas vezes uma para a coluna CRR1 e 2 e ou para CRR3 e 4
                # Processo para coluna 1 e 2
                clicar_elemento(self.navegador,'createitem',By.ID, self)# Clica para criar novo Item
                acessar_iframe(self.navegador,self.tempo_espera, self)# Acessa o Iframe
                enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(self.planilha.iloc[linha]['CNPJ']), self) # Envia CNPJ
                clicar_elemento_dinamico(self.navegador, self) # Clica no CNPJ informado
                if not pd.isna(self.planilha.iloc[linha]['TEXTO3']): # Verifica se o campo TEXTO3 é maior que 3, se sim Envia o TEXTO3
                            enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,self.planilha.iloc[linha]['TEXTO3'], self) # Envia Descrição da coluna TEXTO3
                else:#Se não envia a descrição padrão
                    enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.', self) # Envia Descrição
                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Sim', self)# Envia sim ao campo de rateio
                #Loop para a quantidade de Itens
                for com_rateio in range(2): # Loop para verificar todos os itens (Total 4) com rateio na self.planilha!!
                    if pd.isna(self.planilha.iloc[linha][f'CRR{com_rateio+1}']): # Loop para verificar se o Item está vazio!!
                        pass # Pula o item vazio
                    else:
                        dados_rateio(self.navegador,linha,self.cod_filial,self.cod_uo,self.planilha, self.tempo_espera, self)
                        enviarkey_elemento(self.navegador,'var_codccusto__',By.NAME,str(int(self.planilha.iloc[linha][f'CRR{com_rateio+1}'])), self) # Envia COD PRODUTO
                        clicar_elemento(self.navegador,'searchbutton',By.ID, self) # Clica na Pesquisa
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe da Pesquisa
                        clicar_elemento(self.navegador,'tooltip0',By.ID, self) # Clica no item filtrado
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario3
                        enviarkey_java(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__',self.planilha.iloc[linha][f'VALOR{com_rateio+1}'], self) # Envia Valor
                        clicar_elemento(self.navegador,'action.save',By.NAME, self) # Clica para salvar.
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario
                        contador_1 += 1 # Soma 1 a quantidade de contador, será utiizado para clicar no loop Contador!
                clicar_porcentagem(self.navegador,contador_1,self.tempo_espera, self) # Baseado na soma do Contador clica nos itens

                # Processo para coluna 3 e 4
                clicar_elemento(self.navegador,'createitem',By.ID, self)# Clica para criar novo Item
                acessar_iframe(self.navegador,self.tempo_espera, self)# Acessa o Iframe
                enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(self.planilha.iloc[linha]['CNPJ']), self) # Envia CNPJ
                clicar_elemento_dinamico(self.navegador, self) # Clica no CNPJ informado
                if not pd.isna(self.planilha.iloc[linha]['TEXTO4']): # Verifica se o campo TEXTO4 é maior que 4, se sim Envia o TEXTO4
                    enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,self.planilha.iloc[linha]['TEXTO4'], self) # Envia Descrição da coluna TEXTO4
                else:#Se não envia a descrição padrão
                    enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'IN LOCO COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.', self) # Envia Descrição
                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Sim', self)# Envia sim ao campo de rateio
                #Loop para a quantidade de Itens
                for com_rateio in range(2): # Loop para verificar todos os itens (Total 4) com rateio na self.planilha!!
                    if pd.isna(self.planilha.iloc[linha][f'CRR{com_rateio+3}']): # Loop para verificar se o Item está vazio!!
                        pass # Pula o item vazio
                    else:
                        dados_rateio(self.navegador,linha,self.cod_filial,self.cod_uo,self.planilha, self.tempo_espera, self)
                        enviarkey_elemento(self.navegador,'var_codccusto__',By.NAME,str(int(self.planilha.iloc[linha][f'CRR{com_rateio+3}'])), self) # Envia COD PRODUTO
                        clicar_elemento(self.navegador,'searchbutton',By.ID, self) # Clica na Pesquisa
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe da Pesquisa
                        clicar_elemento(self.navegador,'tooltip0',By.ID, self) # Clica no item filtrado
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario3
                        enviarkey_java(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__',self.planilha.iloc[linha][f'VALOR{com_rateio+3}'], self) # Envia Valor
                        clicar_elemento(self.navegador,'action.save',By.NAME, self) # Clica para salvar.
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario
                        contador_2 += 1 # Soma 1 a quantidade de contador, será utiizado para clicar no loop Contador!
                clicar_porcentagem(self.navegador,contador_2,self.tempo_espera, self) # Baseado na soma do Contador clica nos itens

            # ---------------------- Esta Parte se refere aos Anexos ------------------------
            enviar_anexo(self.navegador,linha,'//*[@id="menu_bar_genericoHistoricoAtendimento"]/li[1]','var_dadosDaCobranca__historico__anexo__','//*[@id="progress-complete-var_dadosDaCobranca__historico__anexo__"]/span','var_dadosDaCobranca__historico__registro__',self.planilha,self.caminho,self.tempo_espera,self) # Envia Anexos
            try:
                if len(self.navegador.find_elements(By.ID, 'id_dadosDaCobranca__acao__')) >= 1:  # Verifica se o campo existe
                    enviarkey_elemento(self.navegador, 'id_dadosDaCobranca__acao__', By.ID, 'Solicitar Nova Medição', self)
            except Exception as e:
                pass
            self.handle_confirmacao_lancamento(nome_cob, aba_original, linha)
            time.sleep(1)
            acessar_iframe(self.navegador,self.tempo_espera, self)
            clicar_elemento_rustico(self.navegador,'clear-input-filter',By.CLASS_NAME, self)#Limpa o campo de Pesquisa
        copiar_para_planilha(self.planilha_destino,self.local_destino)

    def cob_nv(self):
        for linha in range(self.linha_atual, len(self.planilha)):
            self.linha_atual = linha  # Atualiza a linha atual
            acessar_iframe(self.navegador,self.tempo_espera, self)# Acessa o Iframe
            aba_original = self.navegador.window_handles[0] # Identifica Aba Primaria
            clicar_elemento(self.navegador,'btnStartProcess',By.ID, self) # Iniciar novo processo
            clicar_elemento(self.navegador, '//span[text()="Solicitar Cobrança"]', By.XPATH, self)
            WebDriverWait(self.navegador, 10).until(lambda d: len(d.window_handles) > 1)
            nova_aba = self.navegador.window_handles[1]# Identifica nova aba apos iniciar nova Cobrança
            self.navegador.switch_to.window(nova_aba) # Troca para nova Aba
            enviarkey_elemento(self.navegador,'id_informeNucleo__',By.ID,'Núcleo de Faturamento', self)# Envia nucleo - Padrão
            enviarkey_elemento(self.navegador,'id_tipoSolicitacao__',By.ID,'Solicitação de cobrança (FG-176)', self)# Solicitação de cobrança - Padrão
            enviarkey_elemento(self.navegador,'id_plataformaGestaoDaVenda__',By.ID,'Protheus', self)# Plataforma - Padrão
            enviarkey_elemento(self.navegador,'id_tipoDeMedicao__',By.ID,'Variavel', self)# Tipo de medição - Padrão
            if self.planilha.iloc[linha]['PERIODICIDADE'] == 'Não':
                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__APeriodicidadeDoFaturamentoEMensal__',By.ID,'Não', self)# Periodicidade - Não
            else:
                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__APeriodicidadeDoFaturamentoEMensal__',By.ID,'Sim', self)# Periodicidade - Sim
            enviarkey_elemento(self.navegador,"var_dadosDaCobranca__dadosParaHistorico__HouvePrestacaoDeServicos__",By.ID,"Sim", self)
            enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosParaHistorico__numeroContratoProtheus__',By.ID,str(int(self.planilha.iloc[linha]['NUMERO DO CONTRATO'])), self) # Envia Numero de Contrato
            clicar_elemento(self.navegador,'#ac_id_dadosDaCobranca__dadosParaHistorico__numeroContratoProtheus__ ul.ui-autocomplete.ui-front.ui-menu.ui-widget.ui-widget-content.ui-corner-all li.ui-menu-item a.ui-corner-all',By.CSS_SELECTOR, self) # Clica no numero de contrato
            enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosParaHistorico__existeGasOuOS__',By.ID,'Não', self) # Envia Não para Gas ou OS
            enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosParaHistorico__parcelaContrato__',By.ID,'0', self) # Envia 0 para Parcela de Contrato
            enviarkey_elemento(self.navegador, 'var_dadosDaCobranca__dadosParaHistorico__inicioPrestacao__', By.NAME, pd.to_datetime(self.planilha.iloc[linha]['DATA_INICIO']).strftime('%d/%m/%Y'), self) # Envia Data Inicio            
            enviarkey_elemento(self.navegador, 'var_dadosDaCobranca__dadosParaHistorico__finPrestacao__', By.NAME, pd.to_datetime(self.planilha.iloc[linha]['DATA_FIM']).strftime('%d/%m/%Y'), self) # Envia Data Fim
            enviarkey_elemento(self.navegador,'id_dadosDaCobranca__dadosParaHistorico__diaLimiteNFCliente__',By.NAME , int(self.planilha.iloc[linha]['DIA LIMITE']), self) # Envia Dia Limite NF Cliente
            enviarkey_elemento(self.navegador,'var_dadosDaCobranca__cobrancaRelacionadaComConvenio__',By.ID,'Não', self) # Envia Não para Cobrança Relacionada com Convênio
            nome_cob = texto_elemento(self.navegador,'headerTitle',By.ID) # Pega o nome da Cobrança
            data = self.planilha.iloc[linha]['DATA_DESCRIÇÃO'] # Pega data de Descrição
            date = datetime.strptime(data.strftime('%d/%m/%Y'), '%d/%m/%Y') # Transforma data em string
            primeiro_dia, ultimo_dia = primeiro_e_ultimo_dia_do_mes(date.year, date.month) # Pega o mês e dia
            if pd.isna(self.planilha.iloc[linha]['DATA_DE_VENCIMENTO']):
                data_venc = date
            else:
                data_venc = self.planilha.iloc[linha]['DATA_DE_VENCIMENTO'] # Pega data de Vencimento
                if isinstance(data_venc, datetime):# Verifica se a data é um datetime
                    data_venc = data_venc.strftime('%d/%m/%Y')# Considera o dado
                else:
                    pass
            # ---------------------- Esta Parte se refere ao COB sem Rateio ------------------------
            if pd.isna(self.planilha.iloc[linha]['QTD_PARCELA']):
                qtd_repeticao = 1
            else:
                qtd_repeticao = int(self.planilha.iloc[linha]['QTD_PARCELA'])
            for repeticao in range(qtd_repeticao):
                for sem_rateio in range(2):
                    if pd.isna(self.planilha.iloc[linha][f'CR-SR{sem_rateio+1}']): # Verifica se o campo está vazio
                        pass
                    else:
                        clicar_elemento(self.navegador,'createitem',By.ID, self)# Clica para criar novo Item
                        acessar_iframe(self.navegador,self.tempo_espera, self)# Acessa o Iframe
                        enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(self.planilha.iloc[linha]['CNPJ']), self) # Envia CNPJ
                        clicar_elemento_dinamico(self.navegador, self) # Clica no CNPJ informado
                        if sem_rateio == 0: # Difere o primeiro produto do segundo
                            if not pd.isna(self.planilha.iloc[linha]['TEXTO1']): # Verifica se o campo TEXTO1 é maior que 3, se sim Envia o TEXTO1
                                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,self.planilha.iloc[linha]['TEXTO1'], self) # Envia Descrição da coluna TEXTO1
                            else: #Se não envia a descrição padrão
                                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA SESI VIVA+: AEP,PGR,PCMSO,LTCAT \nPERÍODO: {primeiro_dia} a {ultimo_dia}.', self) # Envia Descrição Padrão
                        else:
                            if not pd.isna(self.planilha.iloc[linha]['TEXTO2']): # Verifica se o campo TEXTO2 é maior que 3, se sim Envia o TEXTO2
                                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,self.planilha.iloc[linha]['TEXTO2'], self) # Envia Descrição da coluna TEXTO2
                            else:#Se não envia a descrição padrão
                                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.', self) # Envia Descrição Padrão
                        enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Não', self)# Envia não ao campo de rateio
                        clicar_elemento(self.navegador,'id_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDeCobranca__UOCRProtheus___anchor',By.ID, self)# Clica na pesquisa de produto
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe da Pesquisa
                        clicar_elemento(self.navegador,'vfilter',By.ID, self) # Clica no Filtro
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe do Filtro
                        enviarkey_elemento(self.navegador,'var_codclvlr__',By.NAME,str(int(self.planilha.iloc[linha]['CLASSE DE VALOR'])), self) # Envia Classe de valor Cliente
                        enviarkey_elemento(self.navegador,'var_codfilialprotheus__',By.NAME,self.cod_filial, self) # Envia COD FILIAL - PADRÃO
                        enviarkey_elemento(self.navegador,'var_coduo__',By.NAME,self.cod_uo, self) # Envia COD UO - PADRÃO
                        enviarkey_elemento(self.navegador,'var_codccusto__',By.NAME,str(int(self.planilha.iloc[linha][f'CR-SR{sem_rateio+1}'])), self) # Envia COD PRODUTO
                        clicar_elemento(self.navegador,'searchbutton',By.ID, self) # Clica na Pesquisa
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe da Pesquisa
                        clicar_elemento(self.navegador,'tooltip0',By.ID, self)
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario
                        clicar_elemento(self.navegador,'createitem',By.ID, self) # Clica para adicionar Valor
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe de valor
                        opcoes_pagamento(self.navegador,'//*[@id="mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos_ori"]/option[1]','move_this_right_mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos', self) # Loop para selecionar as opções de pagamento
                        enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__data__',By.NAME,data_venc, self) # Envia data da cobrança
                        enviarkey_java(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__valor__',self.planilha.iloc[linha][f'VALORSR{sem_rateio+1}'], self) # Envia Valor
                        clicar_elemento(self.navegador,'action.save',By.NAME, self) # Clica para salvar.
                        acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario
                        clicar_elemento(self.navegador,'action.save',By.NAME, self) # Clica para salvar.
                        self.navegador.switch_to.default_content()#Volta para o inicio
                # ---------------------- Esta Parte se refere ao COB com Rateio ------------------------
                contador = 0 # Contador utilizado para clicar nos rateios no processo Final!
                contador_1 = 0 # Contador utilizado para clicar nos rateios no processo Final!
                contador_2 = 0 # Contador utilizado para clicar nos rateios no processo Final!
                if pd.isna(self.planilha.iloc[linha]['CRR1']): # Verifica se o primeiro item está vazio, se o mesmo estiver vazio, todo o loop é pulado!
                    pass
                elif pd.isna(self.planilha.iloc[linha]['QTD RATEIO']) or int(self.planilha.iloc[linha]['QTD RATEIO']) == 1 : # Caso não esteja vazio é iniciado o processo de Rateio e a QTD seja um executa todos os rateios em um unico processo
                    clicar_elemento(self.navegador,'createitem',By.ID, self)# Clica para criar novo Item
                    acessar_iframe(self.navegador,self.tempo_espera, self)# Acessa o Iframe
                    enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(self.planilha.iloc[linha]['CNPJ']), self) # Envia CNPJ
                    clicar_elemento_dinamico(self.navegador, self) # Clica no CNPJ informado
                    if not pd.isna(self.planilha.iloc[linha]['TEXTO3']): # Verifica se o campo TEXTO3 é maior que 3, se sim Envia o TEXTO3
                                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,self.planilha.iloc[linha]['TEXTO3'], self) # Envia Descrição da coluna TEXTO3
                    else:#Se não envia a descrição padrão
                        enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.', self) # Envia Descrição
                    enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Sim', self)# Envia sim ao campo de rateio
                    #Loop para a quantidade de Itens
                    for com_rateio in range(4): # Loop para verificar todos os itens (Total 4) com rateio na self.planilha!!
                        if pd.isna(self.planilha.iloc[linha][f'CRR{com_rateio+1}']): # Loop para verificar se o Item está vazio!!
                            pass # Pula o item vazio
                        else:
                            dados_rateio(self.navegador,linha,self.cod_filial,self.cod_uo,self.planilha, self.tempo_espera, self)
                            enviarkey_elemento(self.navegador,'var_codccusto__',By.NAME,str(int(self.planilha.iloc[linha][f'CRR{com_rateio+1}'])), self) # Envia COD PRODUTO
                            clicar_elemento(self.navegador,'searchbutton',By.ID, self) # Clica na Pesquisa
                            acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe da Pesquisa
                            clicar_elemento(self.navegador,'tooltip0',By.ID, self) # Clica no item filtrado
                            acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario3
                            enviarkey_java(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__',self.planilha.iloc[linha][f'VALOR{com_rateio+1}'], self) # Envia Valor
                            clicar_elemento(self.navegador,'action.save',By.NAME, self) # Clica para salvar.
                            acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario
                            contador += 1 # Soma 1 a quantidade de contador, será utiizado para clicar no loop Contador!
                    clicar_porcentagem(self.navegador,contador,self.tempo_espera, self) # Baseado na soma do Contador clica nos itens
                elif int(self.planilha.iloc[linha]['QTD RATEIO']) == 2: # Ira rodar o processo de sem rateio duas vezes uma para a coluna CRR1 e 2 e ou para CRR3 e 4
                    # Processo para coluna 1 e 2
                    clicar_elemento(self.navegador,'createitem',By.ID, self)# Clica para criar novo Item
                    acessar_iframe(self.navegador,self.tempo_espera, self)# Acessa o Iframe
                    enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(self.planilha.iloc[linha]['CNPJ']), self) # Envia CNPJ
                    clicar_elemento_dinamico(self.navegador, self) # Clica no CNPJ informado
                    if not pd.isna(self.planilha.iloc[linha]['TEXTO3']): # Verifica se o campo TEXTO3 é maior que 3, se sim Envia o TEXTO3
                                enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,self.planilha.iloc[linha]['TEXTO3'], self) # Envia Descrição da coluna TEXTO3
                    else:#Se não envia a descrição padrão
                        enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.', self) # Envia Descrição
                    enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Sim', self)# Envia sim ao campo de rateio
                    #Loop para a quantidade de Itens
                    for com_rateio in range(2): # Loop para verificar todos os itens (Total 4) com rateio na self.planilha!!
                        if pd.isna(self.planilha.iloc[linha][f'CRR{com_rateio+1}']): # Loop para verificar se o Item está vazio!!
                            pass # Pula o item vazio
                        else:
                            dados_rateio(self.navegador,linha,self.cod_filial,self.cod_uo,self.planilha, self.tempo_espera, self)
                            enviarkey_elemento(self.navegador,'var_codccusto__',By.NAME,str(int(self.planilha.iloc[linha][f'CRR{com_rateio+1}'])), self) # Envia COD PRODUTO
                            clicar_elemento(self.navegador,'searchbutton',By.ID, self) # Clica na Pesquisa
                            acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe da Pesquisa
                            clicar_elemento(self.navegador,'tooltip0',By.ID, self) # Clica no item filtrado
                            acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario3
                            enviarkey_java(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__',self.planilha.iloc[linha][f'VALOR{com_rateio+1}'], self) # Envia Valor
                            clicar_elemento(self.navegador,'action.save',By.NAME, self) # Clica para salvar.
                            acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario
                            contador_1 += 1 # Soma 1 a quantidade de contador, será utiizado para clicar no loop Contador!
                    clicar_porcentagem(self.navegador,contador_1,self.tempo_espera, self) # Baseado na soma do Contador clica nos itens

                    # Processo para coluna 3 e 4
                    clicar_elemento(self.navegador,'createitem',By.ID, self)# Clica para criar novo Item
                    acessar_iframe(self.navegador,self.tempo_espera, self)# Acessa o Iframe
                    enviarkey_elemento(self.navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,tratar_cnpj(self.planilha.iloc[linha]['CNPJ']), self) # Envia CNPJ
                    clicar_elemento_dinamico(self.navegador, self) # Clica no CNPJ informado
                    if not pd.isna(self.planilha.iloc[linha]['TEXTO4']): # Verifica se o campo TEXTO4 é maior que 4, se sim Envia o TEXTO4
                        enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,self.planilha.iloc[linha]['TEXTO4'], self) # Envia Descrição da coluna TEXTO4
                    else:#Se não envia a descrição padrão
                        enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'IN LOCO COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.', self) # Envia Descrição
                    enviarkey_elemento(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Sim', self)# Envia sim ao campo de rateio
                    #Loop para a quantidade de Itens
                    for com_rateio in range(2): # Loop para verificar todos os itens (Total 4) com rateio na self.planilha!!
                        if pd.isna(self.planilha.iloc[linha][f'CRR{com_rateio+3}']): # Loop para verificar se o Item está vazio!!
                            pass # Pula o item vazio
                        else:
                            dados_rateio(self.navegador,linha,self.cod_filial,self.cod_uo,self.planilha, self.tempo_espera, self)
                            enviarkey_elemento(self.navegador,'var_codccusto__',By.NAME,str(int(self.planilha.iloc[linha][f'CRR{com_rateio+3}'])), self) # Envia COD PRODUTO
                            clicar_elemento(self.navegador,'searchbutton',By.ID, self) # Clica na Pesquisa
                            acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe da Pesquisa
                            clicar_elemento(self.navegador,'tooltip0',By.ID, self) # Clica no item filtrado
                            acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario3
                            enviarkey_java(self.navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__',self.planilha.iloc[linha][f'VALOR{com_rateio+3}'], self) # Envia Valor
                            clicar_elemento(self.navegador,'action.save',By.NAME, self) # Clica para salvar.
                            acessar_iframe(self.navegador,self.tempo_espera, self) # Acessa Iframe primario
                            contador_2 += 1 # Soma 1 a quantidade de contador, será utiizado para clicar no loop Contador!
                    clicar_porcentagem(self.navegador,contador_2,self.tempo_espera, self) # Baseado na soma do Contador clica nos itens
            # ---------------------- Esta Parte se refere aos Anexos ------------------------
            enviar_anexo(self.navegador,linha,'//*[@id="menu_bar_genericoHistoricoAtendimento"]/li[1]','var_dadosDaCobranca__historico__anexo__','//*[@id="progress-complete-var_dadosDaCobranca__historico__anexo__"]/span','var_dadosDaCobranca__historico__registro__',self.planilha,self.caminho,self.tempo_espera,self) # Envia Anexos
            enviar_emails(self.navegador, linha, """li[onclick*="ellist_emailClienteFP__.addNewItem('CreateItens', true);"]""", 'var_emailClienteFP__Email__' ,self.planilha, self.tempo_espera,self)
            self.navegador.switch_to.default_content()
            self.handle_confirmacao_lancamento(nome_cob, aba_original, linha)
            time.sleep(1)
            acessar_iframe(self.navegador,self.tempo_espera, self)

    def pular_linha(self):
        # Atualiza a linha atual para pular para a próxima linha
        self.linha_atual += 1

        # Verifica se há mais de uma aba aberta
        if len(self.navegador.window_handles) > 1:
            self.navegador.close()  # Fecha a aba atual
            self.navegador.switch_to.window(self.navegador.window_handles[0])  # Volta para a aba original
            self.navegador.refresh()  # Atualiza a aba original (F5)
        else:
            self.navegador.switch_to.window(self.navegador.window_handles[0])  # Volta para a aba original
            self.navegador.refresh()  # Atualiza a aba original (F5)

        acessar_iframe(self.navegador, self.tempo_espera, self)
        # Chama o método apropriado
        if self.metodo == 'medicao_vr':
            self.medicao_vr()
        elif self.metodo == 'cob_nv':
            self.cob_nv()

    def repetir_linha(self):
        # Verifica se há mais de uma aba aberta
        if len(self.navegador.window_handles) > 1:
            self.navegador.close()  # Fecha a aba atual
            self.navegador.switch_to.window(self.navegador.window_handles[0])  # Volta para a aba original
            self.navegador.refresh()  # Atualiza a aba original (F5)
        else:
            self.navegador.switch_to.window(self.navegador.window_handles[0])  # Volta para a aba original
            self.navegador.refresh()  # Atualiza a aba original (F5)

        # Não atualiza a linha atual para repetir a execução da linha atual
        acessar_iframe(self.navegador, self.tempo_espera, self)
        if self.metodo == 'medicao_vr':
            self.medicao_vr()
        elif self.metodo == 'cob_nv':
            self.cob_nv()

    def reload_data(self):
        metodo = 'Medição' if self.metodo == 'medicao_vr' else 'Novo'
        """
        Recarrega a planilha com os dados mais recentes.
        """
        self.planilha = pd.read_excel(self.caminho, metodo).apply(lambda col: col.map(lambda x: str(x).replace('\xa0', '') if isinstance(x, str) else x))  # Recarrega a Planilha
        print("Planilha recarregada com dados mais recentes.")

    def inicializacao(self,tipo):
        """
        Inicializa o processo de carregamento e manipulação de uma planilha Excel, além de realizar login em um navegador.

        Returns:
            tuple: Um tuplo contendo o navegador inicializado e a planilha carregada e manipulada.
        """
        self.planilha = pd.read_excel(self.caminho, tipo).apply(lambda col: col.map(lambda x: str(x).replace('\xa0', '') if isinstance(x, str) else x))  # Carrega a Planilha
        copiar_para_planilha(self.local_destino, self.planilha_destino)  # realiza a copia do Historico
        self.navegador, self.chrome_proc = iniciar_navegador()
        root = tk.Tk()
        root.withdraw()  # Oculta a janela principal
        root.attributes('-topmost', True)  # Faz o alerta ficar sempre na frente
        messagebox.showinfo("Login Necessário", "Por favor, faça login no Fusion e clique em OK para continuar.")
        root.destroy()

        return self.navegador, self.chrome_proc, self.planilha

    def fechar_navegador(self):
        """
        Fecha o navegador controlado pelo Selenium e encerra o processo do Chrome se estiver em modo de depuração remota.
        """
        if self.navegador:
            if hasattr(self, 'chrome_proc') and self.chrome_proc:
                try:
                    self.chrome_proc.terminate()
                except Exception:
                    pass 

    def handle_confirmacao_lancamento(self, nome_cob, aba_original, linha):
        resposta = self.confirmacao_lancamento()
        if resposta == "confirmar":
            # Executa as duas ações ao confirmar
            clicar_elemento(self.navegador, 'action.send', By.NAME, self)
            if self.metodo == 'medicao_vr':
                esperar_alerta(self.navegador, nome_cob, aba_original, self.planilha, self.local_destino, 'Medição', linha)
            elif self.metodo == 'cob_nv':
                esperar_alerta(self.navegador,nome_cob, aba_original,self.planilha,self.local_destino,'Novo',linha,nome_cob)
        elif resposta == "repetir_linha":
            # Chama o método repetir_linha
            self.repetir_linha()
        elif resposta == "pular_linha":
            # Chama o método pular_linha
            self.pular_linha()
        elif resposta == "reload_data":
            self.reload_data()  # Recarregar a planilha com dados mais recentes
            self.repetir_linha()
        elif resposta == "confirmar_finalizar":
            # Executa as duas ações e finaliza o programa
            clicar_elemento(self.navegador, 'action.send', By.NAME, self)
            if self.metodo == 'medicao_vr':
                esperar_alerta(self.navegador, nome_cob, aba_original, self.planilha, self.local_destino, 'Medição', linha)
            elif self.metodo == 'cob_nv':
                esperar_alerta(self.navegador,nome_cob, aba_original,self.planilha,self.local_destino,'Novo',linha,nome_cob)
            time.sleep(5)
            copiar_para_planilha(self.planilha_destino,self.local_destino)
            self.fechar_navegador()  # Fecha o navegador
            sys.exit("Programa encerrado pelo usuário.")

    def handle_custom_messagebox_response(self):
        """
        Exibe a caixa de diálogo personalizada e lida com a resposta do usuário.

        Returns:
            bool: True se deve continuar tentando, False se deve parar.
        """
        resposta = self.custom_messagebox()
        if resposta == "try_again":
            return True  # Tentar novamente
        elif resposta == "skip_button":
            return False  # Pular botão
        elif resposta == "repeat_line":
            self.repetir_linha()  # Repetir linha
            return False
        elif resposta == "skip_line":
            self.pular_linha()  # Pular linha
            return False
        elif resposta == "reload_data":
            self.reload_data()  # Recarregar a planilha com dados mais recentes
            self.repetir_linha()
            return False
        elif resposta == "cancel":
            copiar_para_planilha(self.local_destino, self.planilha_destino)
            self.fechar_navegador()  # Fecha o navegador
            sys.exit("Programa encerrado pelo usuário.")  # Encerra o programa
            
    def confirmacao_lancamento(self):
        """
        Exibe uma janela de confirmação com cinco opções: Confirmar, Recusar, Repetir linha, Pular linha e Confirmar e finalizar.

        Returns:
            str: A escolha do usuário ('confirmar', 'recusar', 'repetir_linha', 'pular_linha', 'confirmar_finalizar').
        """
        root = tk.Tk()
        root.withdraw()  # Oculta a janela principal do Tkinter

        # Cria uma nova janela
        confirm_box = tk.Toplevel(root)
        confirm_box.title("Confirmação de Lançamento")
        confirm_box.geometry("800x200")

        # Adiciona uma mensagem
        message = tk.Label(confirm_box, text="Confirma o lançamento?", wraplength=350)
        message.pack(pady=10)

        # Adiciona detalhes
        detail = tk.Label(confirm_box, text="Escolha uma das opções abaixo:", wraplength=350)
        detail.pack(pady=10)

        # Variável para armazenar a resposta
        resposta = tk.StringVar()

        # Funções para definir a resposta e fechar a janela
        def set_resposta(value):
            resposta.set(value)
            confirm_box.destroy()

        # Adiciona botões com o mesmo tamanho
        button_width = 20

        button_confirmar = tk.Button(confirm_box, text="Confirmar", width=button_width, command=lambda: set_resposta("confirmar"))
        button_confirmar.pack(side=tk.LEFT, padx=5, pady=10)

        button_recusar = tk.Button(confirm_box, text="Atualizar", width=button_width, command=lambda: set_resposta("reload_data"))
        button_recusar.pack(side=tk.LEFT, padx=5, pady=10)

        button_repetir_linha = tk.Button(confirm_box, text="Repetir linha", width=button_width, command=lambda: set_resposta("repetir_linha"))
        button_repetir_linha.pack(side=tk.LEFT, padx=5, pady=10)

        button_pular_linha = tk.Button(confirm_box, text="Pular linha", width=button_width, command=lambda: set_resposta("pular_linha"))
        button_pular_linha.pack(side=tk.LEFT, padx=5, pady=10)

        button_confirmar_finalizar = tk.Button(confirm_box, text="Confirmar e finalizar", width=button_width, command=lambda: set_resposta("confirmar_finalizar"))
        button_confirmar_finalizar.pack(side=tk.LEFT, padx=5, pady=10)

        # Espera pela resposta do usuário
        confirm_box.wait_window()

        return resposta.get()
    
    def custom_messagebox(self):
        """
        Exibe uma janela personalizada com cinco botões: Tentar novamente, Pular botão, Repetir linha, Pular linha e Cancelar.

        Returns:
            str: A escolha do usuário ('try_again', 'skip_button', 'repeat_line', 'skip_line', 'cancel').
        """
        root = tk.Tk()
        root.withdraw()  # Oculta a janela principal do Tkinter

        # Cria uma nova janela
        custom_box = tk.Toplevel(root)
        custom_box.title("Alerta")
        custom_box.geometry("800x260")

        # Adiciona uma mensagem
        message = tk.Label(custom_box, text="Botão ou campo não encontrado. O que você gostaria de fazer?", wraplength=350)
        message.pack(pady=10)

        # Adiciona detalhes
        detail = tk.Label(custom_box, text="Aperte 'Tentar novamente' para tentar novamente, 'Pular botão' para continuar para o próximo comando, 'Repetir linha' para repetir a linha, 'Pular linha' para pular para a próxima linha, 'Atualizar' para ler os dados novamente e repetir linha ,ou 'Cancelar' para encerrar o programa.", wraplength=350)
        detail.pack(pady=10)

        # Variável para armazenar a resposta
        resposta = tk.StringVar()

        # Funções para definir a resposta e fechar a janela
        def set_resposta(value):
            resposta.set(value)
            custom_box.destroy()

        # Adiciona botões com o mesmo tamanho
        button_width = 15

        button_try_again = tk.Button(custom_box, text="Tentar novamente", width=button_width, command=lambda: set_resposta("try_again"))
        button_try_again.pack(side=tk.LEFT, padx=5, pady=10)

        button_skip_button = tk.Button(custom_box, text="Pular botão", width=button_width, command=lambda: set_resposta("skip_button"))
        button_skip_button.pack(side=tk.LEFT, padx=5, pady=10)

        button_repeat_line = tk.Button(custom_box, text="Repetir linha", width=button_width, command=lambda: set_resposta("repeat_line"))
        button_repeat_line.pack(side=tk.LEFT, padx=5, pady=10)

        button_skip_line = tk.Button(custom_box, text="Pular linha", width=button_width, command=lambda: set_resposta("skip_line"))
        button_skip_line.pack(side=tk.LEFT, padx=5, pady=10)

        button_skip_line = tk.Button(custom_box, text="Atualizar", width=button_width, command=lambda: set_resposta("reload_data"))
        button_skip_line.pack(side=tk.LEFT, padx=5, pady=10)

        button_cancel = tk.Button(custom_box, text="Cancelar", width=button_width, command=lambda: set_resposta("cancel"))
        button_cancel.pack(side=tk.LEFT, padx=5, pady=10)

        # Espera pela resposta do usuário
        custom_box.wait_window()

        return resposta.get()

    def tratar_erro_critico(self, erro, tipo_operacao):
        """
        Trata erros críticos exibindo uma caixa de diálogo personalizada e executando a ação escolhida pelo usuário.
        """
        navegador_fechado = False
        try:
            # Tenta acessar uma propriedade do navegador
            _ = self.navegador.current_url
        except Exception:
            navegador_fechado = True

        if navegador_fechado or self.navegador is None:
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            tk.messagebox.showerror(
                "Navegador Fechado",
                "O navegador foi fechado!\nPor favor, execute o programa novamente.",
                parent=root
            )
            root.destroy()
            sys.exit("Navegador fechado pelo usuário.")

        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)

        # Cria janela personalizada
        box = tk.Toplevel(root)
        box.title("Erro Crítico")
        box.geometry("600x200")

        msg = tk.Label(box, text=f"Ocorreu um erro:\n{erro}\n\nEscolha uma opção:", wraplength=550)
        msg.pack(pady=15)

        resposta = tk.StringVar()

        def set_resposta(value):
            resposta.set(value)
            box.destroy()

        def reload_page():
            if self.navegador and len(self.navegador.window_handles) > 1:
                main_handle = self.navegador.window_handles[0]
                for handle in self.navegador.window_handles[1:]:
                    self.navegador.switch_to.window(handle)
                    self.navegador.close()
                self.navegador.switch_to.window(main_handle)
                self.navegador.refresh()
                acessar_iframe(self.navegador, self.tempo_espera, self)

        btn_width = 25

        btn_tentar = tk.Button(box, text="Tentar novamente", width=btn_width, command=lambda: set_resposta("tentar"))
        btn_tentar.pack(side=tk.LEFT, padx=10, pady=20)

        btn_pular = tk.Button(box, text="Pular linha e tentar novamente", width=btn_width, command=lambda: set_resposta("pular"))
        btn_pular.pack(side=tk.LEFT, padx=10, pady=20)

        btn_cancelar = tk.Button(box, text="Cancelar", width=btn_width, command=lambda: set_resposta("cancelar"))
        btn_cancelar.pack(side=tk.LEFT, padx=10, pady=20)

        box.wait_window()
        root.destroy()

        escolha = resposta.get()

        if escolha == "tentar":
            # Fecha todas as abas exceto a principal e dá F5
            reload_page()
            if tipo_operacao == "Medição":
                try:
                    self.reload_data()
                    self.medicao_vr()
                except Exception as e2:
                    self.tratar_erro_critico(e2, tipo_operacao)
            elif tipo_operacao == "Novo":
                try:
                    self.reload_data()
                    self.cob_nv()
                except Exception as e2:
                    self.tratar_erro_critico(e2, tipo_operacao)
        elif escolha == "pular":
            # Atualiza linha, fecha abas e tenta novamente
            self.linha_atual += 1
            reload_page()
            if tipo_operacao == "Medição":
                try:
                    self.reload_data()
                    self.medicao_vr()
                except Exception as e2:
                    self.tratar_erro_critico(e2, tipo_operacao)
            elif tipo_operacao == "Novo":
                try:
                    self.reload_data()
                    self.cob_nv()
                except Exception as e2:
                    self.tratar_erro_critico(e2, tipo_operacao)
        else:
            copiar_para_planilha(self.local_destino, self.planilha_destino)
            self.fechar_navegador()  # Fecha o navegador
            sys.exit("Programa encerrado pelo usuário.")
