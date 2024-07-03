from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from tkinter.filedialog import askopenfilename
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import tkinter as tk
import calendar
import time
import locale
import pandas as pd

locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

def acessar_iframe(nav):
    iframe = WebDriverWait(nav, 180).until(EC.presence_of_element_located((By.TAG_NAME, 'iframe'))) # Espera 180 segundos até o iframe aparecer!
    nav.switch_to.frame(iframe) #Troca para o iframe

def acessar_iframe_default(nav, timeout=10, wait_before_switch=2, max_attempts=10):
    time.sleep(wait_before_switch)# Espera antes de mudar para o conteúdo padrão
    nav.switch_to.default_content()
    attempts = 0
    while attempts < max_attempts:
        try:
            iframe = WebDriverWait(nav, timeout).until(EC.presence_of_element_located((By.TAG_NAME, 'iframe')))# Espera até que o iframe esteja presente
            nav.switch_to.frame(iframe)# Troca para o iframe
            return  # Sai da função se for bem-sucedido
        except StaleElementReferenceException as e:
            attempts += 1
            time.sleep(1)

def clicar_elemento(nav,elemento,tipo):
    obj = WebDriverWait(nav, 180).until(EC.presence_of_element_located((tipo, elemento))) # Aguarda 180 segundos até o elemento carregar
    navegador.execute_script("arguments[0].click();", obj) # Clica no objeto utilizando paramentros JavaScript.

def clicar_elemento_rustico(nav,elemento,tipo):
    WebDriverWait(nav, 180).until(EC.presence_of_element_located((tipo, elemento))) # Aguarda 180 segundos até o elemento carregar
    nav.find_element(tipo, elemento).click()# Clicar na Pesquisa da forma tradicional do Selenium!

def enviarkey_elemento(nav,elemento,tipo,texto):
    WebDriverWait(nav, 180).until(EC.presence_of_element_located((tipo, elemento))) # Aguarda 180 segundos até o elemento carregar
    navegador.find_element(tipo, elemento).send_keys(texto) # envia dados para o elemento

def primeiro_e_ultimo_dia_do_mes(ano, mes):
    primeiro_dia = datetime(ano, mes, 1)# Primeiro dia do mês
    ultimo_dia = datetime(ano, mes, calendar.monthrange(ano, mes)[1])# Último dia do mês
    # Formatar as datas no formato dd/mm/aaaa
    primeiro_dia_formatado = primeiro_dia.strftime('%d/%m/%Y')
    ultimo_dia_formatado = ultimo_dia.strftime('%d/%m/%Y')
    return primeiro_dia_formatado, ultimo_dia_formatado

def copiar_linhas_ativas(origem, destino):
    
    planilha_origem = pd.read_excel(origem)# Carrega a planilha original
    linhas_ativas = planilha_origem.dropna(how='all')# Seleciona as linhas ativas (não vazias)
    try:
        book = load_workbook(destino) # Tenta carregar a planilha de destino existente
        sheet = book.active  # Assume que a planilha ativa é onde queremos adicionar os dados
        # Encontra a próxima linha vazia na planilha de destino
        next_row = sheet.max_row + 1
        print(f"Adicionando a partir da linha {next_row}")
        # Adiciona as linhas ativas à planilha de destino
        for r_idx, row in enumerate(dataframe_to_rows(linhas_ativas, index=False, header=False), start=next_row):
            for c_idx, value in enumerate(row, 1):
                sheet.cell(row=r_idx, column=c_idx, value=value)
        # Salva o arquivo de destino
        book.save(destino)
    except Exception as e:
        print(f"Erro ao carregar o arquivo de destino: {e}")

def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title="Selecione a Planilha COB!") # Solciita o usuario selecionar a planilha!
    return caminho_arquivo

print('-----------Automação COB-----------/n')
caminho = selecionar_arquivo()
planilha = pd.read_excel(caminho) # Carrega a Planilha
numero_de_linhas = len(planilha) # Conta a quantidade de linhas
usuario = input('Insira o usuario do Fusion: ')
senha = input('Insira a senha do Fusion: ')

# Start no Navegador Chrome
servico = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions() # Para o mesmo não fechar apos execução
options.add_experimental_option("detach", True) # Para o mesmo não fechar apos execução
navegador = webdriver.Chrome(options=options,service=servico) # Executa o navegador
navegador.get('https://fusion.fiemg.com.br/fusion/portal')
navegador.maximize_window() # Maximiza a janela do navegador

enviarkey_elemento(navegador,'user',By.ID,usuario)# Login
enviarkey_elemento(navegador,'pass',By.ID,senha)# Senha
clicar_elemento(navegador,'btnLogin',By.ID) # Clica no botão de Login
acessar_iframe_default(navegador)# Acessa o Iframe

cod_filial = '01MG0014' # Codigo Filial - Padrão
cod_uo = '10310' # Codigo UO - Padrão

for linha in range(len(planilha)):

    enviarkey_elemento(navegador,'searchBarProcessQuery',By.ID,planilha.iloc[linha]['COB'])#Envio do COB
    while len(navegador.find_elements(By.CLASS_NAME, "item")) == 0: # Loop para aguardar a lista de itens carregar, se a lista não carregar a pesquisa não funciona!
        time.sleep(1)
    time.sleep(1)

    clicar_elemento_rustico(navegador,'//*[@id="page-content-wrapper"]/div/div/div[1]/div[1]/nav/div/form/div/div/span/button',By.XPATH) # Clica no botão de pesquisa inicial
    aba_orignal = navegador.window_handles[0] # Identifica Aba Primaria
    clicar_elemento_rustico(navegador, 'header', By.CLASS_NAME) # Clica no COB pesquisado

    nova_aba = navegador.window_handles[1]# Identifica nova aba apos clicar no COB
    navegador.switch_to.window(nova_aba) # Troca para nova Aba

    # ---------------------- Esta Parte se refere ao COB sem Rateio ------------------------
    for sem_rateio in range(2):
        if pd.isna(planilha.iloc[linha][f'CR-SR{sem_rateio+1}']):
            pass
        else:
            clicar_elemento(navegador,'createitem',By.ID)# Clica para criar novo Item
            acessar_iframe(navegador)# Acessa o Iframe
            enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,str(int(planilha.iloc[linha]['CNPJ'].replace('.','').replace('/','').replace('-','')))) # Envia CNPJ
            clicar_elemento(navegador,'ui-id-11',By.ID) # Clica no CNPJ informado
            data_venc = planilha.iloc[linha]['DATA_DE_VENCIMENTO']
            data = planilha.iloc[linha]['DATA_DESCRIÇÃO']
            date = datetime.strptime(data.strftime('%d/%m/%Y'), '%d/%m/%Y') # Transforma data em string
            primeiro_dia, ultimo_dia = primeiro_e_ultimo_dia_do_mes(date.year, date.month) # Pega o mês e dia
            if sem_rateio == 0:
                enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA SESI VIVA+: AEP,PGR,PCMSO,LTCAT \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
            else:
                enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
            enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Não')# Envia não ao campo de rateio
            clicar_elemento(navegador,'id_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDeCobranca__UOCRProtheus___anchor',By.ID)# Clica na pesquisa de produto
            acessar_iframe_default(navegador) # Acessa Iframe da Pesquisa
            clicar_elemento(navegador,'vfilter',By.ID) # Clica no Filtro
            acessar_iframe_default(navegador) # Acessa Iframe do Filtro
            enviarkey_elemento(navegador,'var_codclvlr__',By.NAME,str(int(planilha.iloc[linha]['CLASSE DE VALOR']))) # Envia Classa de valor Cliente
            enviarkey_elemento(navegador,'var_codfilialprotheus__',By.NAME,cod_filial) # Envia COD FILIAL - PADRÃO
            enviarkey_elemento(navegador,'var_coduo__',By.NAME,cod_uo) # Envia COD UO - PADRÃO
            enviarkey_elemento(navegador,'var_codccusto__',By.NAME,str(int(planilha.iloc[linha][f'CR-SR{sem_rateio+1}']))) # Envia COD PRODUTO
            clicar_elemento(navegador,'searchbutton',By.ID) # Clica na Pesquisa
            acessar_iframe_default(navegador) # Acessa Iframe da Pesquisa
            clicar_elemento(navegador,'tooltip0',By.ID)
            acessar_iframe_default(navegador) # Acessa Iframe primario
            clicar_elemento(navegador,'createitem',By.ID) # Clica para adicionar Valor
            acessar_iframe_default(navegador) # Acessa Iframe de valor

            #Loop para selecionar as opções de pagamento
            for i in range(2):
                navegador.find_element(By.XPATH, '//*[@id="mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos_ori"]/option[1]').click()
                navegador.find_element(By.ID, 'move_this_right_mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__formaDeEntradaDosRecursos').click()
                
            enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__data__',By.NAME,data_venc.strftime('%d/%m/%Y')) # Envia data da cobrança
            script_valor_sr = f"document.getElementsByName('var_dadosDaCobranca__dadosDoFaturamentoVariavel__dataVencimentoValorCobranca__valor__')[0].value='{planilha.iloc[linha][f'VALORSR{sem_rateio+1}']}';"
            navegador.execute_script(script_valor_sr)
            input('enter')
            clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
            acessar_iframe_default(navegador) # Acessa Iframe primario
            enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__numeroContratoProtheus__',By.ID,str(int(planilha.iloc[linha]['NUMERO DO CONTRATO']))) # Envia o numero de contrato
            clicar_elemento(navegador,'//*[@id="ui-id-10"]/li',By.XPATH) # Clica no numero de contrato
            clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
            navegador.switch_to.default_content()#Volta para o inicio

    # ---------------------- Esta Parte se refere ao COB com Rateio ------------------------
    contador = 0 # Contador utilizado para clicar nos rateios no processo Final!
    if pd.isna(planilha.iloc[linha]['CRR1']): # Verifica se o primeiro item está vazio, se o mesmo estiver vazio, todo o loop é pulado!
        pass
    else: # Caso não esteja vazio é iniciado o processo de Rateio
        clicar_elemento(navegador,'createitem',By.ID)# Clica para criar novo Item
        acessar_iframe(navegador)# Acessa o Iframe
        enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoCliente__',By.ID,str(int(planilha.iloc[linha]['CNPJ'].replace('.','').replace('/','').replace('-','')))) # Envia CNPJ
        clicar_elemento(navegador,'ui-id-11',By.ID) # Clica no CNPJ informado
        primeiro_dia, ultimo_dia = primeiro_e_ultimo_dia_do_mes(date.year, date.month) # Pega o mês e dia
        enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__descricaoServico__',By.NAME,f'COBRANÇA CONSULTAS E EXAMES COMPLEMENTARES. \nPERÍODO: {primeiro_dia} a {ultimo_dia}.') # Envia Descrição
        enviarkey_elemento(navegador,'var_dadosDaCobranca__dadosDoFaturamentoVariavel__rateio__', By.NAME,'Sim')# Envia sim ao campo de rateio
        #Loop para a quantidade de Itens

        for com_rateio in range(4): # Loop para verificar todos os itens (Total 4) com rateio na planilha!!
            if pd.isna(planilha.iloc[linha][f'CRR{com_rateio+1}']): # Loop para verificar se o Item está vazio!!
                pass # Pula o item vazio
            else:
                clicar_elemento(navegador,'//*[@id="menu_bar_FINFFCobFaturamentoVariavelCentroDeResultadosXFilialXValor"]/li[1]',By.XPATH) # Clica para abrir campo de produtos
                acessar_iframe_default(navegador) # Acessa Iframe da Pesquisa de produtos

                #Loop para selecionar as opções de pagamento
                for i in range(2):
                    navegador.find_element(By.XPATH, '//*[@id="mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__formaDeEntradaDosRecursos_ori"]/option[1]').click()
                    navegador.find_element(By.ID, 'move_this_right_mul_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__formaDeEntradaDosRecursos').click()

                clicar_elemento(navegador,'id_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__UOCRProtheus___anchor',By.ID) # Clica para abrir campo de pesquisa
                acessar_iframe_default(navegador) # Acessa Iframe da Pesquisa
                clicar_elemento(navegador,'//*[@id="menu_bar_EXTERNOProtheusAmarracaoContabil"]/li',By.XPATH) # Clica para abrir filtro
                acessar_iframe_default(navegador) # Acessa Iframe do Filtro
                enviarkey_elemento(navegador,'var_codclvlr__',By.NAME,str(planilha.iloc[linha]['CLASSE DE VALOR'])) # Envia Classa de valor Cliente
                enviarkey_elemento(navegador,'var_codfilialprotheus__',By.NAME,cod_filial) # Envia COD FILIAL - PADRÃO
                enviarkey_elemento(navegador,'var_coduo__',By.NAME,cod_uo) # Envia COD UO - PADRÃO
                enviarkey_elemento(navegador,'var_codccusto__',By.NAME,str(int(planilha.iloc[linha][f'CRR{com_rateio+1}']))) # Envia COD PRODUTO
                clicar_elemento(navegador,'searchbutton',By.ID) # Clica na Pesquisa
                acessar_iframe_default(navegador) # Acessa Iframe da Pesquisa
                clicar_elemento(navegador,'tooltip0',By.ID) # Clica no item filtrado
                acessar_iframe_default(navegador) # Acessa Iframe primario3
                elemento2 = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.NAME, 'var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__')))
                script_valor_cr = f"document.getElementsByName('var_dadosDaCobranca__dadosDoFaturamentoVariavel__dadosDoRateio__valor__')[0].value='{planilha.iloc[linha][f'VALOR{com_rateio+1}']}';"
                navegador.execute_script(script_valor_cr)
                input('enter')
                clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
                acessar_iframe_default(navegador) # Acessa Iframe primario
                contador += 1 # Soma 1 a quantidade de contador, será utiizado para clicar no loop Contador!

        for i in range(contador): # Baseado na soma do Contador clica nos itens
            clicar_elemento_rustico(navegador,f'//*[@id="{i}"]/td[2]',By.XPATH) # Clica no Item baseado nos indices (No fusion o indice 0 conta!)!!
            acessar_iframe_default(navegador) # Acessa Iframe primario
            clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
            acessar_iframe_default(navegador) # Acessa Iframe primario

        enviarkey_elemento(navegador,'id_txt_dadosDaCobranca__dadosDoFaturamentoVariavel__numeroContratoProtheus__',By.ID,str(int(planilha.iloc[linha]['NUMERO DO CONTRATO']))) # Envia o numero de contrato
        clicar_elemento(navegador,'//*[@id="ui-id-10"]/li',By.XPATH) # Clica no numero de contrato
        clicar_elemento(navegador,'action.save',By.NAME) # Clica para salvar.
        navegador.switch_to.default_content()#Volta para o inicio  
         
    # ---------------------- Esta Parte se refere aos Anexos ------------------------
    for anexo in range(2): 
        if pd.isna(planilha.iloc[linha][f'ARQUIVO{anexo+1}']):
            pass
        else:
            clicar_elemento(navegador,'//*[@id="menu_bar_genericoHistoricoAtendimento"]/li[1]',By.XPATH) # Clica no anexo para enviar arquivo
            acessar_iframe(navegador)#Acesso Iframe
            enviarkey_elemento(navegador,'var_dadosDaCobranca__historico__anexo__',By.ID,fr"{caminho[:-16]}Arquivos\{planilha.iloc[linha][f'ARQUIVO{anexo+1}']}") # Envia o anexo
            while len(navegador.find_elements(By.XPATH, '//*[@id="progress-complete-var_dadosDaCobranca__historico__anexo__"]/span')) == 0: # Loop para aguardar a lista de itens carregar, se a lista não carregar a pesquisa não funciona!
                time.sleep(1)
            time.sleep(1)
            enviarkey_elemento(navegador,'var_dadosDaCobranca__historico__registro__',By.ID,planilha.iloc[linha][f'DESCRICAO{anexo+1}']) # Envia a descrição
            clicar_elemento(navegador,'//*[@id="dibButtons"]/input[1]',By.XPATH) 
            navegador.switch_to.default_content()#Volta para o inicio

    input('Enter')
    clicar_elemento(navegador,'action.send',By.NAME)
    while len(navegador.find_elements(By.CLASS_NAME, 'alert')) == 0: # Loop para aguardar o alerta carregar!
        time.sleep(1)
    time.sleep(1)
    navegador.close() # Fecha o navegador apos Alerta Carregar!!

    navegador.switch_to.window(aba_orignal)
    time.sleep(1)
    acessar_iframe_default(navegador)
    clicar_elemento_rustico(navegador,'clear-input-filter',By.CLASS_NAME)#Limpa o campo de Pesquisa
