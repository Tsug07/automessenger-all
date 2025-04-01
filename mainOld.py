import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import time
import os
import psutil
import re
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException

"""
Automação de mensagem padrão de prorrogação de contrato de experiencia no Onvio Messenger.

"""

# Variável global para controlar o cancelamento do processamento
cancelar = False

#FUNÇÕES DO PROGRAMA
def focar_barra_endereco_e_navegar(driver, termo_busca):
    try:
        
        time.sleep(1)
        # Localiza e clica barra de contato
        focused_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="page-content"]/div/div[2]/div[3]/div[1]/div/div/div/input'))
        )
        
            
        # Simula a pressão da tecla Tab várias vezes
        # for i in range(15):  # Ajuste este número conforme necessário
            
        # Verifica se o elemento focado é a barra de pesquisa
        # focused_element = driver.switch_to.active_element
        if focused_element.get_attribute('placeholder') == "Buscar contatos...":
            focused_element.click()
            atualizar_log(f"Verificando contato {termo_busca}...")
            
            
            # Tenta preencher o campo
            try:
                # Verifica se o texto foi inserido corretamente
                valor_atual = focused_element.get_attribute('value')
                if termo_busca != valor_atual:
                    focused_element.clear()  # Limpa o campo primeiro
                    focused_element.send_keys(termo_busca)
                    atualizar_log(f"Texto '{termo_busca}' inserido com sucesso na barra de pesquisa.")
                    time.sleep(1)  # Espera um pouco para o texto ser inserido
                
                
                elif valor_atual == termo_busca:
                    atualizar_log(f"Texto '{termo_busca}' ja presente na barra de pesquisa.")
                else:
                    atualizar_log(f"Texto inserido não corresponde. Valor atual: '{valor_atual}'")
            except Exception as e:
                atualizar_log(f"Erro ao preencher o campo: {str(e)}")
            
            
            return True
        elif focused_element.get_attribute('placeholder') == "Buscar grupos...":
            atualizar_log(f"Barra de pesquisa encontrada")
            
            
            # Tenta preencher o campo
            try:
                valor_atual = focused_element.get_attribute('value')
                if termo_busca != valor_atual:
                    focused_element.clear()  # Limpa o campo primeiro
                    focused_element.send_keys(termo_busca)
                    time.sleep(1)  # Espera um pouco para o texto ser inserido
                
                # Verifica se o texto foi inserido corretamente
                
                elif valor_atual == termo_busca:
                    atualizar_log(f"Texto '{termo_busca}' inserido com sucesso na barra de pesquisa.")
                else:
                    atualizar_log(f"Texto inserido não corresponde. Valor atual: '{valor_atual}'")
            except Exception as e:
                atualizar_log(f"Erro ao preencher o campo: {str(e)}")
            
            
            return True
                
            
            
            # ActionChains(driver).send_keys(Keys.TAB).perform()
            # time.sleep(1)
            # atualizar_log(f"TAB {i+1} pressionado")
            
        atualizar_log("Barra de pesquisa não encontrada após 10 TABs")
        return False


    except Exception as e:
        atualizar_log(f"Erro ao focar na barra de endereço ou navegar: {str(e)}")
        return False

def encerrar_processos_chrome():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'chrome.exe':
            try:
                proc.terminate()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    time.sleep(2)  # Aguarda um pouco para garantir que os processos foram encerrados

def abrir_chrome_com_url(url):
    encerrar_processos_chrome()
    
    # Caminho para o perfil padrão do Chrome
    user_data_dir = os.path.expanduser('~') + r'\AppData\Local\Google\Chrome\User Data'
    
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument(f"user-data-dir={user_data_dir}")
    chrome_options.add_argument("--profile-directory=Default")
    chrome_options.add_argument("--disable-translate")  # Tenta desabilitar a tradução automática
    chrome_options.add_argument("--lang=pt-BR")  # Define o idioma para português do Brasil
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_experimental_option("prefs", {
        "translate": {"enabled": "false"},
        "profile.default_content_setting_values.notifications": 2
    })
    
    
    service = Service(ChromeDriverManager().install())
    
    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_page_load_timeout(180)  # Aumenta o tempo limite para carregar a página
        driver.get(url)
        atualizar_log(f"Chrome aberto com a URL: {url}")
        return driver
    except Exception as e:
        atualizar_log(f"Erro ao abrir o Chrome: {str(e)}")
        return None

def esperar_carregamento_completo(driver):
    global cancelar
    cancelar = False
    try:
        WebDriverWait(driver, 60).until(
            lambda d: d.execute_script('return document.readyState') == 'complete'
        )
        atualizar_log("Página completamente carregada.")
    except TimeoutException as e:
        atualizar_log(f"Erro ao esperar o carregamento completo: {str(e)}")

def processar_resultados_busca(driver):
    global cancelar
    cancelar = False
    try:
        # Espera pelos resultados da busca
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="page-content"]/div/div[2]/div[3]/div[2]/div/div[1]'))
        )
        
        
        # Localiza e tenta clicar no elemento especificado(1º CONTATO)
        elemento_alvo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="page-content"]/div/div[2]/div[3]/div[2]/div/div[1]'))
        )
        
        # Tenta clicar no elemento
        if elemento_alvo:
            elemento_alvo.click()
            atualizar_log("Clicado no elemento alvo.")
        else:
            atualizar_log("Elemento não encontrado.")
        
        
    except TimeoutException:
        atualizar_log("Timeout ao esperar pelo elemento alvo.", cor='vermelho')
    except Exception as e:
        atualizar_log(f"Erro ao processar resultados da busca ou clicar no elemento: {str(e)}")

def focar_barra_mensagem_enviar(driver, mensagem):
    global cancelar
    if cancelar:
        atualizar_log("Processamento cancelado!", cor="azul")
        return
    try:
        elemento_alvo = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="preview-root"]/div[2]/div[3]/div[1]/div/div[2]/div[2]/div[1]'))
        )
        
        if elemento_alvo.get_attribute('data-placeholder') == "Mensagem":
            elemento_alvo.click()
            # atualizar_log(f"Barra de mensagem encontrada após {i+1} TABs")
            atualizar_log(f"Barra de Mensagem encontrada e clicada!")
            atualizar_log(f"Enviando Mensagem...")
            
            # Tenta preencher o campo
            try:
                    # Divide a mensagem em parágrafos
                paragrafos = re.split(r'\n+', mensagem.strip())

                # Insere cada parágrafo como um novo elemento <p>
                for i, paragrafo in enumerate(paragrafos):
                    if i > 0:
                        # Para todos exceto o primeiro parágrafo, pressiona Shift+Enter para criar um novo <p>
                        ActionChains(driver).key_down(Keys.SHIFT).send_keys(Keys.ENTER).key_up(Keys.SHIFT).perform()
                        time.sleep(0.5)
                        
                    if cancelar:
                        atualizar_log("Processamento cancelado!", cor="azul")
                        return
                    # Insere o texto do parágrafo
                    ActionChains(driver).send_keys(paragrafo).perform()
                    time.sleep(0.5)

                atualizar_log("Mensagem formatada inserida com sucesso.")
                if cancelar:
                    atualizar_log("Processamento cancelado!", cor="azul")
                    return
                # Agora vamos clicar no botão de enviar
                # Primeiro, tentamos localizar o botão por XPath
                # try:
                #     botao_enviar = WebDriverWait(driver, 10).until(
                #         EC.element_to_be_clickable((By.XPATH, '//*[@id="preview-root"]/div[2]/div[3]/div[3]/div[1]/button'))
                #     )
                #     botao_enviar.click()
                #     atualizar_log("Botão de enviar clicado com sucesso.")
                # except:
                #     atualizar_log("\nNão foi possível clicar no botão de enviar por XPath. Tentando por Tab.\n", cor="vermelho")
                    
                #     return False
                    
                
                
                time.sleep(2)  # Espera um pouco para a mensagem ser enviada
                
                return True
                
            except Exception as e:
                atualizar_log(f"Erro ao preencher o campo ou enviar a mensagem: {str(e)}", cor="vermelho")
            
            # break  # Sai do loop se encontrou e interagiu com a barra de mensagem
                    
        else:  # Se o loop terminar sem encontrar a barra de mensagem
            atualizar_log(f"Barra de mensagem não encontrada após {i+1} TABs.")
        
    except Exception as e:
        atualizar_log(f"Erro ao focar na barra de mensagem ou enviar: {str(e)}")
    
    return False

def encontrar_e_clicar_barra_contatos(driver, contato, grupo):
    global cancelar
    if cancelar:
        atualizar_log("Processamento cancelado!", cor="azul")
        return
    try:
        esperar_carregamento_completo(driver)
        atualizar_log("Aguardando elementos específicos da página...")

        time.sleep(5)
        atualizar_log("Focando na aba de contato correspondente...")  
        
    # ABA GRUPO         
        if grupo.upper() != "NONE":
            if cancelar:
                atualizar_log("Processamento cancelado!", cor="azul")
                return
            focar_pagina(driver)
            
            if focar_barra_endereco_e_navegar(driver, grupo):
                if cancelar:
                    atualizar_log("Processamento cancelado!", cor="azul")
                    return
                atualizar_log("Navegação aba grupo.")
                processar_resultados_busca(driver)
                atualizar_log("Navegação bem-sucedida. Busca, ABA GRUPO, realizada.")
                return True
             
    # ABA CONTATO  
        elif contato.upper() != "NONE".upper():
            if cancelar:
                atualizar_log("Processamento cancelado!", cor="azul")
                return
            
            atualizar_log("Navegação aba Contato.")
            time.sleep(1)
            clicar_voltar_lista_contatos(driver)
            if focar_barra_endereco_e_navegar(driver, contato):
                if cancelar:
                    atualizar_log("Processamento cancelado!", cor="azul")
                    return
                processar_resultados_busca(driver)
                atualizar_log("Navegação bem-sucedida. Busca, ABA CONTATO, realizada.") 
                return True
        else:
            atualizar_log("Falha na navegação ou busca.", cor="vermelho")
            return False

    except Exception as e:
        atualizar_log(f"Erro geral ao tentar interagir com a página: {str(e)}", cor="vermelho")
        return False
   
def clicar_voltar_lista_contatos(driver):
    global cancelar
    if cancelar:
        atualizar_log("Processamento cancelado!", cor="azul")
        return
    try:
        
        #xpath perto da barra de contados
        elemento = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="react-tabs-0"]'))
        )
        elemento.click()
        atualizar_log("Clicado na aba Contatos.\n", cor="azul")
        time.sleep(1)  # Espera para a página carregar
        return True
    except Exception as e:
        atualizar_log(f"Erro ao clicar no botão para voltar à lista de contatos: {str(e)}", cor="vermelho")
        return False
    
def focar_pagina(driver):
    global cancelar
    if cancelar:
        atualizar_log("Processamento cancelado!", cor="azul")
        return
    try:    
        # focar na aba grupos dos contatos
        elemento = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="react-tabs-2"]'))
        )
        
        elemento.click()
        atualizar_log("Clicado na aba Grupo.", cor="azul")
        time.sleep(3)  # Espera para a página carregar
        return True
    except Exception as e:
        atualizar_log(f"Erro ao clicar no botão para voltar à lista de contatos: {str(e)}", cor="vermelho")
        return False

def focar_pagina_geral(driver):
    global cancelar
    if cancelar:
        atualizar_log("Processamento cancelado!", cor="azul")
        return
    try:    
        # focar na aba contatos Geral
        elemento = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="page-content"]/div/div[2]/div[1]/div/div/div/div[1]/div/div[1]'))
        )
        
        elemento.click()
        atualizar_log("Clicado no contato geral.")
        time.sleep(3)  # Espera para a página carregar
        return True
    except Exception as e:
        atualizar_log(f"Erro ao clicar no botão para voltar à lista de contatos: {str(e)}", cor="vermelho")
        return False
    
def mensagemPadrao():
    global cancelar
    if cancelar:
        atualizar_log("Processamento cancelado!", cor="azul")
        return
    
    # # Mensagem padrão
    # mensagem = f"Boa tarde!  Espero que esteja bem. O contrato de experiência de {pessoa} vence em {data}. Gostaríamos de saber se haverá prorrogação.  Caso não tenhamos um retorno, consideraremos a prorrogação automática. Agradeço a atenção!"
    
    mensagem = f"Prezado cliente,\n"
    mensagem += " \n"
    mensagem += f"Estamos no período da entrega da Declaração do Imposto de Renda Pessoa Física (IRPF-2024). Para garantir o correto preenchimento e envio da sua declaração à Receita Federal, solicitamos que nos envie as informações necessárias até o dia 15/04/2025. Caso os documentos não sejam encaminhados até essa data, consideraremos que não há informações adicionais e providenciaremos a sua declaração com os dados disponíveis em nosso sistema. Dessa forma, asseguramos o cumprimento do prazo legal e evitamos multas e penalidades.\n"
    mensagem += " \n"
    mensagem += f"Se optar por preencher a sua própria declaração, pedimos que nos informe dentro do mesmo prazo para evitar envios duplicados.\n"
    mensagem += " \n"
    mensagem += "Segue abaixo a relação dos documentos necessários para a elaboração da sua declaração:\n"
    mensagem += " \n"
    mensagem += "Documentos pessoais (Título de Eleitor, CPF, comprovante de residência e dados bancários);\n"
    mensagem += " \n"
    mensagem += "Informe de rendimentos (fornecido pela empresa);\n"
    mensagem += " \n"
    mensagem += "Documentos pessoais dos dependentes (CPF obrigatório, e-mail, celular e comprovante de residência);\n"
    mensagem += " \n"
    mensagem += "Informe de rendimentos financeiros e de aplicações ou extrato de aplicações (fornecidos pelo banco);\n"
    mensagem += " \n"
    mensagem += "Comprovantes de despesas médicas (nome, endereço, CPF ou CNPJ do prestador);\n"
    mensagem += " \n"
    mensagem += "Comprovantes de despesas com ensino;\n"
    mensagem += " \n"
    mensagem += "Extrato de Previdência Privada;\n"
    mensagem += " \n"
    mensagem += "Documentação do Plano de Saúde;\n"
    mensagem += " \n"
    mensagem += "Documentação de imóveis e veículos (inclusive financiados);\n"
    mensagem += " \n"
    mensagem += "Recibos de pagamento ou recebimento de aluguel;\n"
    mensagem += " \n"
    mensagem += "Recibos de doações;\n"
    mensagem += " \n"
    mensagem += "Documentação de consórcios contemplados ou não;\n"
    mensagem += " \n"
    mensagem += "Senha da conta gov.br.\n"
    mensagem += " \n"
    mensagem += "Caso tenha dúvidas, nossa equipe de especialistas está à disposição para orientá-lo.\n"
    mensagem += " \n"
    mensagem += "Atenciosamente,\n"
    mensagem += "Equipe Canella e Santos"
    
    return mensagem


def ler_dados_excel(caminho_excel):
    try:
        # Carrega o workbook
        wb = openpyxl.load_workbook(caminho_excel)
        
        # Seleciona a primeira planilha
        sheet = wb.active
        
        # Lista para armazenar os dados
        dados = {}
         
        # Itera pelas linhas da planilha
        for row in sheet.iter_rows(min_row=2, values_only=True):  # Começa da segunda linha, assumindo que a primeira é cabeçalho
            if row and len(row) >= 4:  # Verifica se a linha tem pelo menos 4 colunas
                codigo, pessoas, nome_contato, nome_grupo = row[:4]
                
                dados[codigo] = {
                    'codigo': codigo,
                    'empresa': pessoas,
                    'nome_contato': nome_contato,
                    'nome_grupo': nome_grupo
                }
            else:
                atualizar_log(f"Linha ignorada por não ter todas as colunas necessárias (esperado 4, encontrado {len(row) if row else 0}): {row}")
        
        if not dados:
            atualizar_log("Nenhum dado válido encontrado no arquivo Excel.", cor="vermelho")
            return None
        
        return dados
    
    except Exception as e:
        atualizar_log(f"Erro ao ler o arquivo Excel: {str(e)}", cor="vermelho")
        return None

def extrair_cod_nome_contatos_e_grupos(dados):
    codigo = []
    nome_contato = []
    nome_grupo = []
    empresas = []
    
    # Iterar sobre o dicionário, onde a chave é o código da empresa
    for cod, info in dados.items():
        codigo.append(cod)  # A chave é o código da empresa
        nome_contato.append(info['nome_contato'])  # Extrair o nome do contato
        nome_grupo.append(info['nome_grupo'])  # Extrair o nome do grupo
        empresas.append(info['empresa'])  # Extrair o nome da empresa
    
    return codigo, empresas, nome_contato, nome_grupo

        
# Função para selecionar o arquivo Excel
def selecionar_excel():
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=(("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*"))
    )
    if arquivo:
        caminho_excel.set(arquivo)
        atualizar_log(f"Arquivo Excel selecionado: {arquivo}")

# Função para iniciar o processamento dos dados
def iniciar_processamento():
    global cancelar
    cancelar = False
    
    excel = caminho_excel.get()
    if excel:
        atualizar_log("Iniciando processamento...", cor="azul")
        botao_iniciar.config(state=tk.DISABLED)  # Desabilitar o botão para evitar múltiplos cliques
        thread = threading.Thread(target=processar_dados, args=(excel,))
        thread.start()
    else:
        messagebox.showwarning("Atenção", "Por favor, selecione o arquivo Excel.")

# Função do seu código que já está pronta para processar os dados
def processar_dados(excel):
    global cancelar

    url_desejada = "https://app.gestta.com.br/attendance/#/chat/contact-list"
    driver = abrir_chrome_com_url(url_desejada)
    
    # Arquivo Excel
    caminho_excel = excel
    dados = ler_dados_excel(caminho_excel)
    codigos, empresas, nome_contatos, nome_grupos = extrair_cod_nome_contatos_e_grupos(dados)
    # print(codigo)
      
    if not driver:
        atualizar_log("Falha ao abrir o Chrome.", cor="vermelho")
        return

    try:
        atualizar_log("Chrome aberto com sucesso.", cor="azul")
        atualizar_log("Aguardando 10 segundos para garantir carregamento completo...")
        time.sleep(10)
        

        #RODAR NOME DE CONTATOS E GRUPOS
        for codigo, empresa, nome_contato, nome_grupo in zip(codigos, empresas, nome_contatos, nome_grupos):
            if cancelar:
                atualizar_log("Processamento cancelado!", cor="azul")
                return 
            
            atualizar_log(f"\nProcessando contato da pessoa {codigo} - {empresa}: Contato: {nome_contato}, Grupo: {nome_grupo}\n", cor="azul")

            # atualizar_log(f"\nProcessando contato da pessoa {codigo} - {pessoa_individual}\n", cor="azul")
            mensagem = mensagemPadrao()
                    
            if nome_contato.upper() != "NONE":
                if cancelar:
                    atualizar_log("Processamento cancelado!", cor="azul")
                    return 
                # Encontra a barra de Pesquisa e processa o contato
                if encontrar_e_clicar_barra_contatos(driver, nome_contato, nome_grupo):
                    
                    atualizar_log(f"Contato {nome_contato} na Aba Contato, encontrado e clicado.")
                    try: 
                        # Espera pela area da mensagem
                        area_mensagem = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="preview-root"]/div[2]'))
                        )
                        
                        if area_mensagem:
                            atualizar_log(f"carregamento Barra de Mensagens 6 segundos ...")
                            # Aguarda a página do chat carregar
                            time.sleep(6)
                                
                            # Foca na barra de mensagem e envia
                            if focar_barra_mensagem_enviar(driver, mensagem):
                                if cancelar:
                                    atualizar_log("Processamento cancelado!", cor="azul")
                                    return

                                atualizar_log(f"\nAviso enviado para {nome_contato}, {codigo} - {empresa}.\n", cor="verde")
                            else:
                                atualizar_log(f"\nFalha ao enviar mensagem para {nome_contato}\n", cor="vermelho")                       
                                
                            time.sleep(5)  # Espera um pouco antes de processar o próximo contato
                                
                            # Clica no botão para voltar à lista de contatos
                            if not clicar_voltar_lista_contatos(driver):
                                if cancelar:
                                    atualizar_log("Processamento cancelado!", cor="azul")
                                    return
                                
                                atualizar_log("Falha ao voltar para a lista de contatos. Tentando continuar...", cor="vermelho")
                                
                                time.sleep(5)  # Espera adicional após voltar à lista de contatos
                                
                            else:
                                atualizar_log(f"Iniciando novo contato...", cor="azul")
                    
                    except TimeoutException:
                                if cancelar:
                                    atualizar_log("Processamento cancelado!", cor="azul")
                                    return 
                                atualizar_log("\nBug ao tentar clicar no contato na Aba Contato!\n", cor="vermelho")
                                atualizar_log("Clicando na Aba Grupo para solucionar o problema", cor="azul")
                                # focar na aba grupos dos contatos
                                elemento = WebDriverWait(driver, 10).until(
                                    EC.element_to_be_clickable((By.XPATH, '//*[@id="react-tabs-2"]'))
                                )
                                
                                elemento.click()  
                                atualizar_log("Clicando no contato em Grupo.")
                                processar_resultados_busca(driver)     
                                
                                atualizar_log(f"carregamento Barra de Mensagens 6 segundos ...")
                                # Aguarda a página do chat carregar
                                time.sleep(6)
                                    
                                # Foca na barra de mensagem e envia
                                if focar_barra_mensagem_enviar(driver, mensagem):
                                    if cancelar:
                                        atualizar_log("Processamento cancelado!", cor="azul")
                                        return
                                    
    
                                    atualizar_log(f"\nAviso enviado para {nome_contato}, {codigo} - {empresa}.\n", cor="verde")
                                    
                                else:
                                    atualizar_log(f"Falha ao enviar mensagem para {nome_contato}", cor="vermelho")                       
                                    
                                time.sleep(5)  # Espera um pouco antes de processar o próximo contato
                                    
                                # Clica no botão para voltar à lista de contatos
                                if not focar_pagina(driver):
                                    atualizar_log("Falha ao voltar para a lista de contatos. Tentando continuar...", cor="vermelho")
                                    
                                    time.sleep(5)  # Espera adicional após voltar à lista de contatos
                                    
                                else:
                                    atualizar_log(f"Iniciando novo contato...", cor="azul")
            #GRUPO   
            elif nome_grupo.upper() != "NONE":
                if cancelar:
                    atualizar_log("Processamento cancelado!", cor="azul")
                    return 
                # Encontra a barra de Pesquisa e processa o contato
                if encontrar_e_clicar_barra_contatos(driver, nome_contato, nome_grupo):
                    
                    atualizar_log(f"Contato {nome_grupo} na Aba Grupo, encontrado e clicado.")
                    
                    try:
                        
                        # Espera pela TITULO area da mensagem
                        area_mensagem = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="ChatHeader"]/div[1]'))
                        )
                        
                        if area_mensagem:
                            
                            atualizar_log(f"carregamento Barra de Mensagens 6 segundos ...")
                            # Aguarda a página do chat carregar
                            time.sleep(6)

                            # Foca na barra de mensagem e envia
                            if focar_barra_mensagem_enviar(driver, mensagem):
                                atualizar_log(f"\nAviso enviado para {nome_grupo}, {codigo} - {empresa}.\n", cor="verde")
                            else:
                                atualizar_log(f"Falha ao enviar mensagem para {nome_grupo}", cor="vermelho")
                                
                            time.sleep(5)  # Espera um pouco antes de processar o próximo contato
                                
                            # Clica no botão para voltar à lista de contatos
                            if not focar_pagina(driver):
                                atualizar_log("Falha ao voltar para a lista de contatos. Tentando continuar...", cor="vermelho")
                                
                                time.sleep(5)  # Espera adicional após voltar à lista de contatos
                                
                            else:
                                atualizar_log(f"Iniciando novo contato...", cor="azul")
                    except TimeoutException: 
                        if cancelar:
                            atualizar_log("Processamento cancelado!", cor="azul")
                            return 
                        atualizar_log("Bug ao tentar clicar no contato na Aba Grupo!", cor="vermelho")
                        atualizar_log("Clicando na Aba Contato para solucionar o problema", cor="azul")
                        # Clicar aba Contato
                        elemento = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, '//*[@id="react-tabs-0"]'))
                        )
                        elemento.click()
                        atualizar_log("Clicando no contato em Grupo.")
                        processar_resultados_busca(driver)     
                        
                        atualizar_log(f"carregamento Barra de Mensagens 6 segundos ...")
                        # Aguarda a página do chat carregar
                        time.sleep(6)
                        
                        # Foca na barra de mensagem e envia
                        if focar_barra_mensagem_enviar(driver, mensagem):
                            atualizar_log(f"\nAviso enviado para {nome_grupo}, {codigo} - {empresa}.\n", cor="verde")
                        else:
                            atualizar_log(f"Falha ao enviar mensagem para {nome_grupo}", cor="vermelho")
                            
                        time.sleep(5)  # Espera um pouco antes de processar o próximo contato
                            
                        # Clica no botão para voltar à lista de contatos
                        if not clicar_voltar_lista_contatos(driver):
                            atualizar_log("Falha ao voltar para a lista de contatos. Tentando continuar...", cor="vermelho")
                            
                            time.sleep(5)  # Espera adicional após voltar à lista de contatos
                            
                        else:
                            atualizar_log(f"Iniciando novo contato...", cor="azul")
            else:
                atualizar_log("\nContato Inexistente!", cor="vermelho")
                atualizar_log(f"Pulando pessoa {codigo} - {empresa}", cor="azul")
                atualizar_log("Pulando em 3 Seg...")
                time.sleep(3)   
                
                            
            

        atualizar_log("Processamento finalizado com sucesso!")
        finalizar_programa()
    except Exception as e:
        atualizar_log(f"Ocorreu um erro inesperado: {str(e)}", cor="vermelho")
    

    
    
    
# Função para cancelar o processamento
def cancelar_processamento():
    global cancelar
    cancelar = True
    atualizar_log("Cancelando processamento...", cor="azul")
    botao_fechar.config(state=tk.NORMAL)  # Habilitar o botão de fechar o programa

# Função para cancelar e fechar o programa
def fechar_programa():
    janela.quit()

# Função para finalizar o programa com uma mensagem
def finalizar_programa():
    messagebox.showinfo("Processo Finalizado", "O processamento foi concluído com sucesso!")
    botao_fechar.config(state=tk.NORMAL)  # Habilitar o botão de fechar o programa
    botao_iniciar.config(state=tk.NORMAL)  # Reabilitar o botão de iniciar

# Função para atualizar o log na área de texto
def atualizar_log(mensagem, cor=None):
    log_text.config(state=tk.NORMAL)  # Habilitar edição temporária
    if cor == "vermelho":
        log_text.insert(tk.END, mensagem + "\n", "vermelho")  # Inserir nova mensagem com tag 'vermelho'
    elif cor == "verde":
        log_text.insert(tk.END, mensagem + "\n", "verde")  # Inserir nova mensagem com tag 'verde'
    elif cor == "azul":
        log_text.insert(tk.END, mensagem + "\n", "azul")
    else:
        log_text.insert(tk.END, mensagem + "\n")  # Inserir nova mensagem sem tag
    log_text.config(state=tk.DISABLED)  # Desabilitar edição novamente
    log_text.see(tk.END)  # Scroll automático para a última linha

# Função para configurar a tag de cor no log
def configurar_tags_log():
    log_text.tag_config("vermelho", foreground="red")  # Configura a cor vermelha para a tag 'vermelho'
    log_text.tag_config("verde", foreground="green")  # Configura a cor vermelha para a tag 'verde'
    log_text.tag_config("azul", foreground="blue") # Configura a cor azul para a tag 'azul'
# Função main para encapsular a lógica do programa
def main():
    global janela, caminho_pasta, caminho_excel, botao_fechar, botao_iniciar, log_text

    

    # Criar a janela principal
    janela = tk.Tk()
    janela.title("Prorrogação Contrato Experiência AutoMessenger")
    janela.geometry("600x400")
    janela.resizable(False, False)

    # Variáveis para armazenar os caminhos
    caminho_pasta = tk.StringVar()
    caminho_excel = tk.StringVar()

    # Frame para seleção de pasta e arquivo
    frame_selecao = tk.Frame(janela)
    frame_selecao.pack(pady=10)


    # Label e Botão para selecionar o arquivo Excel
    label_excel = tk.Label(frame_selecao, text="Arquivo Excel:")
    label_excel.grid(row=1, column=0, pady=5, padx=5)

    entrada_excel = tk.Entry(frame_selecao, textvariable=caminho_excel, width=50, state='readonly')
    entrada_excel.grid(row=1, column=1, padx=5)

    botao_excel = tk.Button(frame_selecao, text="Selecionar Excel", command=selecionar_excel)
    botao_excel.grid(row=1, column=2, padx=5)

    # Botão para iniciar o processamento
    botao_iniciar = tk.Button(janela, text="Iniciar Processamento", command=iniciar_processamento)
    botao_iniciar.pack(pady=10)

    # Botão para cancelar e fechar o programa
    botao_cancelar = tk.Button(janela, text="Cancelar Processamento", command=cancelar_processamento)
    botao_cancelar.pack(pady=5)

    # Botão para fechar o programa (desabilitado até o processamento terminar)
    botao_fechar = tk.Button(janela, text="Fechar Programa", command=fechar_programa, state=tk.DISABLED)
    botao_fechar.pack(pady=5)

    # Frame para o log
    frame_log = tk.Frame(janela)
    frame_log.pack(pady=10, fill=tk.BOTH, expand=True)

    # Área de texto para o log com barra de rolagem
    log_text = scrolledtext.ScrolledText(frame_log, wrap=tk.WORD, height=10, state=tk.DISABLED)
    log_text.pack(fill=tk.BOTH, expand=True)

    # Configurar as tags de cor para o log
    configurar_tags_log()
    
    # Iniciar o loop da interface
    janela.mainloop()

# Garantir que o código só execute se este arquivo for o principal
if __name__ == '__main__':
    main()
