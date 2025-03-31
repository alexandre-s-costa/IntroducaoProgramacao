from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd
import datetime
import time
import os

def capturar_dados_processo(driver):
    #Captura os dados de um processo na página de detalhes.
    dados = {}  
    # Mapeamento de campos e seus XPaths
    campos = {
        "Número": (By.ID, "span_proc_numero"),
        "Área": (By.XPATH, "//div[contains(text(), 'Área')]/following-sibling::span"),
        "Classe": (By.XPATH, "//div[contains(text(), 'Classe')]/following-sibling::span"),
        "Assunto(s)": (By.XPATH, "//div[contains(text(), 'Assunto(s)')]/following-sibling::span"),
        "Valor da Causa": (By.XPATH, "//div[contains(text(), 'Valor da Causa')]/following-sibling::span"),
        "Fase Processual": (By.XPATH, "//div[contains(text(), 'Fase Processual')]/following-sibling::span"),
        "Dt. Distribuição": (By.XPATH, "//div[contains(text(), 'Dt. Distribui')]/following-sibling::span"),
        "Custas": (By.XPATH, "//div[contains(text(), 'Custas')]/following-sibling::span"),
        "Status": (By.XPATH, "//div[contains(text(), 'Status')]/following-sibling::span")
    }
    # Extrair campos básicos
    for chave, (by, path) in campos.items():
        try:
            elemento = WebDriverWait(driver, 3).until(EC.presence_of_element_located((by, path)))
            dados[chave] = elemento.text.strip()
        except:
            dados[chave] = "Não encontrado"          
    
    # Capturar Polo Ativo
    try:
        polo_ativo_legend = driver.find_element(By.XPATH, "//legend[contains(text(), 'Polo Ativo')]")
        fieldset = polo_ativo_legend.find_element(By.XPATH, "./ancestor::fieldset")
        nomes_ativos = fieldset.find_elements(By.XPATH, ".//span[@class='span1 nomes']")
        
        if nomes_ativos:
            dados["Polo Ativo"] = "; ".join([nome.text.strip() for nome in nomes_ativos])
        else:
            dados["Polo Ativo"] = "Não encontrado"
    except:
        dados["Polo Ativo"] = "Não encontrado"
    
    # Capturar Polo Passivo
    """
    try:
        polo_passivo_legend = driver.find_element(By.XPATH, "//legend[contains(text(), 'Polo Passivo | Executado')]")
        fieldset = polo_passivo_legend.find_element(By.XPATH, "./ancestor::fieldset")
        nomes_passivos = fieldset.find_elements(By.XPATH, ".//span[@title='Nome da Parte']")
        
        if nomes_passivos:
            dados["Polo Passivo"] = "; ".join([nome.get_attribute('title').strip() for nome in nomes_passivos])
        else:
            dados["Polo Passivo"] = "Não encontrado"
    except:
        dados["Polo Passivo"] = "Não encontrado"
   """  
    return dados
   
def processar_tabela(driver, tabela_xpath):
    # Processa uma tabela de resultados, extraindo dados de cada processo.
    dados_processos = []
    
    # Obter a quantidade de linhas
    linhas = driver.find_elements(By.XPATH, f"{tabela_xpath}//tr[contains(@class, 'TabelaLinha')]")
    quantidade_linhas = len(linhas)
    
    for i in range(quantidade_linhas):
        try:
            # Refaz a busca das linhas para evitar "stale elements"
            linhas = driver.find_elements(By.XPATH, f"{tabela_xpath}//tr[contains(@class, 'TabelaLinha')]")
            if i >= len(linhas):
                break
                
            linha = linhas[i]
            
            # Clica no botão "Editar"
            botao_editar = WebDriverWait(linha, 5).until(
                EC.element_to_be_clickable((By.XPATH, './/i[contains(@class, "fa-pen-to-square")]'))
            )
            botao_editar.click()
            
            # Espera a página de detalhes carregar
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "span_proc_numero")))
            time.sleep(1)
            
            # Captura os dados e adiciona à lista
            dados = capturar_dados_processo(driver)
            dados_processos.append(dados)
            
            # Volta para a página de resultados
            driver.back()
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//table')))
            time.sleep(1)
            
        except Exception as e:
            print(f"Erro ao processar a linha {i}: {e}")
            try:
                driver.back()
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//table')))
            except:
                print("Não foi possível retornar à página de resultados")
    
    return dados_processos

def processar_paginas(driver):
    # Processa todas as páginas de resultados.
    dados_processos = []
    
    # Capturar o total de processos informado
    try:
        total_element = driver.find_element(By.XPATH, "//div[@id='Paginacao']/b[contains(text(), 'Total de:')]")
        total_processos = int(total_element.text.split('Total de:')[1].strip())
        print(f"Total de processos informado na página: {total_processos}")
    except:
        print("Não foi possível capturar o total de processos")
        total_processos = None
    
    pagina_atual = 1
    ultima_pagina = False
    
    while not ultima_pagina:
        print(f"Processando a página {pagina_atual}...")
        
        # Processar primeira tabela
        dados_processos += processar_tabela(driver, "//table[1]")
        
        # Verificar se existe uma segunda tabela
        try:
            botao_tabela2 = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Tabela 2')]"))
            )
            botao_tabela2.click()
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//table')))
            dados_processos += processar_tabela(driver, "//table[2]")
        except:
            print("Segunda tabela não encontrada ou vazia. Possivelmente é a última página.")
        
        # Tentar navegar para a próxima página
        try:
            # Método 1: Link direto para a próxima página
            try:
                proxima_pagina = pagina_atual + 1
                xpath_proxima = f"//div[@id='Paginacao']/a[contains(text(), '{proxima_pagina}')]"
                botao_proxima = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, xpath_proxima))
                )
                botao_proxima.click()
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//table')))
                pagina_atual = proxima_pagina
            except:
                ultima_pagina = True    
        except:
            ultima_pagina = True
    
    return dados_processos, total_processos

def consulta_por_nome(driver, nome_parte, numero=None, cpf_cnpj=None):
    # Realiza consulta por nome da parte.
    url = 'https://projudi.tjgo.jus.br/BuscaProcesso?PaginaAtual=4'
    driver.get(url)
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, "NomeParte")))
    
    # Preencher campos de busca
    campo_nome_parte = driver.find_element(By.NAME, "NomeParte")
    campo_nome_parte.clear()
    campo_nome_parte.send_keys(nome_parte)
    
    if numero:
        campo_processo_numero = driver.find_element(By.NAME, "ProcessoNumero")
        campo_processo_numero.clear()
        campo_processo_numero.send_keys(numero)
    
    if cpf_cnpj:
        campo_cpf_cnpj = driver.find_element(By.NAME, "CpfCnpjParte")
        campo_cpf_cnpj.clear()
        campo_cpf_cnpj.send_keys(cpf_cnpj)
    
    # Submeter consulta
    botao_consulta = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.NAME, "imgSubmeter"))
    )
    botao_consulta.click()
    
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//table')))
    
    return processar_paginas(driver)

def consulta_por_arquivo(driver, arquivo_processos):
    # Realiza consulta a partir de um arquivo com números de processo.
    # Ler números de processo do arquivo
    try:
        df_processos = pd.read_excel(arquivo_processos)
        print(f"Encontrados {len(df_processos)} processos para consulta.")
        
        if ('PROCESSO' or 'processo' or 'Processos' or 'processos') not in df_processos.columns:
            print("A coluna 'PROCESSO' não foi encontrada no arquivo.")
            return [], None
        
        dados_todos_processos = []
        
        # Consultar cada processo individualmente
        for i, processo in enumerate(df_processos['PROCESSO']):
            print(f"Processando {i+1}/{len(df_processos)}: {processo}")
            
            # Acessar a página de detalhes do processo
            url = f"https://projudi.tjgo.jus.br/BuscaProcesso?PaginaAtual=2&ProcessoNumero={processo}"
            driver.get(url)
            
            try:
                # Verificar se a página de detalhes carregou
                WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "span_proc_numero")))
                dados = capturar_dados_processo(driver)
                dados_todos_processos.append(dados)
            except:
                print(f"Erro ao acessar o processo {processo}. Verifique se o número está correto no arquivo.")
                dados_todos_processos.append({"Número": processo, "Erro": "Processo não encontrado"})
            
            time.sleep(1)
        
        return dados_todos_processos, len(df_processos)
    
    except Exception as e:
        print(f"Erro ao processar o arquivo: {e}")
        return [], None

def Consulta(tipo_consulta, **kwargs):
    hora_inicio = datetime.datetime.now()
    
    # Configuração do driver
    chrome_options = Options()
    chrome_options.add_argument('--disable-dev-shm-usage')
    #chrome_options.add_argument('--headless')  # Descomente para executar sem abrir o navegador
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    try:
        if tipo_consulta.lower() == 'nome':
            nome_parte = kwargs.get('NomeParte')
            if not nome_parte:
                print("Nome da parte é obrigatório para consulta por nome.")
                driver.quit()
                return None
            
            numero = kwargs.get('numero')
            cpf_cnpj = kwargs.get('CPF_CNPJ')
            
            dados_processos, total_processos = consulta_por_nome(driver, nome_parte, numero, cpf_cnpj)
        
        elif tipo_consulta.lower() == 'arquivo':
            arquivo = kwargs.get('arquivo')
            if not arquivo or not os.path.exists(arquivo):
                print("Arquivo não encontrado ou não especificado.")
                driver.quit()
                return None
            
            dados_processos, total_processos = consulta_por_arquivo(driver, arquivo)
        
        else:
            print("Tipo de consulta inválido. Use 'nome' ou 'arquivo'.")
            driver.quit()
            return None
        
        driver.quit() # Fecha a janela do Chrome

        # Medição de tempo de execução
        hora_fim = datetime.datetime.now()
        tempo_execucao = hora_fim - hora_inicio
        print(f"Tempo de execução: {tempo_execucao}")
        
        # Criar DataFrame e salvar resultados
        df = pd.DataFrame(dados_processos)
        
        # Validar total de registros
        registros_coletados = len(df)
        print(f"Registros coletados: {registros_coletados}")
        
        if total_processos is not None and registros_coletados != total_processos:
            print(f"ATENÇÃO: Registros coletados ({registros_coletados}) diferente do total informado ({total_processos}).")
        
        # Verificar caminho atual para salvar o arquivo
        #caminho_arquivo = kwargs.get('C:/Users/asilvacosta/Documents/PRTI/PROJETOS/PROJETO PROGRAMACAO')      
        caminho_arquivo = kwargs.get('C:/Users/xanda/OneDrive/Documentos/RESIDENCIA TI/INTRODUCAO PROGRAMACAO/')

        # Salvar em Excel
        nome_arquivo = f'Processos_TJGO_{hora_fim.strftime("%d-%m-%Y_%H-%M-%S")}.xlsx'
        df.to_excel(nome_arquivo, index=False)
        print(f"Dados exportados para {nome_arquivo}")
        
        return df
    
    except Exception as e:
        print(f"Erro durante a execução: {e}")
        driver.quit()
        return None
# =========================== #
#########  Função Main ########
# =========================== #
if __name__ == "__main__":
    # Interface simples de linha de comando
    print("Consulta de Processos do Projudi")
    print("1 - Consultar por nome da parte")
    print("2 - Consultar por arquivo de processos")
    
    opcao = input("Escolha uma opção (1 ou 2): ")
    # Consultar por Nome da Parte
    if opcao == "1":
        nome = input("Nome da parte: ")
        numero = input("Número do processo (opcional, pressione Enter para pular): ")
        cpf_cnpj = input("CPF/CNPJ (opcional, pressione Enter para pular): ")
        
        Consulta('nome', NomeParte=nome, 
                 numero=numero if numero else None, 
                 CPF_CNPJ=cpf_cnpj if cpf_cnpj else None)
    # Consultar por arquivo
    elif opcao == "2":
        arquivo = input("Caminho do arquivo Excel com a coluna 'PROCESSO': ")
        Consulta('arquivo', arquivo=arquivo)
    
    else:
        print("Opção inválida.")