import time
import os
import urllib3

# --- CORREÇÃO DO ERRO DE SSL (REDE CORPORATIVA) ---
os.environ['WDM_SSL_VERIFY'] = '0'
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select  # <--- IMPORT NECESSÁRIO PARA O FILTRO
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURAÇÕES DE USUÁRIO ---
USUARIO = 'albeny-fjaa'
SENHA = 'Fj@@0788'


# URLs e Pastas
URL_LOGIN = "https://10.32.24.172:8444/sigad-ica/pages/login/login.jsf"
PASTA_DOWNLOAD = os.path.join(os.getcwd(), "downloads_sigadaer_2025")

if not os.path.exists(PASTA_DOWNLOAD):
    os.makedirs(PASTA_DOWNLOAD)

# --- CONFIGURAÇÃO DO CHROME ---
chrome_options = Options()
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument('--allow-running-insecure-content')
chrome_options.add_argument('--start-maximized')

prefs = {
    "download.default_directory": PASTA_DOWNLOAD,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True
}
chrome_options.add_experimental_option("prefs", prefs)

# Inicializa o Navegador
print("Verificando versão do Chrome Driver...")
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(driver, 20)

try:
    # ---------------------------------------------------------
    # 1. LOGIN 
    # ---------------------------------------------------------
    print("Acessando SIGADAER...")
    driver.get(URL_LOGIN)

    print("Realizando login...")
    # Verifica se os campos existem antes de tentar interagir
    try:
        wait.until(EC.presence_of_element_located((By.ID, "username")))
        driver.find_element(By.ID, "username").send_keys(USUARIO)
        driver.find_element(By.XPATH, "//input[@type='password']").send_keys(SENHA)
        driver.find_element(By.XPATH, "//span[text()='Entrar'] | //input[@value='Entrar']").click()
    except:
        print("Login automático pulado ou já logado (verifique se precisa logar manualmente).")

    # ---------------------------------------------------------
    # 2. AUTOMATIZAÇÃO DO FILTRO (CORRIGIDO)
    # ---------------------------------------------------------
    print("\n--- NAVEGANDO PARA DOCUMENTOS DO SETOR ---")
    
    # Clica no menu lateral "Documentos do Setor"
    link_setor = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Documentos do Setor")))
    link_setor.click()
    
    print("Aguardando carregamento da tela (5 segundos)...")
    time.sleep(5)  # PAUSA OBRIGATÓRIA PARA O CARREGAMENTO DO JSF
    
    print("Selecionando Ano 2025...")
    
    # ESTRATÉGIA ROBUSTA: Encontrar o SELECT que possui a opção '2025' dentro dele
    # Isso evita erros se o ID mudar um pouco
    xpath_select_ano = "//select[.//option[@value='2025']]"
    
    try:
        select_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath_select_ano)))
        
        # Garante que o elemento está visível na tela (Scroll até ele)
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_element)
        time.sleep(1)

        select_ano = Select(select_element)
        select_ano.select_by_value("2025")
        print("Ano 2025 selecionado com sucesso.")
        
    except Exception as e:
        print(f"Não foi possível encontrar o filtro de ano: {e}")
        # Tenta imprimir o HTML da página para debug se falhar (opcional)
        # print(driver.page_source)
        raise e # Para o script aqui se não conseguir filtrar

    print("Aplicando Filtros...")
    # Procura o botão que contenha o texto "Aplicar Filtros" ou parte do ID
    xpath_botao_aplicar = "//button[contains(@id, 'aplicarFiltrosButton')]"
    
    botao_aplicar = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_botao_aplicar)))
    driver.execute_script("arguments[0].click();", botao_aplicar) # Click via JS é mais garantido no PrimeFaces
    
    print("Aguardando atualização da tabela...")
    time.sleep(5) # Tempo para a tabela recarregar

    # ---------------------------------------------------------
    # 3. LOOP DE DOWNLOAD
    # ---------------------------------------------------------
    print("Iniciando varredura de downloads...")

    # Pega o número de linhas na tabela
    # Procura tbody que contenha 'data' no ID (comum em PrimeFaces: form:dataTable_data)
    wait.until(EC.presence_of_element_located((By.XPATH, "//tbody[contains(@id, 'data')]//tr")))
    linhas_tabela = driver.find_elements(By.XPATH, "//tbody[contains(@id, 'data')]//tr")
    total_docs = len(linhas_tabela)
    
    # Se encontrar apenas 1 linha e ela disser "Nenhum registro", ajusta a contagem
    if total_docs == 1 and "Nenhum" in linhas_tabela[0].text:
        print("Nenhum documento encontrado para 2025.")
        total_docs = 0
    else:
        print(f"Encontrados {total_docs} documentos.")

    for i in range(total_docs):
        try:
            print(f"\nProcessando {i+1}/{total_docs}...")

            # Atualiza a referência da lista
            linhas_atuais = driver.find_elements(By.XPATH, "//tbody[contains(@id, 'data')]//tr")
            if i >= len(linhas_atuais): 
                print("Índice fora do alcance (tabela mudou?), parando.")
                break
                
            linha = linhas_atuais[i]
            
            # Clica no primeiro link da linha
            link_doc = linha.find_element(By.XPATH, ".//a")
            link_doc.click()
            
            # Espera botão Download aparecer
            xpath_btn = "//button[.//span[text()='Download']]"
            botao_download = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_btn)))
            
            # Scroll se necessário
            driver.execute_script("arguments[0].scrollIntoView();", botao_download)
            time.sleep(1) 

            # Clica
            botao_download.click()
            time.sleep(3) # Espera iniciar download
            
            # Volta para lista
            driver.back()
            
            # Espera a tabela recarregar antes de ir para o próximo
            wait.until(EC.presence_of_element_located((By.XPATH, "//tbody[contains(@id, 'data')]")))
            time.sleep(1.5) 
            
        except Exception as e:
            print(f"Erro no doc {i+1}: {e}")
            try: 
                driver.back()
                time.sleep(2) 
            except: 
                pass

    print(f"\nFinalizado! Pasta: {PASTA_DOWNLOAD}")

except Exception as e:
    print(f"Erro Fatal: {e}")

finally:
    # Comente a linha abaixo se quiser manter o navegador aberto para debug
    driver.quit()