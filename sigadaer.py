import time
import os
import urllib3

# --- CORREÇÃO DO ERRO DE SSL (REDE CORPORATIVA) ---
# Isso força o Python a ignorar o erro de certificado ao baixar o Driver
os.environ['WDM_SSL_VERIFY'] = '0'
# Desabilita avisos de conexão insegura no console para limpar a saída
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
chrome_options.add_argument('--ignore-certificate-errors') # Ignora erro de SSL do site SIGADAER
chrome_options.add_argument('--allow-running-insecure-content')
chrome_options.add_argument('--start-maximized')

prefs = {
    "download.default_directory": PASTA_DOWNLOAD,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True
}
chrome_options.add_experimental_option("prefs", prefs)

# Inicializa o Navegador (Agora com o erro de SSL tratado)
print("Verificando versão do Chrome Driver (isso pode demorar um pouco na primeira vez)...")
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
    wait.until(EC.presence_of_element_located((By.ID, "username")))
    driver.find_element(By.ID, "username").send_keys(USUARIO)
    
    # Tenta senha pelo tipo
    driver.find_element(By.XPATH, "//input[@type='password']").send_keys(SENHA)
    
    # Botão Entrar
    driver.find_element(By.XPATH, "//span[text()='Entrar'] | //input[@value='Entrar']").click()

    # ---------------------------------------------------------
    # 2. FILTRO MANUAL
    # ---------------------------------------------------------
    print("\n--- LOGIN EFETUADO ---")
    print("O script aguardará 45 segundos.")
    print(">> AÇÃO NECESSÁRIA: Vá em 'Documentos do Setor', filtre '2025' e clique em 'Aplicar'.")
    print(">> NÃO ABRA NENHUM DOCUMENTO AINDA.")
    
    time.sleep(45) 

    # ---------------------------------------------------------
    # 3. LOOP DE DOWNLOAD
    # ---------------------------------------------------------
    print("Iniciando varredura...")

    # Pega o número de linhas na tabela
    linhas_tabela = driver.find_elements(By.XPATH, "//tbody[contains(@id, 'data')]//tr")
    total_docs = len(linhas_tabela)
    print(f"Encontrados {total_docs} documentos.")

    for i in range(total_docs):
        try:
            print(f"\nProcessando {i+1}/{total_docs}...")

            # Atualiza a lista para evitar erro de elemento velho
            linhas_atuais = driver.find_elements(By.XPATH, "//tbody[contains(@id, 'data')]//tr")
            if i >= len(linhas_atuais): break # Previne erro se a lista mudar
                
            linha = linhas_atuais[i]
            
            # Clica no primeiro link da linha (geralmente o Número Sigad)
            link_doc = linha.find_element(By.XPATH, ".//a")
            link_doc.click()
            
            # Espera botão Download aparecer
            # XPATH específico do seu print
            xpath_btn = "//button[.//span[text()='Download']]"
            botao_download = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_btn)))
            
            # Garante que o botão está visível
            driver.execute_script("arguments[0].scrollIntoView();", botao_download)
            time.sleep(1) 

            # Clica
            botao_download.click()
            time.sleep(3) # Espera iniciar download
            
            # Volta para lista
            driver.back()
            
            # Espera a tabela recarregar
            wait.until(EC.presence_of_element_located((By.XPATH, "//tbody[contains(@id, 'data')]")))
            time.sleep(1.5) 
            
        except Exception as e:
            print(f"Erro no doc {i+1}: {e}")
            try: driver.back(); time.sleep(2) # Tenta recuperar se travar
            except: pass

    print(f"\nFinalizado! Pasta: {PASTA_DOWNLOAD}")

except Exception as e:
    print(f"Erro Fatal: {e}")

finally:
    driver.quit()