import time
import os
import urllib3

# --- CORREÇÃO DO ERRO DE SSL ---
os.environ['WDM_SSL_VERIFY'] = '0'
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- CONFIGURAÇÕES ---
USUARIO = 'albeny-fjaa'  # Preencha
SENHA = 'Fj@@0788'    # Preencha
URL_LOGIN = "https://10.32.24.172:8444/sigad-ica/pages/login/login.jsf"
PASTA_DOWNLOAD = os.path.join(os.getcwd(), "downloads_sigadaer_2025")

if not os.path.exists(PASTA_DOWNLOAD):
    os.makedirs(PASTA_DOWNLOAD)

# --- CHROME OPTIONS ---
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

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(driver, 20)

# =============================================================================
# FUNÇÃO: RESTAURAR O ESTADO DA LISTA
# =============================================================================
def garantir_tabela_carregada(driver, wait):
    """
    Garante que estamos na tela da lista, reaplica o filtro se necessário
    e aguarda a tabela aparecer.
    """
    print(" > Verificando estado da tabela...")
    
    # 1. Verifica se o Select do ano está visível (sinal que a página carregou)
    xpath_select_ano = "//select[.//option[@value='2025']]"
    try:
        # Espera curta para ver se o elemento já está lá
        select_element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, xpath_select_ano))
        )
    except TimeoutException:
        print(" > Filtro não encontrado. Tentando recarregar a seção...")
        # Se não achou o filtro, talvez tenha caído fora da página. 
        # Clica no menu novamente como 'Reset'
        try:
            driver.find_element(By.LINK_TEXT, "Documentos do Setor").click()
            select_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath_select_ano)))
        except:
            return False # Falha crítica

    # 2. Verifica se a tabela tem dados reais
    # O seletor [@data-ri] pega SOMENTE linhas de dados do PrimeFaces
    xpath_linhas_dados = "//tbody[contains(@id, 'data')]//tr[@data-ri]"
    linhas = driver.find_elements(By.XPATH, xpath_linhas_dados)
    
    if len(linhas) > 0:
        print(" > Tabela já possui dados carregados.")
        return True

    print(" > Tabela vazia ou filtro perdido. Reaplicando filtro 2025...")
    
    # 3. Reaplicar Filtro
    try:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_element)
        select_ano = Select(select_element)
        select_ano.select_by_value("2025")
        
        xpath_botao = "//button[contains(@id, 'aplicarFiltrosButton')]"
        btn = driver.find_element(By.XPATH, xpath_botao)
        driver.execute_script("arguments[0].click();", btn)
        
        # Aguarda spinner de carregamento sumir (se houver) ou tabela aparecer
        time.sleep(3) 
        wait.until(EC.presence_of_element_located((By.XPATH, xpath_linhas_dados)))
        print(" > Filtro aplicado e tabela recarregada.")
        return True
    except Exception as e:
        print(f" > Erro ao reaplicar filtro: {e}")
        return False

# =============================================================================
# FLUXO PRINCIPAL
# =============================================================================
try:
    print("--- LOGIN ---")
    driver.get(URL_LOGIN)
    try:
        wait.until(EC.presence_of_element_located((By.ID, "username")))
        driver.find_element(By.ID, "username").send_keys(USUARIO)
        driver.find_element(By.XPATH, "//input[@type='password']").send_keys(SENHA)
        driver.find_element(By.XPATH, "//span[text()='Entrar'] | //input[@value='Entrar']").click()
    except:
        print("Login pulado.")

    print("\n--- ACESSANDO MENU ---")
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Documentos do Setor"))).click()
    time.sleep(2)

    # Preparação Inicial
    garantir_tabela_carregada(driver, wait)

    # Contagem Real (Usando @data-ri para ignorar linhas vazias/cabeçalhos)
    xpath_linhas = "//tbody[contains(@id, 'data')]//tr[@data-ri]"
    linhas_iniciais = driver.find_elements(By.XPATH, xpath_linhas)
    total_docs = len(linhas_iniciais)
    print(f"\nDOCUMENTOS REAIS ENCONTRADOS: {total_docs}")

    if total_docs == 0:
        print("Nenhum documento encontrado com o seletor rigoroso. Verifique o filtro.")
        exit()

    # LOOP
    for i in range(total_docs):
        print(f"\n--- Documento {i+1} / {total_docs} ---")
        
        try:
            # 1. Garantir que a lista está na tela
            garantir_tabela_carregada(driver, wait)
            
            # 2. Re-capturar a linha específica
            # Usamos wait.until para garantir que as linhas existam no DOM
            linhas_atuais = wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath_linhas)))
            
            if i >= len(linhas_atuais):
                print("ERRO: Índice fora de alcance. A paginação mudou?")
                break
            
            linha = linhas_atuais[i]
            
            # 3. Extrair Link e Clicar
            try:
                # Busca link dentro da linha específica
                link_doc = linha.find_element(By.XPATH, ".//a")
                print(f" > Abrindo documento: {link_doc.text[:30]}...")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", link_doc)
                time.sleep(0.5)
                link_doc.click() # Clica para entrar no detalhe
                
            except NoSuchElementException:
                print(" > AVISO: Linha sem link encontrada. Pulando...")
                print(f" > Conteúdo da linha: {linha.text}")
                continue

            # 4. Tela de Detalhe -> Download
            xpath_btn_download = "//button[.//span[text()='Download']]"
            try:
                # Espera o botão de download aparecer na nova tela
                btn_down = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_btn_download)))
                driver.execute_script("arguments[0].click();", btn_down)
                print(" > Download iniciado.")
                time.sleep(2) # Espera o download engatilhar
            except TimeoutException:
                print(" > Erro: Botão de download não encontrado na tela de detalhe.")
            
            # 5. Voltar para a lista
            print(" > Voltando para a lista...")
            driver.back()
            
            # Espera CRÍTICA para o navegador processar o 'Back' antes de tentar achar a tabela de novo
            time.sleep(2) 

        except Exception as e:
            print(f"Erro genérico no item {i}: {e}")
            driver.back()
            time.sleep(2)

    print("\nProcesso Finalizado!")

except Exception as e:
    print(f"Erro Fatal: {e}")