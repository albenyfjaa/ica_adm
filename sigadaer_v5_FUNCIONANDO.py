import time
import os
import urllib3

# --- CORREÇÃO SSL ---
os.environ['WDM_SSL_VERIFY'] = '0'
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
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
# FUNÇÃO: VERIFICAR SE O ARQUIVO JÁ EXISTE
# =============================================================================
def arquivo_ja_baixado(texto_link, pasta):
    """
    Verifica se existe algum arquivo na pasta que contenha o número do documento.
    O SIGADAER troca '/' por '_' no nome do arquivo.
    """
    if not texto_link: return False
    # Sanitiza o texto do link para o formato de arquivo (troca / por _)
    termo_busca = texto_link.strip().replace('/', '_').replace('\\', '_')
    arquivos_locais = os.listdir(pasta)
    for arquivo in arquivos_locais:
        if termo_busca in arquivo:
            return True 
    return False

# =============================================================================
# FUNÇÃO: PREPARAR A LISTA (FILTRO 2025 + VISUALIZAÇÃO 1000)
# =============================================================================
def preparar_lista_1000(driver, wait):
    print(" > Preparando ambiente...")
    
    # 1. Aplicar Filtro 2025 (Se necessário)
    xpath_select_ano = "//select[.//option[@value='2025']]"
    try:
        # Tenta achar o filtro. Se não achar, clica no menu para resetar.
        try:
            select_element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath_select_ano)))
        except:
            print(" > Menu não detectado, clicando em 'Documentos do Setor'...")
            driver.find_element(By.LINK_TEXT, "Documentos do Setor").click()
            select_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath_select_ano)))

        # Aplica 2025
        select_ano = Select(select_element)
        # Verifica se já está selecionado para ganhar tempo
        if select_ano.first_selected_option.get_attribute("value") != "2025":
            print(" > Selecionando Ano 2025...")
            select_ano.select_by_value("2025")
            btn = driver.find_element(By.XPATH, "//button[contains(@id, 'aplicarFiltrosButton')]")
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(3) # Espera recarregar tabela inicial
        else:
            print(" > Ano 2025 já estava selecionado.")

    except Exception as e:
        print(f" > Erro ao filtrar ano: {e}")
        return False

    # 2. Selecionar 1000 Itens por Página (NOVO)
    print(" > Alterando visualização para 1000 itens...")
    try:
        # ID identificado na imagem: caixaForm:tableDocumentos_rppDD
        id_select_paginacao = "caixaForm:tableDocumentos_rppDD"
        
        # Espera o select de paginação aparecer (ele fica no rodapé da tabela)
        sel_pag_element = wait.until(EC.presence_of_element_located((By.ID, id_select_paginacao)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", sel_pag_element)
        
        select_pag = Select(sel_pag_element)
        
        # Só altera se não estiver em 1000 ainda
        if select_pag.first_selected_option.text != "1000":
            select_pag.select_by_value("1000") # Valor baseado na imagem (10, 25, 50, 1000)
            print(" > Aguardando 5 segundos para atualização da tabela...")
            time.sleep(5) # Pausa obrigatória solicitada
        else:
            print(" > Visualização já está em 1000.")

        return True

    except Exception as e:
        print(f" > Erro ao selecionar 1000 itens: {e}")
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
        print("Login automático pulado.")

    print("\n--- ACESSANDO MENU ---")
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Documentos do Setor"))).click()
    
    # Prepara a lista (2025 + 1000 itens)
    preparar_lista_1000(driver, wait)

    # LOOP ÚNICO (SEM PAGINAÇÃO)
    print(f"\n==============================================")
    print(f" INICIANDO VARREDURA (MODO 1000 ITENS)")
    print(f"==============================================")

    # Captura todas as linhas de uma vez
    xpath_linhas = "//tbody[contains(@id, 'data')]//tr[@data-ri]"
    wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath_linhas)))
    linhas = driver.find_elements(By.XPATH, xpath_linhas)
    total_docs = len(linhas)
    
    print(f" > Total de documentos listados: {total_docs}")
    janela_principal = driver.current_window_handle

    for i in range(total_docs):
        try:
            # Recaptura leve para garantir integridade (caso o DOM mude minimamente)
            # Como não recarregamos a página, isso é rápido
            linhas = driver.find_elements(By.XPATH, xpath_linhas)
            if i >= len(linhas): break
            
            linha = linhas[i]
            link_doc = linha.find_element(By.XPATH, ".//a")
            texto_link = link_doc.text.strip()
            
            # --- VERIFICAÇÃO DE JÁ BAIXADO ---
            if arquivo_ja_baixado(texto_link, PASTA_DOWNLOAD):
                print(f" [{i+1}/{total_docs}] JÁ EXISTE: {texto_link}")
                continue
            # ---------------------------------

            print(f" [{i+1}/{total_docs}] BAIXANDO: {texto_link}...")

            # Abre em NOVA ABA (Essencial para não perder a view de 1000 itens)
            ActionChains(driver).key_down(Keys.CONTROL).click(link_doc).key_up(Keys.CONTROL).perform()
            
            wait.until(EC.number_of_windows_to_be(2))
            janelas = driver.window_handles
            nova_janela = [w for w in janelas if w != janela_principal][0]
            driver.switch_to.window(nova_janela)
            
            # Download na nova aba
            try:
                xpath_btn_download = "//button[.//span[text()='Download']]"
                btn_down = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_btn_download)))
                driver.execute_script("arguments[0].click();", btn_down)
                time.sleep(3) # Tempo para iniciar download
            except TimeoutException:
                print("     > ERRO: Botão não encontrado.")
            
            # Fecha a aba e volta imediatamente
            driver.close()
            driver.switch_to.window(janela_principal)
            
        except Exception as e:
            print(f"     > Erro no item {i+1}: {e}")
            # Recuperação de erro de janela
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(janela_principal)

    print(f"\nScript Finalizado! Pasta: {PASTA_DOWNLOAD}")

except Exception as e:
    print(f"Erro Fatal: {e}")