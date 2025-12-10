import time
import os
import urllib3
import re

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
USUARIO = 'albeny-fjaa'
SENHA = 'Fj@@0788'
URL_LOGIN = "https://10.32.24.172:8444/sigad-ica/pages/login/login.jsf"
PASTA_DOWNLOAD = os.path.join(os.getcwd(), "downloads_sigadaer_2024_v2")

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
# 1. FUNÇÃO: LIMPEZA DE NOME
# =============================================================================
def sanitizar_nome(nome):
    """Remove caracteres inválidos para Windows e espaços extras."""
    nome = re.sub(r'[<>:"/\\|?*]', '_', nome)
    return " ".join(nome.split())[:180]

# =============================================================================
# 2. FUNÇÃO: VERIFICAÇÃO RIGOROSA (COM EXCEÇÃO PARA S/N)
# =============================================================================
def arquivo_ja_baixado_rigoroso(texto_link, pasta):
    """
    Verifica se o arquivo existe, MAS FORÇA O DOWNLOAD se o ID for genérico (S/N).
    """
    if not texto_link: return False
    
    # --- NOVA LÓGICA: DETECTAR ID GENÉRICO ---
    # Se o ID for s/n, s.n., s/nº, o script vai ignorar a verificação 
    # e baixar sempre, evitando pular documentos diferentes que compartilham o 's/n'
    id_lower = texto_link.strip().lower()
    # Remove pontos, º, ª para padronizar (ex: "s.n.º" vira "sn")
    id_limpo = re.sub(r'[º°ª\.]', '', id_lower)
    
    if id_limpo in ['sn', 's/n', 's-n', 'sem numero', 'sem número']:
        return False # Retorna False para FORÇAR o download
    # -----------------------------------------

    # Sanitiza o ID buscado (ex: '51/1DCT' -> '51_1DCT')
    id_sanitizado = re.sub(r'[<>:"/\\|?*]', '_', texto_link.strip())
    
    if len(id_sanitizado) < 2: 
        return False
        
    id_proc = id_sanitizado.lower()
    
    try:
        arquivos_locais = os.listdir(pasta)
    except FileNotFoundError:
        return False
    
    for f in arquivos_locais:
        nome_arq = f.lower()
        
        if id_proc in nome_arq:
            idx_inicio = nome_arq.find(id_proc)
            idx_fim = idx_inicio + len(id_proc)
            
            if idx_inicio > 0:
                char_antes = nome_arq[idx_inicio - 1]
                if char_antes.isalnum(): continue

            if idx_fim < len(nome_arq):
                char_depois = nome_arq[idx_fim]
                if char_depois.isalnum(): continue
            
            return True
                
    return False

# =============================================================================
# 3. FUNÇÃO: MONITORAR E RENOMEAR
# =============================================================================
def monitorar_e_renomear(pasta, arquivos_antes, nome_desejado):
    print("     > Aguardando arquivo...")
    fim_tempo = time.time() + 60
    
    while time.time() < fim_tempo:
        arquivos_agora = set(os.listdir(pasta))
        novos = arquivos_agora - arquivos_antes
        
        for arquivo in novos:
            if arquivo.endswith('.crdownload') or arquivo.endswith('.tmp'):
                continue
                
            caminho_completo = os.path.join(pasta, arquivo)
            try:
                if os.path.getsize(caminho_completo) > 0:
                    # Se vier com nome genérico (download...), renomeia
                    if "download" in arquivo.lower():
                        novo_nome = f"{nome_desejado}.pdf"
                        dest = os.path.join(pasta, novo_nome)
                        
                        # Evita sobrescrever duplicatas adicionando (1), (2)
                        c = 1
                        while os.path.exists(dest):
                            dest = os.path.join(pasta, f"{nome_desejado} ({c}).pdf")
                            c += 1
                            
                        os.rename(caminho_completo, dest)
                        print(f"     > RENOMEADO PARA: {os.path.basename(dest)}")
                        return True
                    else:
                        print(f"     > Nome original mantido: {arquivo}")
                        return True
            except:
                pass
        time.sleep(1)
    return False

# =============================================================================
# 4. PREPARAÇÃO DA LISTA (1000 ITENS)
# =============================================================================
def preparar_lista_1000(driver, wait):
    xpath_select_ano = "//select[.//option[@value='2024']]"
    try:
        try:
            sel = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath_select_ano)))
        except:
            driver.find_element(By.LINK_TEXT, "Documentos do Setor").click()
            sel = wait.until(EC.presence_of_element_located((By.XPATH, xpath_select_ano)))

        select_ano = Select(sel)
        if select_ano.first_selected_option.get_attribute("value") != "2024":
            select_ano.select_by_value("2024")
            driver.find_element(By.XPATH, "//button[contains(@id, 'aplicarFiltrosButton')]").click()
            time.sleep(3)
            
        # Paginador 1000
        sel_pag = wait.until(EC.presence_of_element_located((By.ID, "caixaForm:tableDocumentos_rppDD")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", sel_pag)
        select_pag = Select(sel_pag)
        
        if select_pag.first_selected_option.text != "1000":
            select_pag.select_by_value("1000")
            print(" > Modo 1000 itens ativado. Aguardando 5s...")
            time.sleep(5)
        return True
    except Exception as e:
        print(f" > Erro preparação: {e}")
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
        pass # Já logado

    print("\n--- ACESSANDO MENU ---")
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Documentos do Setor"))).click()
    preparar_lista_1000(driver, wait)

    print("\n--- INICIANDO PROCESSAMENTO ---")
    xpath_linhas = "//tbody[contains(@id, 'data')]//tr[@data-ri]"
    wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath_linhas)))
    linhas = driver.find_elements(By.XPATH, xpath_linhas)
    total_docs = len(linhas)
    print(f" > Total encontrado: {total_docs}")

    janela_principal = driver.current_window_handle

    for i in range(total_docs):
        try:
            # Recaptura leve
            linhas = driver.find_elements(By.XPATH, xpath_linhas)
            if i >= len(linhas): break
            linha = linhas[i]
            link_doc = linha.find_element(By.XPATH, ".//a")
            
            numero_doc = link_doc.text.strip() # Ex: 3/CAR2
            
            # --- VERIFICAÇÃO RIGOROSA ---
            if arquivo_ja_baixado_rigoroso(numero_doc, PASTA_DOWNLOAD):
                print(f" [{i+1}/{total_docs}] Já existe: {numero_doc}")
                continue
            # ----------------------------

            print(f" [{i+1}/{total_docs}] Processando: {numero_doc}...")
            arquivos_antes = set(os.listdir(PASTA_DOWNLOAD))

            # Abre nova aba
            ActionChains(driver).key_down(Keys.CONTROL).click(link_doc).key_up(Keys.CONTROL).perform()
            wait.until(EC.number_of_windows_to_be(2))
            
            # Muda para nova aba
            janelas = driver.window_handles
            driver.switch_to.window([w for w in janelas if w != janela_principal][0])

            # --- DENTRO DA NOVA ABA ---
            try:
                # 1. ESPERA CRÍTICA DE CARREGAMENTO (Resolve o erro da aba de detalhe)
                # Espera até que o label 'Assunto:' esteja visível no DOM
                wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Assunto:')]")))
                
                # 2. CAPTURA ASSUNTO
                assunto_texto = ""
                try:
                    elem_assunto = driver.find_element(By.XPATH, "//*[contains(text(), 'Assunto:')]/..")
                    assunto_texto = elem_assunto.text.replace("Assunto:", "").strip()
                except:
                    pass

                # Monta nome
                nome_base = f"{numero_doc} - {assunto_texto}" if assunto_texto else numero_doc
                nome_final = sanitizar_nome(nome_base)

                # 3. DOWNLOAD
                xpath_btn = "//button[.//span[text()='Download']]"
                btn = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_btn)))
                driver.execute_script("arguments[0].click();", btn)
                
                # 4. MONITORAR
                monitorar_e_renomear(PASTA_DOWNLOAD, arquivos_antes, nome_final)

            except Exception as e:
                print(f"     > ERRO NA ABA: {e}")

            # Fecha e volta
            driver.close()
            driver.switch_to.window(janela_principal)

        except Exception as e:
            print(f"Erro genérico item {i}: {e}")
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(janela_principal)

    print("\nFim!")

except Exception as e:
    print(f"Erro Fatal: {e}")