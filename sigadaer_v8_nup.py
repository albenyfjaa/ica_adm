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
PASTA_DOWNLOAD = os.path.join(os.getcwd(), "downloads_sigadaer_2024_v3_NUP") # Nova pasta para evitar misturar

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
    # O NUP tem barras (/) e pontos (.), o regex abaixo troca barra por underline
    nome = re.sub(r'[<>:"/\\|?*]', '_', nome)
    return " ".join(nome.split())[:180]

# =============================================================================
# 2. FUNÇÃO: VERIFICAÇÃO (AGORA USANDO NUP)
# =============================================================================
def arquivo_ja_baixado_nup(nup_texto, pasta):
    """
    Verifica se existe algum arquivo na pasta que contenha o NUP informado.
    """
    if not nup_texto: return False
    
    # Sanitiza o NUP (Ex: '67609.001951/2024-11' vira '67609.001951_2024-11')
    nup_limpo = sanitizar_nome(nup_texto).lower()
    
    if len(nup_limpo) < 5: return False # Segurança contra NUP vazio
    
    try:
        arquivos = os.listdir(pasta)
    except FileNotFoundError:
        return False

    for f in arquivos:
        if nup_limpo in f.lower():
            return True
            
    return False

# =============================================================================
# 3. FUNÇÃO: MONITORAR E RENOMEAR
# =============================================================================
def monitorar_e_renomear(pasta, arquivos_antes, nome_desejado):
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
                    if "download" in arquivo.lower() or "documento" in arquivo.lower():
                        ext = os.path.splitext(arquivo)[1]
                        if not ext: ext = ".pdf"
                        
                        novo_nome = f"{nome_desejado}{ext}"
                        dest = os.path.join(pasta, novo_nome)
                        
                        c = 1
                        while os.path.exists(dest):
                            dest = os.path.join(pasta, f"{nome_desejado} ({c}){ext}")
                            c += 1
                            
                        os.rename(caminho_completo, dest)
                        return True
                    else:
                        # Se já baixou com nome certo (raro no SIGAD)
                        return True
            except:
                pass
        time.sleep(1)
    return False

# =============================================================================
# PREPARAÇÃO
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
        pass 

    print("\n--- ACESSANDO MENU ---")
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Documentos do Setor"))).click()
    preparar_lista_1000(driver, wait)

    print("\n--- INICIANDO PROCESSAMENTO (VIA NUP) ---")
    xpath_linhas = "//tbody[contains(@id, 'data')]//tr[@data-ri]"
    wait.until(EC.presence_of_all_elements_located((By.XPATH, xpath_linhas)))
    linhas = driver.find_elements(By.XPATH, xpath_linhas)
    total_docs = len(linhas)
    print(f" > Total encontrado: {total_docs}")

    janela_principal = driver.current_window_handle
    relatorio_erros = []

    for i in range(total_docs):
        identificador = "DESCONHECIDO"
        try:
            # Recaptura
            linhas = driver.find_elements(By.XPATH, xpath_linhas)
            if i >= len(linhas): break
            linha = linhas[i]
            
            # --- NOVA LÓGICA: PEGAR NUP ---
            # O NUP está na 7ª coluna (td[7]), mas o XPATH conta a partir de 1
            # Estrutura visual: [Chk][Abacaxi][Bola][PDF][Num(Link)][Sigad][NUP]
            # Indices Python: 0, 1, 2, 3, 4, 5, 6
            colunas = linha.find_elements(By.TAG_NAME, "td")
            
            # O link para clicar continua sendo no Número (coluna 4 ou 5)
            link_doc = linha.find_element(By.XPATH, ".//a")
            
            # Pega o texto do NUP (Coluna índice 6 = 7ª coluna)
            if len(colunas) > 6:
                nup_texto = colunas[6].text.strip()
            else:
                nup_texto = "SEM_NUP"

            identificador = nup_texto # Usaremos o NUP como ID principal
            
            # --- VERIFICAÇÃO ---
            if arquivo_ja_baixado_nup(identificador, PASTA_DOWNLOAD):
                print(f" [{i+1}/{total_docs}] Já existe (NUP): {identificador}")
                continue
            # -------------------

            print(f" [{i+1}/{total_docs}] Processando NUP: {identificador}...")
            arquivos_antes = set(os.listdir(PASTA_DOWNLOAD))

            # Abre nova aba
            ActionChains(driver).key_down(Keys.CONTROL).click(link_doc).key_up(Keys.CONTROL).perform()
            wait.until(EC.number_of_windows_to_be(2))
            
            # Muda para nova aba
            janelas = driver.window_handles
            driver.switch_to.window([w for w in janelas if w != janela_principal][0])

            # --- DENTRO DA NOVA ABA ---
            try:
                # 1. VERIFICAÇÃO DE ERRO DE SERVIDOR (Igual imagem enviada)
                src = driver.page_source
                if "ERR_RESPONSE_HEADERS_MULTIPLE_CONTENT_DISPOSITION" in src or "Esta página não está funcionando" in src:
                    msg = f"{identificador} - Erro Servidor (Header Duplo)"
                    print(f"     > [FALHA] {msg}")
                    relatorio_erros.append(msg)
                    driver.close()
                    driver.switch_to.window(janela_principal)
                    continue

                # 2. ESPERA CRÍTICA
                try:
                    wait_curto = WebDriverWait(driver, 5)
                    wait_curto.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Assunto:')]")))
                except TimeoutException:
                    msg = f"{identificador} - Timeout (Página em branco/não carregou)"
                    print(f"     > [FALHA] {msg}")
                    relatorio_erros.append(msg)
                    driver.close()
                    driver.switch_to.window(janela_principal)
                    continue
                
                # 3. CAPTURA ASSUNTO
                assunto_texto = ""
                try:
                    elem_assunto = driver.find_element(By.XPATH, "//*[contains(text(), 'Assunto:')]/..")
                    assunto_texto = elem_assunto.text.replace("Assunto:", "").strip()
                except:
                    pass

                # Monta nome usando NUP
                # Ex: 67609.001951_2024-11 - Assunto do Documento.pdf
                nome_base = f"{identificador} - {assunto_texto}"
                nome_final = sanitizar_nome(nome_base)

                # 4. DOWNLOAD
                try:
                    xpath_btn = "//button[.//span[text()='Download']]"
                    if len(driver.find_elements(By.XPATH, xpath_btn)) > 0:
                        btn = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_btn)))
                        driver.execute_script("arguments[0].click();", btn)
                        
                        sucesso = monitorar_e_renomear(PASTA_DOWNLOAD, arquivos_antes, nome_final)
                        if not sucesso:
                            relatorio_erros.append(f"{identificador} - Timeout Download")
                    else:
                        msg = f"{identificador} - Botão Download NÃO encontrado (Doc Invalidado?)"
                        print(f"     > [FALHA] {msg}")
                        relatorio_erros.append(msg)

                except Exception as e:
                    relatorio_erros.append(f"{identificador} - Erro no clique: {str(e)[:50]}")

            except Exception as e:
                print(f"     > ERRO NA ABA: {e}")
                relatorio_erros.append(f"{identificador} - Erro Genérico Aba")

            # Fecha e volta
            if len(driver.window_handles) > 1:
                driver.close()
            driver.switch_to.window(janela_principal)

        except Exception as e:
            print(f"Erro item {i}: {e}")
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(janela_principal)

    print("\n" + "="*50)
    print(f"FIM! Falhas: {len(relatorio_erros)}")
    if relatorio_erros:
        print("\n".join(relatorio_erros))

except Exception as e:
    print(f"Erro Fatal: {e}")