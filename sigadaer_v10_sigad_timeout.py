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

# Pasta específica para esta versão organizada por SIGAD ID
PASTA_DOWNLOAD = r"D:\Nuvem\OneDrive - DECEA\01-ICA\SIGADAER\CAR2\downloads_sigadaer_2025"

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
# 2. FUNÇÃO: VERIFICAÇÃO (POR NÚMERO SIGAD)
# =============================================================================
def arquivo_ja_baixado_sigad(sigad_id, pasta):
    """
    Verifica se existe algum arquivo na pasta que comece com o ID Sigad.
    """
    if not sigad_id or sigad_id == "SEM_SIGAD": return False
    
    id_limpo = sigad_id.strip()
    
    try:
        arquivos = os.listdir(pasta)
    except FileNotFoundError:
        return False

    for f in arquivos:
        # Verifica se o arquivo COMEÇA com o ID
        if f.startswith(id_limpo):
            resto = f[len(id_limpo):]
            if not resto: return True
            
            primeiro_char_resto = resto[0]
            if primeiro_char_resto in [' ', '-', '_', '.']:
                return True
            
    return False

# =============================================================================
# 3. FUNÇÃO: MONITORAR E RENOMEAR
# =============================================================================
def monitorar_e_renomear(pasta, arquivos_antes, nome_desejado):
    fim_tempo = time.time() + 60 # Timeout de 60 segundos
    while time.time() < fim_tempo:
        arquivos_agora = set(os.listdir(pasta))
        novos = arquivos_agora - arquivos_antes
        
        for arquivo in novos:
            if arquivo.endswith('.crdownload') or arquivo.endswith('.tmp'):
                continue
                
            caminho_completo = os.path.join(pasta, arquivo)
            try:
                if os.path.getsize(caminho_completo) > 0:
                    if "download" in arquivo.lower() or "documento" in arquivo.lower() or ".pdf" in arquivo.lower():
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
                        return True
            except:
                pass
        time.sleep(1)
    return False

# =============================================================================
# PREPARAÇÃO (FILTRO 2024 + 1000 ITENS)
# =============================================================================
def preparar_lista_1000(driver, wait):
    xpath_select_ano = "//select[.//option[@value='2025']]"
    try:
        try:
            sel = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath_select_ano)))
        except:
            driver.find_element(By.LINK_TEXT, "Documentos do Setor").click()
            sel = wait.until(EC.presence_of_element_located((By.XPATH, xpath_select_ano)))

        select_ano = Select(sel)
        if select_ano.first_selected_option.get_attribute("value") != "2025":
            select_ano.select_by_value("2025")
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

    print("\n--- INICIANDO PROCESSAMENTO (OTIMIZADO) ---")
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
            linhas = driver.find_elements(By.XPATH, xpath_linhas)
            if i >= len(linhas): break
            linha = linhas[i]
            
            colunas = linha.find_elements(By.TAG_NAME, "td")
            link_doc = linha.find_element(By.XPATH, ".//a")
            
            # 1. PEGAR ID SIGAD (Coluna 5)
            if len(colunas) > 5:
                sigad_texto = colunas[5].text.strip()
            else:
                sigad_texto = "SEM_SIGAD"

            identificador = sigad_texto 
            
            # 2. VERIFICAÇÃO ANTES DE ABRIR
            if arquivo_ja_baixado_sigad(identificador, PASTA_DOWNLOAD):
                print(f" [{i+1}/{total_docs}] Já existe (SIGAD): {identificador}")
                continue
            
            print(f" [{i+1}/{total_docs}] Baixando SIGAD: {identificador}...")
            arquivos_antes = set(os.listdir(PASTA_DOWNLOAD))

            # Abre nova aba
            ActionChains(driver).key_down(Keys.CONTROL).click(link_doc).key_up(Keys.CONTROL).perform()
            wait.until(EC.number_of_windows_to_be(2))
            
            janelas = driver.window_handles
            driver.switch_to.window([w for w in janelas if w != janela_principal][0])

            # --- DENTRO DA NOVA ABA (AJUSTE CRÍTICO DE TIMEOUT) ---
            try:
                # DÁ UM RESPIRO PARA O CHROME RENDERIZAR O ERRO
                time.sleep(1.5) 

                # A. VERIFICAÇÃO RÁPIDA DE ERRO
                src = driver.page_source
                
                # Se a página for muito pequena (< 200 chars), provavelmente é tela branca de erro
                if len(src) < 200:
                    msg = f"{identificador} - Erro Tela Branca (Conteúdo Vazio)"
                    print(f"     > [FALHA RÁPIDA] {msg}")
                    relatorio_erros.append(msg)
                    driver.close()
                    driver.switch_to.window(janela_principal)
                    continue

                if "ERR_RESPONSE_HEADERS_MULTIPLE_CONTENT_DISPOSITION" in src or "Esta página não está funcionando" in src:
                    msg = f"{identificador} - Erro Servidor (Header Duplo)"
                    print(f"     > [FALHA RÁPIDA] {msg}")
                    relatorio_erros.append(msg)
                    driver.close()
                    driver.switch_to.window(janela_principal)
                    continue

                # B. ESPERA CRÍTICA REDUZIDA (3 SEGUNDOS)
                try:
                    # Se não aparecer em 3s, assume que travou
                    wait_curto = WebDriverWait(driver, 3) 
                    wait_curto.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Assunto:')]")))
                except TimeoutException:
                    msg = f"{identificador} - Timeout de 3s (Página travada)"
                    print(f"     > [FALHA] {msg}")
                    relatorio_erros.append(msg)
                    driver.close()
                    driver.switch_to.window(janela_principal)
                    continue
                
                # C. CAPTURA ASSUNTO
                assunto_texto = ""
                try:
                    elem_assunto = driver.find_element(By.XPATH, "//*[contains(text(), 'Assunto:')]/..")
                    assunto_texto = elem_assunto.text.replace("Assunto:", "").strip()
                except:
                    pass

                # D. MONTAGEM DO NOME: [SIGAD] - [ASSUNTO]
                nome_base = f"{identificador} - {assunto_texto}"
                nome_final = sanitizar_nome(nome_base)

                # E. CLIQUE NO DOWNLOAD
                try:
                    xpath_btn = "//button[.//span[text()='Download']]"
                    if len(driver.find_elements(By.XPATH, xpath_btn)) > 0:
                        btn = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_btn)))
                        driver.execute_script("arguments[0].click();", btn)
                        
                        sucesso = monitorar_e_renomear(PASTA_DOWNLOAD, arquivos_antes, nome_final)
                        if not sucesso:
                            relatorio_erros.append(f"{identificador} - Timeout no Download")
                    else:
                        msg = f"{identificador} - Botão Download NÃO encontrado (Doc Invalidado?)"
                        print(f"     > [FALHA] {msg}")
                        relatorio_erros.append(msg)

                except Exception as e:
                    relatorio_erros.append(f"{identificador} - Erro no clique: {str(e)[:50]}")

            except Exception as e:
                print(f"     > ERRO NA ABA: {e}")
                relatorio_erros.append(f"{identificador} - Erro Genérico Aba")

            # Fecha aba e volta
            if len(driver.window_handles) > 1:
                driver.close()
            driver.switch_to.window(janela_principal)

        except Exception as e:
            print(f"Erro item {i}: {e}")
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(janela_principal)

    # RELATÓRIO FINAL
    print("\n" + "="*50)
    print(f"FIM DO PROCESSAMENTO! Falhas: {len(relatorio_erros)}")
    if relatorio_erros:
        print("Lista de Erros:")
        print("\n".join(relatorio_erros))
        
        # --- ALTERAÇÃO: Salva log NA MESMA PASTA DOS DOCUMENTOS ---
        caminho_log = os.path.join(PASTA_DOWNLOAD, "erros_sigad_final.txt")
        with open(caminho_log, "w", encoding="utf-8") as f:
            f.write("\n".join(relatorio_erros))
        print(f"\nLog de erros salvo em: {caminho_log}")

except Exception as e:
    print(f"Erro Fatal: {e}")