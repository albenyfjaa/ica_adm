import time
import os
import urllib3
import glob
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
# FUNÇÃO: LIMPEZA DE NOME DE ARQUIVO
# =============================================================================
def sanitizar_nome(nome):
    """Remove caracteres inválidos para nomes de arquivos no Windows."""
    # Remove caracteres proibidos: < > : " / \ | ? *
    nome = re.sub(r'[<>:"/\\|?*]', '_', nome)
    # Remove espaços duplos e quebras de linha
    nome = " ".join(nome.split())
    # Limita o tamanho para evitar erro de caminho muito longo
    return nome[:150]

# =============================================================================
# FUNÇÃO: MONITORAR E RENOMEAR DOWNLOAD
# =============================================================================
def monitorar_e_renomear(pasta, arquivos_antes, nome_desejado):
    """
    Aguarda um novo arquivo aparecer na pasta.
    Se o nome for genérico ('download...'), renomeia para 'nome_desejado'.
    """
    print("     > Aguardando conclusão do download...")
    fim_tempo = time.time() + 60  # Timeout de 60 segundos
    
    while time.time() < fim_tempo:
        # Lista arquivos atuais
        arquivos_agora = set(os.listdir(pasta))
        novos_arquivos = arquivos_agora - arquivos_antes
        
        if novos_arquivos:
            # Pega o novo arquivo (ignora .crdownload ou .tmp enquanto baixa)
            for arquivo in novos_arquivos:
                if arquivo.endswith('.crdownload') or arquivo.endswith('.tmp'):
                    continue
                
                caminho_completo = os.path.join(pasta, arquivo)
                
                # Verifica se o download terminou (tamanho > 0 e acessível)
                try:
                    if os.path.getsize(caminho_completo) > 0:
                        print(f"     > Arquivo detectado: {arquivo}")
                        
                        # VERIFICAÇÃO CRÍTICA: O nome é genérico?
                        # O Chrome salva como 'download', 'download (1)', etc.
                        if "download" in arquivo.lower():
                            novo_caminho = os.path.join(pasta, f"{nome_desejado}.pdf")
                            
                            # Se já existir um arquivo com o nome ideal, adiciona contador
                            contador = 1
                            while os.path.exists(novo_caminho):
                                novo_caminho = os.path.join(pasta, f"{nome_desejado}_{contador}.pdf")
                                contador += 1
                                
                            os.rename(caminho_completo, novo_caminho)
                            print(f"     > RENOMEADO PARA: {os.path.basename(novo_caminho)}")
                            return True
                        else:
                            print("     > O arquivo já veio com nome correto. Mantendo.")
                            return True
                except OSError:
                    pass # Arquivo ainda sendo escrito, tenta no próximo loop
                    
        time.sleep(1)
        
    print("     > Tempo esgotado ou download falhou.")
    return False

# =============================================================================
# FUNÇÃO: PREPARAR A LISTA
# =============================================================================
def preparar_lista_1000(driver, wait):
    # (Mesma lógica robusta do script anterior)
    xpath_select_ano = "//select[.//option[@value='2025']]"
    try:
        try:
            select_element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath_select_ano)))
        except:
            driver.find_element(By.LINK_TEXT, "Documentos do Setor").click()
            select_element = wait.until(EC.presence_of_element_located((By.XPATH, xpath_select_ano)))

        select_ano = Select(select_element)
        if select_ano.first_selected_option.get_attribute("value") != "2025":
            select_ano.select_by_value("2025")
            btn = driver.find_element(By.XPATH, "//button[contains(@id, 'aplicarFiltrosButton')]")
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(3)
            
        # Selecionar 1000
        id_select_paginacao = "caixaForm:tableDocumentos_rppDD"
        sel_pag_element = wait.until(EC.presence_of_element_located((By.ID, id_select_paginacao)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", sel_pag_element)
        select_pag = Select(sel_pag_element)
        
        if select_pag.first_selected_option.text != "1000":
            select_pag.select_by_value("1000")
            print(" > Modo 1000 itens ativado. Aguardando tabela...")
            time.sleep(5)
        return True
    except Exception as e:
        print(f" > Erro na preparação: {e}")
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
            # Recaptura rápida
            linhas = driver.find_elements(By.XPATH, xpath_linhas)
            if i >= len(linhas): break
            linha = linhas[i]
            link_doc = linha.find_element(By.XPATH, ".//a")
            
            # Dados Básicos
            numero_doc = link_doc.text.strip() # Ex: 149/DAIN/22989
            
            # Verifica se já existe (verificação simples pelo número)
            termo_busca = numero_doc.replace('/', '_').replace('\\', '_')
            ja_existe = False
            for f in os.listdir(PASTA_DOWNLOAD):
                if termo_busca in f:
                    ja_existe = True
                    break
            
            if ja_existe:
                print(f" [{i+1}/{total_docs}] Já existe: {numero_doc}")
                continue

            print(f" [{i+1}/{total_docs}] Processando: {numero_doc}...")

            # SNAPSHOT DA PASTA (Para monitorar o novo arquivo)
            arquivos_antes = set(os.listdir(PASTA_DOWNLOAD))

            # Abre nova aba
            ActionChains(driver).key_down(Keys.CONTROL).click(link_doc).key_up(Keys.CONTROL).perform()
            wait.until(EC.number_of_windows_to_be(2))
            janelas = driver.window_handles
            nova_janela = [w for w in janelas if w != janela_principal][0]
            driver.switch_to.window(nova_janela)

            # --- DENTRO DA NOVA ABA ---
            try:
                # 1. TENTA CAPTURAR O ASSUNTO (Para compor o nome)
                # Procura por um elemento que contenha "Assunto:" e pega o texto
                assunto_texto = ""
                try:
                    # XPath genérico para pegar o texto após o label 'Assunto:'
                    # Baseado em layout padrão JSF/Primefaces
                    elem_assunto = driver.find_element(By.XPATH, "//*[contains(text(), 'Assunto:')]/..")
                    assunto_texto = elem_assunto.text.replace("Assunto:", "").strip()
                except:
                    # Se falhar, tenta pegar apenas o texto da página e procurar
                    pass

                # Monta o nome desejado
                if assunto_texto:
                    nome_base = f"{numero_doc} - {assunto_texto}"
                else:
                    nome_base = numero_doc # Fallback só com número se não achar assunto
                
                nome_final = sanitizar_nome(nome_base)

                # 2. CLICA NO DOWNLOAD
                xpath_btn_download = "//button[.//span[text()='Download']]"
                btn_down = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_btn_download)))
                driver.execute_script("arguments[0].click();", btn_down)
                
                # 3. CHAMA A FUNÇÃO DE RENOMEAR
                monitorar_e_renomear(PASTA_DOWNLOAD, arquivos_antes, nome_final)

            except Exception as e:
                print(f"     > Erro na aba de detalhe: {e}")

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