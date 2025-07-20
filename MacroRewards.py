import nltk
from nltk.corpus import floresta
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
import time
import random
import os
import subprocess
from datetime import datetime, date

def check_if_ran_today():
    log_file = "process_log.txt"
    hoje = date.today().strftime("%Y-%m-%d")
    if os.path.exists(log_file):
        with open(log_file, 'r') as f:
            data_salva = f.read().strip()
            if data_salva == hoje:
                print(f"Script já foi executado hoje ({hoje}). Encerrando.")
                return True
    return False

def log_execution():
    log_file = "process_log.txt"
    hoje = date.today().strftime("%Y-%m-%d")
    with open(log_file, 'w') as f:
        f.write(f"{hoje}\n")
    print(f"Execução de hoje ({hoje}) registrada.")

def verificar_conexao():
    url = "http://www.google.com"
    timeout = 5
    while True:
        try:
            _ = requests.get(url, timeout=timeout)
            print("Conexão com a internet verificada!")
            return True
        except requests.ConnectionError:
            print("Sem conexão com a internet. Tentando novamente em 5 segundos...")
            time.sleep(5)

def garantir_fora_da_tela(navegador):
    navegador.set_window_position(-2000, -2000)
    time.sleep(0.1)

def limpar_processos():
    try:
        subprocess.run(
            ['taskkill', '/im', 'chromedriver.exe', '/f', '/t'],
            creationflags=subprocess.CREATE_NO_WINDOW,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=False
        )
        subprocess.run(
            ['taskkill', '/im', 'chrome.exe', '/f', '/t'],
            creationflags=subprocess.CREATE_NO_WINDOW,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=False
        )
        print("Processos do Chrome e WebDriver encerrados com sucesso.")
    except Exception as e:
        print(f"Erro ao encerrar processos: {e}")

def executar_script():
    try:
        chrome_profile_path = os.path.expanduser("~\\AppData\\Local\\Google\\Chrome\\User Data\\SeleniumProfile")

        chrome_options = ChromeOptions()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--remote-debugging-port=9222')
        chrome_options.add_argument('--disable-software-rasterizer')
        chrome_options.add_argument('--enable-logging')
        chrome_options.add_argument('--log-level=1')
        chrome_options.add_argument('--force-device-scale-factor=1')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument(f"--user-data-dir={chrome_profile_path}")

        navegador = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
        navegador.set_window_position(-2000, -2000)
        navegador.set_window_size(1280, 720)
        print("Navegador inicializado com o perfil do Chrome padrão.")

        wait = WebDriverWait(navegador, 5)
        navegador.get("https://www.microsoft.com/pt-br/rewards/about?rtc=1")
        garantir_fora_da_tela(navegador)

        abrir = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="card-body-highlight-615a0958-3258-424f-b04a-0eb65b1b0fa0"]/div[3]/a')))
        navegador.refresh()
        garantir_fora_da_tela(navegador)
        abrir = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="card-body-highlight-615a0958-3258-424f-b04a-0eb65b1b0fa0"]/div[3]/a')))
        time.sleep(random.uniform(2, 4))
        navegador.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", abrir)
        time.sleep(random.uniform(2, 5))
        abrir.click()
        garantir_fora_da_tela(navegador)

        navegador.switch_to.window(navegador.window_handles[1])
        garantir_fora_da_tela(navegador)
        tentativas = 0

        while tentativas < 30:
            try:
                elementos = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="daily-sets"]')))
                if elementos:
                    print("Página de missões carregada!")
                    break
            except Exception:
                print(f"Tentativa {tentativas + 1} falhou. Recarregando a página...")
                navegador.refresh()
                garantir_fora_da_tela(navegador)
                tentativas += 1

        else:
            print("Não foi possível carregar a página de missões após 30 tentativas.")
            navegador.quit()
            limpar_processos()
            return False

        indice_aba = 2

        while True:
            navegador.switch_to.window(navegador.window_handles[1])
            garantir_fora_da_tela(navegador)
            print("Mudou para a aba 2 para buscar os elementos.")

            navegador.refresh()
            garantir_fora_da_tela(navegador)
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="daily-sets"]')))

            try:
                elementos = navegador.find_elements(By.CLASS_NAME, 'mee-icon-AddMedium')
                if not elementos:
                    check_elementos = navegador.find_elements(By.CLASS_NAME, 'mee-icon-SkypeCircleCheck')
                    if check_elementos:
                        print("Todas as missões foram completadas!")
                        break
                    else:
                        print("Nenhum elemento encontrado.")
                        navegador.quit()
                        limpar_processos()
                        return False

                for i, elemento in enumerate(elementos):
                    try:
                        elemento.click()
                        print(f"Elemento {i + 1} clicado.")
                        time.sleep(2)
                        if len(navegador.window_handles) > indice_aba:
                            navegador.switch_to.window(navegador.window_handles[indice_aba])
                            garantir_fora_da_tela(navegador)
                        else:
                            continue

                        pagina = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="b_content"]')))
                        if pagina:
                            time.sleep(random.uniform(3, 5))
                        else:
                            navegador.refresh()
                            garantir_fora_da_tela(navegador)
                            time.sleep(random.uniform(2, 5))

                        try:
                            aceitar = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="bnp_btn_accept"]')))
                            if aceitar:
                                aceitar.click()
                        except:
                            pass

                    except Exception as e:
                        print(f"Erro ao clicar no elemento {i + 1}: {e}")

                    navegador.switch_to.window(navegador.window_handles[1])
                    garantir_fora_da_tela(navegador)
                    indice_aba += 1

            except Exception as e:
                print(f"Erro ao buscar elementos: {e}")
                navegador.quit()
                limpar_processos()
                return False

        executar_buscas_floresta(navegador, wait)
        navegador.quit()
        limpar_processos()
        return True

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        navegador.quit()
        limpar_processos()
        return False

def executar_buscas_floresta(navegador, wait):
    nltk.download('floresta')
    palavras = [palavra.lower() for (palavra, tipo) in floresta.tagged_words() if len(palavra) > 6]

    def palavra_aleatoria():
        return random.choice(palavras)

    try:
        for i in range(random.randint(5, 11)):
            termo_sorteado = palavra_aleatoria()
            navegador.execute_script("window.open('');")
            navegador.switch_to.window(navegador.window_handles[-1])
            garantir_fora_da_tela(navegador)

            navegador.get("https://www.bing.com/")
            garantir_fora_da_tela(navegador)

            psq = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="sb_form_c"]')))
            psq.click()
            text = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="sb_form_q"]')))
            text.send_keys(termo_sorteado)
            text.send_keys(Keys.ENTER)

            WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, '//ol[@id="b_results"]')))
            garantir_fora_da_tela(navegador)

            time.sleep(random.uniform(2, 5))

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    time.sleep(random.uniform(1.5, 3))
    navegador.quit()
    limpar_processos()

def main():
    if check_if_ran_today():
        return

    tentativas = 0
    max_tentativas = 10
    sucesso = False

    while not sucesso and tentativas < max_tentativas:
        tentativas += 1
        print(f"Tentativa {tentativas} de {max_tentativas}")
        sucesso = executar_script()

        if not sucesso:
            print(f"Tentativa {tentativas} falhou. Tentando novamente em 5 segundos...")
            time.sleep(5)

    if sucesso:
        print("Script executado com sucesso!")
        log_execution()
    else:
        print("O script falhou após várias tentativas.")

    limpar_processos()

if __name__ == "__main__":
    if verificar_conexao():
        main()