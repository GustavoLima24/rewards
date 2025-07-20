import nltk
from nltk.corpus import floresta
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.options import Options
import time
import random
import os
from datetime import datetime, date

def check_if_ran_today():
    """Verifica se o script já foi executado hoje."""
    log_file = "process_log.txt"
    hoje = date.today().strftime("%Y-%m-%d")
    
    try:
        if os.path.exists(log_file):
            with open(log_file, 'r') as f:
                data_salva = f.read().strip()
                if data_salva == hoje:
                    print(f"Script já foi executado hoje ({hoje}). Encerrando.")
                    return True
        return False
    except Exception as e:
        print(f"Erro ao ler o arquivo de log: {e}")
        return False

def log_execution():
    """Grava a data de hoje no log, sobrescrevendo o conteúdo anterior."""
    log_file = "process_log.txt"
    hoje = date.today().strftime("%Y-%m-%d")
    
    try:
        with open(log_file, 'w') as f:  # modo 'w' sobrescreve tudo
            f.write(f"{hoje}\n")
        print(f"Execução de hoje ({hoje}) registrada.")
    except Exception as e:
        print(f"Erro ao escrever no log: {e}")

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
    """Função para garantir que a janela do navegador fique fora da tela."""
    navegador.set_window_position(-2000, -2000)
    time.sleep(0.1)

def limpar_processos():
    """Encerra todos os processos do Edge e WebDriver em segundo plano."""
    try:
        os.system('taskkill /im msedgedriver.exe /f /t')
        os.system('taskkill /im msedge.exe /f /t')
        print("Processos do Edge e WebDriver encerrados com sucesso.")
    except Exception as e:
        print(f"Erro ao encerrar processos: {e}")

def executar_script():
    try:
        edge_options = webdriver.EdgeOptions()
        edge_options.use_chromium = True
        edge_options.add_argument('--no-sandbox')
        edge_options.add_argument('--disable-dev-shm-usage')
        edge_options.add_argument('--remote-debugging-port=9222')
        edge_options.add_argument('--disable-gpu')
        edge_options.add_argument('--disable-software-rasterizer')
        edge_options.add_argument('--enable-logging')
        edge_options.add_argument(r"--user-data-dir=C:\Users\gusou\AppData\Local\Microsoft\Edge\User Data")
        edge_options.add_argument("--profile-directory=Default")
        edge_options.add_argument('--log-level=1')
        edge_options.add_argument('--force-device-scale-factor=1')
        edge_options.add_argument('--disable-blink-features=AutomationControlled')

        try:
            navegador = webdriver.Edge(
                service=EdgeService(
                    EdgeChromiumDriverManager().install(),
                    log_path='msedgedriver.log'
                ),
                options=edge_options
            )
            navegador.set_window_position(-2000, -2000)
            navegador.set_window_size(1280, 720)
            print("Navegador inicializado fora da tela com tamanho 1280x720.")
        except Exception as e:
            print(f"Falha crítica: {e}")
            limpar_processos()
            return False

        navegador.get("edge://settings/")
        garantir_fora_da_tela(navegador)
        print("Perfil do Edge carregado. Verifique se as configurações estão visíveis.")

        tempo_maximo = 5
        wait = WebDriverWait(navegador, tempo_maximo)

        navegador.get("https://www.microsoft.com/pt-br/rewards/about?rtc=1")
        garantir_fora_da_tela(navegador)
        print("Página carregada com sucesso.")

        abrir = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="card-body-highlight-615a0958-3258-424f-b04a-0eb65b1b0fa0"]/div[3]/a')))
        navegador.refresh()
        garantir_fora_da_tela(navegador)
        abrir = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="card-body-highlight-615a0958-3258-424f-b04a-0eb65b1b0fa0"]/div[3]/a')))
        time.sleep(random.uniform(2, 4))
        navegador.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", abrir)
        time.sleep(random.uniform(2, 5))
        if abrir:
            time.sleep(2)
            abrir.click()
            garantir_fora_da_tela(navegador)

        try:
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

        except Exception as e:
            print(f"Erro ao carregar a página de missões: {e}")
            navegador.quit()
            limpar_processos()
            return False

        time.sleep(2)
        indice_aba = 2

        while True:
            navegador.switch_to.window(navegador.window_handles[1])
            garantir_fora_da_tela(navegador)
            print("Mudou para a aba 2 para buscar os elementos.")

            # Refresh page and wait for it to load
            navegador.refresh()
            garantir_fora_da_tela(navegador)
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="daily-sets"]')))
            print("Página de missões recarregada e carregada completamente.")

            try:
                elementos = navegador.find_elements(By.CLASS_NAME, 'mee-icon-AddMedium')
                
                if not elementos:
                    check_elementos = navegador.find_elements(By.CLASS_NAME, 'mee-icon-SkypeCircleCheck')
                    if check_elementos:
                        print("Encontrados apenas elementos 'mee-icon-SkypeCircleCheck'. Todas as missões foram completadas!")
                        break  # Exit the loop when no more 'mee-icon-AddMedium' elements are found
                    else:
                        print("Nenhum elemento 'mee-icon-AddMedium' ou 'mee-icon-SkypeCircleCheck' encontrado.")
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
                            print(f"Alterou para a aba {indice_aba + 1}")
                        else:
                            print("Não há mais abas disponíveis para alterar.")
                            continue

                        pagina = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="b_content"]')))
                        if pagina:
                            print("A página foi carregada!")
                            time.sleep(random.uniform(3, 5))
                        else:
                            navegador.refresh()
                            garantir_fora_da_tela(navegador)
                            print("A página não foi carregada corretamente!")
                            time.sleep(random.uniform(2, 5))

                        try:
                            aceitar = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="bnp_btn_accept"]')))
                            if aceitar:
                                print("Aceitou os termos!")
                                aceitar.click()
                        except:
                            print("Não perguntou sobre aceitar")

                        print(f"Cliquei em {elemento} missão realizada!")
                    except Exception as e:
                        print(f"Erro ao clicar no elemento {i + 1}: {e}")

                    navegador.switch_to.window(navegador.window_handles[1])
                    garantir_fora_da_tela(navegador)
                    print("Retornou para a aba 2 para buscar novamente.")
                    indice_aba += 1

            except Exception as e:
                print(f"Erro ao buscar elementos: {e}")
                navegador.quit()
                limpar_processos()
                return False

        # Proceed to search after all missions are completed
        executar_buscas_floresta(navegador, wait)
        print("Navegador fechado.")
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
            print(f"Aba {i + 1} aberta.")

            navegador.get(f"https://www.bing.com/")
            garantir_fora_da_tela(navegador)
            print(f"Busca {i + 1} pelo termo '{termo_sorteado}' realizada.")

            psq = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="sb_form_c"]')))
            psq.click()
            text = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="sb_form_q"]')))
            text.send_keys(termo_sorteado)
            text.send_keys(Keys.ENTER)
            print(f"Pesquisa por '{termo_sorteado}' enviada.")

            WebDriverWait(navegador, 10).until(EC.visibility_of_element_located((By.XPATH, '//ol[@id="b_results"]')))
            garantir_fora_da_tela(navegador)
            print(f"Resultados de busca {i + 1} carregados.")

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