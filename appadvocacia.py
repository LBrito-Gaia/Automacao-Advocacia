import os
import time
import json
import re
import shutil
from docx import Document
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def configurar_driver(pasta_temp):
    options = webdriver.ChromeOptions()
    caminho_abs = os.path.abspath(pasta_temp)
    settings = {
        "recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
    prefs = {
        "printing.print_preview_sticky_settings.appState": json.dumps(settings),
        "savefile.default_directory": caminho_abs,
        "download.default_directory": caminho_abs,
        "download.prompt_for_download": False,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument('--kiosk-printing')
    return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

def processar_comprovantes_advocacia(arquivo_word, pasta_cliente):
    pasta_temp = 'Temp_Downloads'
    pasta_pagos = os.path.join(pasta_cliente, 'Comprovantes_Pagos')
    pasta_pendentes = os.path.join(pasta_cliente, 'Comprovantes_Pendentes')
    
    for p in [pasta_temp, pasta_pagos, pasta_pendentes]:
        if not os.path.exists(p): os.makedirs(p)

    doc = Document(arquivo_word)
    links = re.findall(r'(https?://[^\s]+)', "\n".join([p.text for p in doc.paragraphs]))
    
    if not links:
        print(f"   - Nenhum link encontrado no arquivo!")
        return 0

    driver = configurar_driver(pasta_temp)
    comprovantes_baixados = 0

    try:
        for i, url in enumerate(links):
            print(f"   [{i+1}/{len(links)}] Acessando: {url}")
            driver.get(url)

            time.sleep(3)

            texto_pagina = driver.find_element(By.TAG_NAME, "body").text.upper()
            e_pago = any(termo in texto_pagina for termo in ["PAGO", "RECEBIDO", "CONFIRMADO"])
            status = "PAGO" if e_pago else "PENDENTE"
            pasta_destino = pasta_pagos if e_pago else pasta_pendentes

            try:
                script_get_href = """
                var elementos = document.querySelectorAll('atlas-link');
                for (var i = 0; i < elementos.length; i++) {
                    var texto = elementos[i].getAttribute('data-track-link-label');
                    if (texto && texto.toLowerCase().includes('visualizar comprovante')) {
                        return elementos[i].href || elementos[i].getAttribute('href');
                    }
                }
                var links = document.querySelectorAll('a');
                for (var i = 0; i < links.length; i++) {
                    if (links[i].href && links[i].href.includes('comprovantes')) {
                        return links[i].href;
                    }
                }
                return null;
                """
                href_comprovante = driver.execute_script(script_get_href)

                if href_comprovante:
                    driver.get(href_comprovante)
                    time.sleep(3)
                else:
                    corpo = driver.find_element(By.TAG_NAME, 'body')
                    corpo.send_keys(Keys.HOME)
                    time.sleep(1)
                    for _ in range(8):
                        corpo.send_keys(Keys.TAB)
                        time.sleep(0.3)
                    corpo.send_keys(Keys.ENTER)
                    time.sleep(3)

                time.sleep(10)
                driver.execute_script('window.print();')
                time.sleep(2)

                arquivos = [f for f in os.listdir(pasta_temp) if f.endswith('.pdf')]
                if arquivos:
                    arquivos.sort(key=lambda x: os.path.getmtime(os.path.join(pasta_temp, x)), reverse=True)
                    arquivo_origem = os.path.join(pasta_temp, arquivos[0])
                    nome_final = f"Comprovante_{i+1}_{status}.pdf"
                    shutil.move(arquivo_origem, os.path.join(pasta_destino, nome_final))
                    comprovantes_baixados += 1

                driver.get("about:blank")

            except Exception as e:
                print(f"      - Erro: {e}")

    finally:
        driver.quit()
    
    return comprovantes_baixados

def processar_todos_clientes(pasta_entrada):
    # Lista todos os arquivos .docx na pasta de entrada
    arquivos_docx = [f for f in os.listdir(pasta_entrada) if f.endswith('.docx')]
    
    if not arquivos_docx:
        print("Nenhum arquivo .docx encontrado na pasta!")
        return

    print(f"\n===Encontrados {len(arquivos_docx)} arquivos===\n")
    
    total_comprovantes = 0
    
    for arquivo in arquivos_docx:
        # Remove a extensão .docx para obter o nome do cliente
        nome_cliente = arquivo.replace('.docx', '')
        caminho_arquivo = os.path.join(pasta_entrada, arquivo)
        
        print(f"\n{'='*50}")
        print(f"Processando: {nome_cliente}")
        print(f"{'='*50}")
        
        # Cria pasta do cliente
        pasta_cliente = os.path.join('Comprovantes_Clientes', nome_cliente)
        
        # Processa os comprovantes
        qtd = processar_comprovantes_advocacia(caminho_arquivo, pasta_cliente)
        total_comprovantes += qtd
        
        print(f"\n   -> {qtd} comprovantes baixados para {nome_cliente}")

    print(f"\n{'='*50}")
    print(f"PROCESSO FINALIZADO!")
    print(f"Total de comprovantes baixados: {total_comprovantes}")
    print(f"Arquivos processados: {len(arquivos_docx)}")
    print(f"{'='*50}")

if __name__ == "__main__":
    # Define a pasta onde estão os arquivos .docx
    pasta_entrada = 'Clientes'
    
    # Cria a pasta de entrada se não existir
    if not os.path.exists(pasta_entrada):
        os.makedirs(pasta_entrada)
        print(f"Pasta '{pasta_entrada}' criada. Coloque seus arquivos .docx nela!")
    else:
        processar_todos_clientes(pasta_entrada)
