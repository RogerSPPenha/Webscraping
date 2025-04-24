from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import pandas as pd
import time
from selenium.common.exceptions import TimeoutException

chrome_driver_path = "C:\Program Files\chromedriver-win64\chromedriver-win64\chromedriver.exe"


service_chrome = Service(chrome_driver_path)
options_chrome = webdriver.ChromeOptions() 
options_chrome.add_argument("--disable-gpu") 
options_chrome.add_argument("--window-size=1920,1080")
options_chrome.add_argument("--headless")

driver = webdriver.Chrome(service=service_chrome, options=options_chrome)

url_base = "https://www.kabum.com.br/hardware/placa-de-video-vga"
driver.get(url_base)
time.sleep(5)

dict_produtos = {
    "marca": [],
    "preco": [],
    "parcela": [],
    "desconto": [],
    "frete": [],
    "avaliacoes": [],
    "estrelas": []
}

pagina = 1

while True:
    print(f"\n Coletando dados da página {pagina}...")

    try:
        WebDriverWait(driver, 10).until(
            ec.presence_of_all_elements_located((By.CLASS_NAME, "productCard"))
        )
        print("Elementos encontrados com sucesso!")
    except TimeoutException:
        print("Tempo de espera excedido!")

    produtos = driver.find_elements(By.CLASS_NAME, "productCard")

    for produto in produtos:
        try:
            nome = produto.find_element(By.CLASS_NAME, "nameCard").text.strip()
            preco = produto.find_element(By.CLASS_NAME, "priceCard").text.strip()
            parcela = produto.find_element(By.CLASS_NAME, "priceTextCard").text.strip()
            try:
                desconto = produto.find_element(By.CLASS_NAME, "rounded-8").text.strip()
            except:
                desconto = "não há desconto"
            try:
                produto.find_element(By.CLASS_NAME, "bg-success-500").text.strip()
                frete = "Grátis"
            except:
                frete = "Pago"
            try:
                avaliacoes = produto.find_element(By.CLASS_NAME, "pt-4").text.strip()
            except:
                avaliacoes = "não há avaliações"
            try:
                produto.find_element(By.CSS_SELECTOR, "div[aria-label='Classificação: 5 de 5 estrelas.']").text.strip()
                estrelas = 5
                dict_produtos["estrelas"].append(estrelas)
            except:
                try:
                    produto.find_element(By.CSS_SELECTOR, "div[aria-label='Classificação: 4 de 5 estrelas.']").text.strip()
                    estrelas = 4
                    dict_produtos["estrelas"].append(estrelas)
                except:
                    try:
                        produto.find_element(By.CSS_SELECTOR, "div[aria-label='Classificação: 3 de 5 estrelas.']").text.strip()
                        estrelas = 3
                        dict_produtos["estrelas"].append(estrelas)
                    except:
                        try:
                            produto.find_element(By.CSS_SELECTOR, "div[aria-label='Classificação: 2 de 5 estrelas.']").text.strip()
                            estrelas = 2
                            dict_produtos["estrelas"].append(estrelas)
                        except:
                            try:
                                produto.find_element(By.CSS_SELECTOR, "div[aria-label='Classificação: 1 de 5 estrelas.']").text.strip()
                                estrelas = 1
                                dict_produtos["estrelas"].append(estrelas)
                            except:
                                estrelas = "Não há estrelas"
                                dict_produtos["estrelas"].append(estrelas)

            print(f"{nome} / {preco} / {parcela} / {desconto} / {frete} / {avaliacoes} / {estrelas}")


            dict_produtos["marca"].append(nome)
            dict_produtos["preco"].append(preco)
            dict_produtos["parcela"].append(parcela)
            dict_produtos["desconto"].append(desconto)
            dict_produtos["frete"].append(frete)
            dict_produtos["avaliacoes"].append(avaliacoes)
        except Exception as e:
            print("Erro ao coletar dados:", e)

    try:
        botao_proximo = WebDriverWait(driver, 10).until(
            ec.element_to_be_clickable((By.CLASS_NAME, "nextLink"))
        )
        if botao_proximo:
            driver.execute_script("arguments[0].scrollIntoView();", botao_proximo)
            time.sleep(8)

            driver.execute_script("arguments[0].click();", botao_proximo)
            print(f"Indo para a página {pagina}")
            pagina += 1
            time.sleep(8)
        else:
            print("Você chegou na última página!")
            break

    except Exception as e:
        print("Erro ao tentar avançar para a próxima página", e)
        break

driver.quit()

df = pd.DataFrame(dict_produtos)
df.to_excel("Placas_de_Video.xlsx", index=False)

print(f"Arquivo 'Placas_de_Video' salvo com sucesso! ({len(df)} produtos capturados!)")