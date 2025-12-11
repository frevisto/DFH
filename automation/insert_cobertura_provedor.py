#
# SCRIPT PARA INCLUSÃO DE COBERTURAS GPON AUTOMATIZADA NO PORTAL SDWAN DA DFH
#

import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.alert import Alert
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import os

# ================================
# SETUP
# ================================
load_dotenv()
URL_SDWAN = os.getenv("URL_SDWAN")
SD_USER = os.getenv("SD_USER")
SD_PASS = os.getenv("SD_PASS")
URL_ADD_COBERTURA = os.getenv("URL_ADD_COBERTURA")

# ================================
# CARREGAR EXCEL (opcional)
# ================================
EXCEL_ARQUIVO = "../data/Pasta1.xlsx"
df = pd.read_excel(EXCEL_ARQUIVO, header=None)  # Sem cabeçalho
print("Total de linhas carregadas:", len(df))

# ================================
# INICIAR WEBDRIVER
# ================================
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)
wait = WebDriverWait(driver, 10)

# ================================
# LOGIN NO SDWAN
# ================================
driver.get(URL_SDWAN)

# Preencher usuário
campo_user = wait.until(EC.presence_of_element_located((By.ID, "set_Login")))
campo_user.clear()
campo_user.send_keys(SD_USER)

# Preencher senha
campo_pass = wait.until(EC.presence_of_element_located((By.ID, "set_pass")))
campo_pass.clear()
campo_pass.send_keys(SD_PASS)
time.sleep(1)

# Clicar no botão Login usando XPath filtrando pelo texto
bot_login = wait.until(
    EC.element_to_be_clickable(
        (By.XPATH, "//button[normalize-space(text())='Login']")
    )
)
driver.execute_script("arguments[0].click();", bot_login)


# Esperar redirecionamento para portal
wait.until(EC.url_contains("portal.html"))
print("Login concluído. Redirecionamento detectado.")

# ================================
# LOOP PRINCIPAL
# ================================
for index, row in df.iterrows():
    cidade = str(row[0]).strip()
    estado = str(row[1]).strip()
    valor_concat = f"{cidade} - {estado}"

    print(f"\n>>> Processando: {valor_concat}")

    try:
        # ----------------------------
        # Navegar para página de cobertura
        # ----------------------------
        driver.get(URL_ADD_COBERTURA)
        time.sleep(1)  # garantir carregamento

        # ----------------------------
        # 1) Preencher cb_nome
        # ----------------------------
        campo_nome = wait.until(EC.presence_of_element_located((By.ID, "cb_nome")))
        campo_nome.clear()
        campo_nome.send_keys(valor_concat)

        # ----------------------------
        # 2) Alterar tecnologia
        # ----------------------------
        campo_tec = driver.find_element(By.ID, "cb_Tecnologia")
        driver.execute_script("arguments[0].value = 'GPON (Fibra) Banda Larga';", campo_tec)

        # ----------------------------
        # 3) Alterar address
        # ----------------------------
        campo_addr = driver.find_element(By.ID, "address")
        campo_addr.clear()
        campo_addr.send_keys(valor_concat)

        # ----------------------------
        # 4) Localizar ponto
        # ----------------------------
        bot_localizar = driver.find_element(By.XPATH, "//input[@value='Localizar Ponto']")
        bot_localizar.click()
        wait.until(lambda d: d.find_element(By.ID, "cb_Cidade").get_attribute("value").strip() != "")

        # ----------------------------
        # 5) Incluir cobertura
        # ----------------------------
        bot_incluir = driver.find_element(By.XPATH, "//input[@value='Incluir Cobertura']")
        bot_incluir.click()

        # ----------------------------
        # 6) Aguardar alerta de sucesso
        # ----------------------------
        try:
            alerta = wait.until(EC.alert_is_present())
            print("Alerta recebido:", alerta.text)
            alerta.accept()
        except:
            print("⚠ Nenhum alerta encontrado — verificar comportamento do site")

        print(f"✔ {valor_concat}: Sucesso")

    except Exception as e:
        print(f"⚠ {valor_concat}: Falhou — {e}")
        continue  # continuar para próxima linha

print("\n=== Finalizado! ===")
driver.quit()