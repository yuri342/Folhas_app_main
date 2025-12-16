import subprocess
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import json
import getpass  # pega nome do usuário automaticamente

# --- 1️⃣ Executar Chrome com Remote Debugging ---
user = getpass.getuser()  # pega usuário do Windows
chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"  # ajuste se necessário

profile_path = rf"C:\Users\{user}\AppData\Local\Google\Chrome\User Data"

# Comando para abrir o Chrome se não estiver aberto
subprocess.Popen([
    chrome_path,
    f"--remote-debugging-port=9222",
    f"--user-data-dir={profile_path}",
    "--profile-directory=Default"
])

# Pequena pausa para garantir que o Chrome abriu
time.sleep(5)

# --- 2️⃣ Conectar via Selenium usando CDP ---
chrome_options = Options()
chrome_options.debugger_address = "127.0.0.1:9222"

service = Service("chromedriver.exe")  # ajuste caminho se necessário
driver = webdriver.Chrome(service=service, options=chrome_options)

# --- 3️⃣ Abrir a plataforma Senior ---
driver.get("https://platform.senior.com.br/senior-x/#/")
time.sleep(3)  # esperar carregar cookies

# --- 4️⃣ Capturar cookies ---
cookies = driver.get_cookies()
driver.quit()

# --- 5️⃣ Filtrar o token ---
token_cookie = None
for c in cookies:
    if "token" in c["name"].lower():  # pega cookie que contenha "token"
        token_cookie = c["value"]
        break

if not token_cookie:
    print("❌ Nenhum token encontrado no navegador.")
else:
    os.makedirs("tokens", exist_ok=True)
    arquivo = os.path.join("tokens", "token_cookie.json")
    with open(arquivo, "w", encoding="utf-8") as f:
        json.dump({"token_cookie": token_cookie}, f, indent=4)
    print(f"✔ Token capturado e salvo em: {arquivo}")


