import requests

# URL da API
url = "https://platform.senior.com.br/t/senior.com.br/bridge/1.0/rest/platform/user/queries/getUser"

# Cabeçalhos HTTP necessários
headers = {
    "Authorization": "Bearer RLeJJ6H4g325yGZe6NbklXP4CF3erMHG",
    "Content-Type": "application/json;charset=UTF-8",
    "Accept": "application/json, text/plain, */*"
}

# Corpo da requisição (vazio, mas obrigatório no POST)
payload = {
    'admin':True
}

try:
    response = requests.post(url, headers=headers, json=payload)
    
    if response.status_code == 200:
        print("✔ Requisição bem-sucedida!\n")
        print(response.json())  # Mostra os dados do usuário
    else:
        print(f"Erro {response.status_code}:")
        print(response.text)

    response = requests.post(url, headers=headers, json={})


    response = requests.post(url, headers=headers, json=payload)
    
    print("=== HEADERS ENVIADOS ===")
    for nome, valor in response.request.headers.items():
        print(f"{nome}: {valor}")

except Exception as e:
    print("Erro ao fazer requisição:", e)
