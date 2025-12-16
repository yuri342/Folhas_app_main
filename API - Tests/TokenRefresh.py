import json
from datetime import datetime, timedelta
import os
import json
import requests
from datetime import datetime, timedelta

def refresh_token(refresh_token):
    url = "https://platform.senior.com.br/t/senior.com.br/bridge/1.0/rest/platform/authentication/actions/refreshToken"

    headers = {
        "Authorization": f"Bearer {refresh_token}",
        "Content-Type": "application/json;charset=UTF-8",
        "Accept": "application/json"
    }

    payload = {
        "refreshToken": refresh_token
    }

    response = requests.post(url, headers=headers, json=payload)

    return response.json()

def verificar_data():

    # --- Carregar token salvo ---
    with open("tokens/tokens.json", "r", encoding="utf-8") as f:
        token_data = json.load(f)

    # --- Ler data de expiração ---
    expiration_str = token_data.get("expiration_date")
    if not expiration_str:
        raise ValueError("Data de expiração não encontrada no JSON.")

    expiration_date = datetime.fromisoformat(expiration_str)
    current_date = datetime.now()

    # --- Verificar se faltam 4 dias ou menos ---
    dias_restantes = (expiration_date - current_date).days

    if dias_restantes <= 4:
        print(f"⚠ O token vai expirar em {dias_restantes} dias. Precisa renovar!")
        return 1
    else:
        print(f"✔ Token ainda válido por {dias_restantes} dias.")
        return 0

def refresh_token_json(refreshTokenValue):

    api_response = refresh_token(refreshTokenValue)

    # extrair JSON interno
    token_data = json.loads(api_response["jsonToken"])

    # calcular expiração correta
    expires_seconds = token_data["expires_in"]
    expiration_date = datetime.now() + timedelta(seconds=expires_seconds)

    # objeto final
    result = {
        "access_token": token_data["access_token"],
        "refresh_token": token_data["refresh_token"],
        "expiration_date": expiration_date.isoformat(),
        "expires_in_seconds": expires_seconds
    }

    # salvar automaticamente
    os.makedirs("tokens", exist_ok=True)
    file_path = os.path.join("tokens", "tokens.json")

    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=4, ensure_ascii=False)

    print("✔ Novo token salvo com sucesso!")
    print(json.dumps(result, indent=4, ensure_ascii=False))


if __name__ == "__main__":
   value = verificar_data()
   if value == 1:
     with open("tokens/tokens.json", "r", encoding="utf-8") as f:
          token_data = json.load(f)
     refresh_token_value = token_data.get("refresh_token")
     if not refresh_token_value:
          raise ValueError("Refresh token não encontrado no JSON.")
     refresh_token_json(refresh_token_value)
   else:
        print("Nenhuma ação necessária.")