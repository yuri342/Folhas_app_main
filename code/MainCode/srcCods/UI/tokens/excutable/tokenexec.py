
# -*- coding: utf-8 -*-
import os
import sys
import json
import requests
from datetime import datetime, timedelta
from code.MainCode.UI.UI import TokenApp

# --------------------
# Paths compatíveis com PyInstaller
# --------------------
def resource_path(relative_path):
    """Retorna caminho correto para arquivos externos quando rodando como .exe ou .py."""
    if getattr(sys, 'frozen', False):
        # Quando é um .exe → usa o diretório onde o exe está
        base_path = os.path.dirname(sys.executable)
    else:
        # Quando é código Python normal → usa pasta do arquivo
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# --------------------
# Configurações
# --------------------
TOKEN_DIR = resource_path("tokens")
TOKEN_FILE = os.path.join(TOKEN_DIR, "tokens.json")
LOG_FILE = os.path.join(TOKEN_DIR, "log.txt")
REFRESH_THRESHOLD_DAYS = 4
REFRESH_URL = "https://platform.senior.com.br/t/senior.com.br/bridge/1.0/rest/platform/authentication/actions/refreshToken"

# --------------------
# Utilitários
# --------------------
def log(msg):
    os.makedirs(TOKEN_DIR, exist_ok=True)
    # rotação simples para não crescer indefinidamente
    try:
        if os.path.exists(LOG_FILE) and os.path.getsize(LOG_FILE) > 2 * 1024 * 1024:
            backup = LOG_FILE + ".1"
            if os.path.exists(backup):
                os.remove(backup)
            os.replace(LOG_FILE, backup)
    except Exception:
        pass

    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now().isoformat()}] {msg}\n")
    print(msg)

def carregar_token():
    log(f"Procurando TOKEN em: {TOKEN_FILE}")
    if not os.path.exists(TOKEN_FILE):
        return None
    try:
        with open(TOKEN_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        log(f"Erro ao ler token: {e}")
        return None

def salvar_token_atomic(token_data: dict):
    """Escrita atômica para evitar corrupção de arquivo em caso de interrupção."""
    os.makedirs(TOKEN_DIR, exist_ok=True)
    tmp = TOKEN_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(token_data, f, indent=4, ensure_ascii=False)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, TOKEN_FILE)

# --------------------
# Validação e expiração
# --------------------
def _parse_iso_datetime(s: str):
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None

def token_valido(token_data: dict):
    """
    Validação mínima:
    - JSON precisa ser dict
    - precisa ter 'access_token' OU 'refresh_token'
    - se houver 'expiration_date', não pode estar expirada
    - se NÃO houver 'expiration_date', então precisa ter 'refresh_token' para podermos renovar
    Retorna (is_valid, issues: list[str])
    """
    issues = []
    if not isinstance(token_data, dict):
        return False, ["JSON inválido (não é objeto)."]

    access = token_data.get("access_token")
    refresh = token_data.get("refresh_token")
    expiration_str = token_data.get("expiration_date")

    if not access and not refresh:
        issues.append("Faltam 'access_token' e 'refresh_token'.")

    if expiration_str:
        exp_dt = _parse_iso_datetime(expiration_str)
        if not exp_dt:
            issues.append("Campo 'expiration_date' inválido (não é ISO).")
        elif exp_dt <= datetime.now():
            issues.append("Token expirado.")
    else:
        # Sem expiração explícita, só é aceitável se houver refresh para permitir renovação
        if not refresh:
            issues.append("Sem 'expiration_date' e sem 'refresh_token' (não é possível renovar).")

    return (len(issues) == 0), issues

def precisa_renovar(expiration_str: str):
    """
    Retorna (renovar: bool, dias_restantes: int).
    Se não houver 'expiration_date' ou estiver inválida, preferimos renovar se houver refresh.
    """
    if not expiration_str:
        return True, -999
    exp_dt = _parse_iso_datetime(expiration_str)
    if not exp_dt:
        return True, -998
    dias_restantes = (exp_dt - datetime.now()).days
    return dias_restantes <= REFRESH_THRESHOLD_DAYS, dias_restantes

# --------------------
# Requisição de refresh
# --------------------
def renovar_token(refresh_token: str, timeout_sec: int = 20):
    headers = {
        "Authorization": f"Bearer {refresh_token}",
        "Content-Type": "application/json;charset=UTF-8",
        "Accept": "application/json",
    }
    payload = {"refreshToken": refresh_token}

    try:
        response = requests.post(REFRESH_URL, headers=headers, json=payload, timeout=timeout_sec)
        if response.status_code != 200:
            log(f"Falha ao renovar token ({response.status_code}): {response.text}")
            return None

        body = response.json()
        # Senior geralmente retorna {"jsonToken": "<string JSON com tokens>"}
        if "jsonToken" in body:
            try:
                novo = json.loads(body["jsonToken"])
            except Exception as e:
                log(f"jsonToken inválido: {e}")
                return None
        else:
            novo = body  # fallback: caso venha direto

        # Normalizar campos
        access = novo.get("access_token") or novo.get("accessToken")
        refresh = novo.get("refresh_token") or novo.get("refreshToken") or refresh_token
        expires_in = novo.get("expires_in") or novo.get("expiresIn")

        # Converter expires_in para int, se possível
        if isinstance(expires_in, (int, float)):
            expires_seconds = int(expires_in)
        else:
            try:
                expires_seconds = int(expires_in) if expires_in is not None else None
            except Exception:
                expires_seconds = None

        expiration_date = (datetime.now() + timedelta(seconds=expires_seconds)).isoformat() if expires_seconds else None

        resultado = {
            "access_token": access,
            "refresh_token": refresh,
            "expiration_date": expiration_date,
            "expires_in_seconds": expires_seconds,
        }
        return resultado

    except requests.RequestException as e:
        log(f"Erro de rede no refresh: {e}")
        return None
    except Exception as e:
        log(f"Erro ao processar resposta de refresh: {e}")
        return None

# --------------------
# Abrir UI somente quando necessário
# --------------------
def abrir_ui_e_recarregar():
    """
    Abre a UI para o usuário inserir/colar o token.
    A UI deve salvar em TOKEN_FILE (mesmo diretório).
    """
    log("Abrindo interface para captura/ajuste de token...")
    # UI.TokenApp deve aceitar um parâmetro opcional para diretório de salvamento, se você já implementou.
    # Caso não aceite, e a UI salva em TOKEN_FILE por padrão, basta instanciar sem argumentos:
    try:
        app = TokenApp(save_dir=TOKEN_DIR)  # se sua UI aceita save_dir
    except TypeError:
        app = TokenApp()  # fallback: UI antiga sem parâmetro
    app.mainloop()
    return carregar_token()

# --------------------
# Fluxo principal
# --------------------
def main():
    # Garantir diretório
    os.makedirs(TOKEN_DIR, exist_ok=True)

    token_data = carregar_token()

    # 1) Se não há token → abrir UI
    if token_data is None:
        log("Nenhum token encontrado.")
        token_data = abrir_ui_e_recarregar()
        if token_data is None:
            log("Nenhum token salvo pela interface. Encerrando.")
            return

    # 2) Validar token
    is_valid, issues = token_valido(token_data)
    if not is_valid:
        log("Token inválido: " + "; ".join(issues))
        token_data = abrir_ui_e_recarregar()
        if token_data is None:
            log("Nenhum token válido salvo pela interface. Encerrando.")
            return
        is_valid, issues = token_valido(token_data)
        if not is_valid:
            log("Token segue inválido após UI: " + "; ".join(issues))
            return

    # 3) Decidir renovação (NÃO acessar token_data["expiration_date"] sem checar)
    renovar, dias_restantes = precisa_renovar(token_data.get("expiration_date"))
    if renovar:
        log(f"Token próximo da expiração (dias restantes: {dias_restantes}). Renovando automaticamente...")
        refresh_value = token_data.get("refresh_token")
        if not refresh_value:
            log("Não há 'refresh_token' para renovar. Abrindo UI...")
            token_data = abrir_ui_e_recarregar()
            return

        novo_token = renovar_token(refresh_value)
        if novo_token:
            salvar_token_atomic(novo_token)
            log("Token renovado com sucesso!")
        else:
            log("Falha na renovação automática. Abrindo UI para ajuste manual...")
            token_data = abrir_ui_e_recarregar()
            if token_data is None:
                log("Nenhum token salvo. Encerrando.")
                return
    else:
        log(f"Token válido por mais {dias_restantes} dias.")

# Execução segura para PyInstaller
if __name__ == "__main__":
    main()
