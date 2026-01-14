import os
import pathlib as pl
from datetime import datetime
import json
import sys
import pathlib as pl
import requests
import certifi

SSL_VERIFY = certifi.where()
TABELASIDSNOMES = []
LINHAS = []

def get_tokens_json_path() -> pl.Path:

    # onde o exe está
    if getattr(sys, "frozen", False):
        base = pl.Path(sys.executable).parent
    else:
        base = pl.Path(__file__).resolve().parent

    candidatos = [

        # O CAMINHO QUE SEMPRE FUNCIONA NO EXE
        base / "tokens" / "tokens.json",

        # Caminho original no projeto (para rodar no Python)
        base.parent / "tokens" / "excutable" / "dist" / "tokens" / "tokens.json",
        base / "tokens" / "excutable" / "dist" / "tokens" / "tokens.json",
        base / "tokens" / "tokens.json",
    ]

    for c in candidatos:
        if c.exists():
            return c

    return None



def get_base_path():#
    if getattr(sys, 'frozen', False):
        # Rodando como .exe (PyInstaller)
        return pl.Path(sys.executable).parent
    else:
        # Rodando como script Python normal
        return pl.Path(__file__).parent

BASEPATH = get_base_path()
BASEPATH2 = get_tokens_json_path()
print(f"Base Path: {BASEPATH}")
print(f"Base Path 2: {BASEPATH2}")



def carregar_token(Token_FILE):
    if Token_FILE is None:
        print("ERRO: Token_FILE veio como None — get_tokens_json_path() não achou tokens.json")
        return None

    if not os.path.exists(Token_FILE):
        print("ERRO: Caminho não existe:", Token_FILE)
        return None

    try:
        with open(Token_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return None
    
def extrair_token():
    token_data = carregar_token(BASEPATH2)
    if token_data and "access_token" in token_data:
        return token_data["access_token"]
    return None

#REQUESTS - SENIOR

def request_tabela_salariais_Nomes(token):
    url = "https://platform.senior.com.br/t/senior.com.br/bridge/1.0/rest/hcm/remuneration/queries/wageScales"
    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {token}"
    }

    response = requests.get(url, headers=headers, verify=False)

    if response.status_code == 200:
        return response.json()
    else:
        print(f"Erro na requisição: {response.status_code} - {response.text}")
        return None
    

def addObjsTabelas(nome,id):
    obj = {
        "names": str(nome),
        "ids":  id
    }

    TABELASIDSNOMES.append(obj)

def listarJson(token):
    if not token:
        print("Token inválido.")
        return []

    jsonTabelas = request_tabela_salariais_Nomes(token)

    if not jsonTabelas or "wageScales" not in jsonTabelas:
        print("Erro ao carregar tabelas.")
        return []

    TABELASIDSNOMES.clear()  # evita duplicação

    for tabela in jsonTabelas.get("wageScales", []):
        tabelaId = tabela.get("id")
        nome = tabela.get("name")

        if tabelaId is not None and nome is not None:
            addObjsTabelas(nome, tabelaId)

    return TABELASIDSNOMES.copy()


def print_tabelas(lista):
    print(f"{'NOME':<40} | ID")
    print("-" * 90)

    for item in lista:
        print(f"{item['names']:<40} | {item['ids']}")


def request_tabela_salariais_Detalhes(token, tabela_id, keepAlive, wageRevisionId=None):
    url = "https://platform.senior.com.br/t/senior.com.br/bridge/1.0/rest/hcm/remuneration/queries/getWageScale"


    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
        
    }

    if keepAlive:
        headers["Connection"] = "keep-alive"
    
    payload = {
        "wageScaleId": tabela_id,
        "wageRevisionId": wageRevisionId
    }

    response = requests.post(url, headers=headers, json=payload, verify=False)

    if response.status_code == 200:
        return response.json()
    else:
        print(f"Erro na requisição: {response.status_code} - {response.text}")
        return None

def pegarTabelaRevisions(token, tabela_id=None):
    url = "https://platform.senior.com.br/t/senior.com.br/bridge/1.0/rest/hcm/remuneration/queries/getWageScale"

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
        "keep-alive": "true"
    }

    tabelas = []

    payload = {
        "wageScaleId": tabela_id,
    }

    response = requests.post(url, headers=headers, json=payload, verify=False)

    if response.status_code != 200:
        raise Exception(f"Erro na requisição: {response.status_code} - {response.text}")

    data = response.json()

    return data.get("revisions", [])

def pegarTabelaRevisionsG(token, tabela_id=None):
    revis = pegarTabelaRevisions(token, tabela_id=tabela_id)
    revision_nomes = [rev.get("startDate") for rev in revis]
    revision_ids = [rev.get("id") for rev in revis]

    URL_API = (
        "https://platform.senior.com.br/"
        "t/senior.com.br/bridge/1.0/rest/"
        "hcm/remuneration/queries/getWageScale"
    )

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
        "keep-alive": "true"
    }

    tabelas = []

    for tabela_rid in revision_ids:
        nomeRev = revision_nomes[revision_ids.index(tabela_rid)]
        payload = {"wageScaleId": tabela_id, "wageRevisionId": tabela_rid}
        response = requests.post(URL_API, json=payload, headers=headers, verify=False)

        if response.status_code != 200:
            print(f"Erro no RID {tabela_rid}")
            continue

        tabela = response.json()
        linhas = []

        for classe in tabela.get("classes", []):
            valores = [v.get("value") for v in classe.get("values", [])]
            nomes = [v.get("name") for v in classe.get("values", [])]

            linhas.append({
                "fonte": f"{nomeRev} - {tabela.get('name')}",
                "name": classe.get("name"),
                "valores": valores,
                "alfa": nomes
            })

        tabelas.append(linhas)

    return tabelas


#_---------------------------_---------------------------------#
# função que concatena e substitui e outra que em uma nova aba

#FUNÇÃO QUE FAZ VÁRIAS REQUISIÇÕES KEEP-ALIVE
def obter_dados_multiplos_ids(ids):
    import requests

    token = extrair_token()
    URL_API = (
        "https://platform.senior.com.br/"
        "t/senior.com.br/bridge/1.0/rest/"
        "hcm/remuneration/queries/getWageScale"
    )

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Connection": "keep-alive"
    }

    FILIAIS_SP = [
        "5003", "5042", "5046", "5061", "5064", "5070", "5067",
        "5018", "5041", "5043", "5024", "5028", "5040", "5047",
        "5048", "50030", "50420"
    ]

    tabelas = []

    for tabela_id in ids:
        payload = {"wageScaleId": tabela_id}
        response = requests.post(URL_API, json=payload, headers=headers, verify=False)

        if response.status_code != 200:
            print(f"Erro no ID {tabela_id}")
            continue

        tabela = response.json()

        is_sp = (
            tabela.get("name") == "Tabela Salarial SP"
            or tabela.get("id") == "74070F901FE74A358F8B1740EDF60F06"
        )

        # ---------- BASE DA TABELA ----------
        linhas_base = []
        for classe in tabela.get("classes", []):
            linhas_base.append({
                "name": classe.get("name"),
                "valores": [v.get("value") for v in classe.get("values", [])],
                "alfa": [v.get("name") for v in classe.get("values", [])]
            })

        # ---------- FLUXO SP ----------
        if is_sp:
            for filial in FILIAIS_SP:
                linhas = []
                for linha in linhas_base:
                    linhas.append({
                        "fonte": f"Tabela Salarial {filial}",
                        "name": linha["name"],
                        "valores": linha["valores"],
                        "alfa": linha["alfa"]
                    })
                tabelas.append(linhas)

        # ---------- FLUXO NORMAL ----------
        else:
            linhas = []
            for linha in linhas_base:
                linhas.append({
                    "fonte": tabela.get("name"),
                    "name": linha["name"],
                    "valores": linha["valores"],
                    "alfa": linha["alfa"]
                })
            tabelas.append(linhas)

    return tabelas

#FUNÇÂO RECURSIVA PARA ADICIONAR LINHAS
def adicionar_linha(nivel, valores, letras, font, idx):
    linha = {
        "fonte": str(font),
        "name": str(nivel),
        "valores": valores,
        "alfa": letras,
        "idx": idx
    }
    LINHAS.append(linha)
#------------------------------------------#


def criarTabelaNoPadrao(token, tabela_id, keepalive, revision_id=False):
    """
    Popula LINHAS.
    - keepalive=True  → dados já normalizados
    - keepalive=False → JSON bruto
    """
    # NÃO limpar LINHAS aqui
    # ===== MODO KEEPALIVE =====
    if keepalive and revision_id is False:
        tabelas = obter_dados_multiplos_ids(tabela_id)

        for idx, linhas_tabela in enumerate(tabelas):
            for linha in linhas_tabela:
                linha["idx"] = idx
                LINHAS.append(linha)

        return
    # ===== MODO KEEPALIVE COM REVISION ID =====
    if keepalive and revision_id:
        tabelas = pegarTabelaRevisionsG(token, tabela_id)


        for idx, linhas_tabela in enumerate(tabelas):
            for linha in linhas_tabela:
                linha["idx"] = idx
                LINHAS.append(linha)

        return

    # ===== MODO NORMAL =====
    tabela = request_tabela_salariais_Detalhes(
        token, tabela_id, keepAlive=False, wageRevisionId=revision_id
    )
    if not tabela:
        print("Nenhuma tabela encontrada.")
        return

    namejson = tabela.get("name")
    classes = tabela.get("classes", [])

    for linhas in classes:
        valores = [v.get("value") for v in linhas.get("values", [])]
        nomes = [v.get("name") for v in linhas.get("values", [])]

        adicionar_linha(
            linhas.get("name"),
            valores,
            nomes,
            namejson,
            idx=0
        )


from collections import defaultdict
def agrupar_linhas_por_tabela(LINHAS):
    tabelas = defaultdict(list)
    for linha in LINHAS:
        tabelas[linha["idx"]].append(linha)
    return list(tabelas.values())


def obter_workbook(caminho_excel: str):
    from openpyxl import load_workbook, Workbook
    """Retorna um workbook existente, ou cria um novo caso o arquivo não exista."""
    if os.path.exists(caminho_excel):
        print("📄 Abrindo Excel existente…")
        return load_workbook(caminho_excel)
    else:
        print("📄 Criando novo Excel…")
        return Workbook()

def criartabela(
    CaminhoPasta,
    tabela_id,
    nometabela,
    aba=False,
    nomeAba=None,
    onef=False,
    keepalive=False,
    revision_id=None,
    revisionIds = False
):
    import os
    from openpyxl.styles import Alignment, Font, Border, Side, PatternFill


    colunas = [
        "Nível", "STEP1", "STEP2", "STEP3", "STEP4",
        "100%", "STEP6", "STEP7", "STEP8", "STEP9", "FILIAL"
    ]

    token = extrair_token()

    # ================= BUSCA DE DADOS =================
    LINHAS.clear()

    if keepalive and not revisionIds:
        criarTabelaNoPadrao(token, tabela_id, True, revision_id=revisionIds)
        tabelas = agrupar_linhas_por_tabela(LINHAS)
    elif keepalive and revisionIds:
        criarTabelaNoPadrao(token, tabela_id, True, revision_id=revisionIds)
        tabelas = agrupar_linhas_por_tabela(LINHAS)
    else:
        criarTabelaNoPadrao(token, tabela_id, False, revision_id=revision_id)
        tabelas = [LINHAS]

    caminho_excel = os.path.join(CaminhoPasta, f"{nometabela}.xlsx")
    wb = obter_workbook(caminho_excel)

    # ================= ESTILOS =================
    fonteCabecalho = Font(name='Aptos Narrow', size=10, color='FFFFFF')
    fonteCabecalho2 = Font(name='Aptos Narrow', size=10, bold=True, color='FFFFFF')
    fontenivel = Font(name="Segoe UI Variable Small", size=7, bold=True, color="FFFFFF")
    fonteValores = Font(name="Segoe UI Variable Small", size=7, color="000000")

    prenchimento1 = PatternFill("solid", fgColor="E97132")
    prenchimento2 = PatternFill("solid", fgColor="7030A0")
    prenchimento3 = PatternFill("solid", fgColor="FEECEC")

    alinharCentro = Alignment(horizontal="center", vertical="center")

    bordaFina = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # ===== LOOP PRINCIPAL DE TABELAS =====

    for idx, LINHAS_TABELA in enumerate(tabelas, start=1):
        if aba:
            fonte_tabela = LINHAS_TABELA[0].get("fonte", f"ABA_{idx}")
            nome = nomeAba if nomeAba else fonte_tabela
            if nome in wb.sheetnames:
                nome = f"{nome}_{idx}"

            ws = wb.create_sheet(nome)
            escrever_cabecalho = True
            linha_base = 1
        else:
            ws = wb.active

            if ws.max_row <= 1:
                escrever_cabecalho = True
                linha_base = 1
            else:
                escrever_cabecalho = False
                linha_base = ws.max_row + 1

        # --- CABEÇALHO (APENAS SE NECESSÁRIO) ---
        if escrever_cabecalho:
            for col_idx, nome_col in enumerate(colunas, start=1):
                cell = ws.cell(row=linha_base, column=col_idx, value=nome_col)
                cell.font = fonteCabecalho2 if nome_col == "100%" else fonteCabecalho
                cell.fill = prenchimento2 if nome_col == "100%" else prenchimento1
                cell.alignment = alinharCentro
                cell.border = bordaFina

        linha_inicio = linha_base + (1 if escrever_cabecalho else 0)

        # --- DADOS ---
        for i, linha in enumerate(LINHAS_TABELA, start=linha_inicio):

            # NÍVEL
            cell = ws.cell(row=i, column=1, value=linha["name"])
            cell.font = fontenivel
            cell.fill = prenchimento2
            cell.alignment = alinharCentro
            cell.border = bordaFina

            # VALORES
            for j in range(2, len(colunas)):
                try:
                    valor = linha["valores"][j - 2]
                except (IndexError, TypeError):
                    valor = 0

                if not isinstance(valor, (int, float)):
                    valor = 0

                cell = ws.cell(row=i, column=j, value=valor)
                cell.fill = prenchimento3 if j == 6 else PatternFill("solid", fgColor="FFFFFF")
                cell.number_format = 'R$ #,##0.00'
                cell.font = fonteValores
                cell.alignment = alinharCentro
                cell.border = bordaFina

            # FILIAL (SEMPRE ÚLTIMA COLUNA)
            col_filial = len(colunas)
            fonte = linha.get("fonte", "")
            numero = fonte.split()[-1] if fonte else ""

            cell = ws.cell(row=i, column=col_filial, value=numero)
            cell.font = fontenivel
            cell.fill = prenchimento2
            cell.alignment = alinharCentro
            cell.border = bordaFina

    wb.save(caminho_excel)
    LINHAS.clear()

def gerar_tabela_SP_engessada(CaminhoPasta: str, tabela_id):
    """
    Regra fixa SP (engessada):
    - Busca UMA tabela salarial
    - Repete os mesmos dados
    - Cria 4 abas (5003, 5041, 5048, 50030)
    - Nome da aba = filial
    - Coluna FILIAL = filial
    """

    import os
    import requests
    from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

    FILIAIS_SP =[
        "5003",
        "5042",
        "5046",
        "5061",
        "5064",
        "5070",
        "5067",
        "5018",
        "5041",
        "5043",
        "5024",
        "5028",
        "5040",
        "5047",
        "5048",
        "50030",
        "50420"
    ]

    
    ARQUIVO = "Tabela Salarial SP.xlsx"

    token = extrair_token()
    if not token:
        raise Exception("Token inválido ou não encontrado.")

    # ===== CHAMADA DIRETA À API =====
    URL_API = (
        "https://platform.senior.com.br/"
        "t/senior.com.br/bridge/1.0/rest/"
        "hcm/remuneration/queries/getWageScale"
    )

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    payload = {"wageScaleId": tabela_id}
    response = requests.post(URL_API, json=payload, headers=headers, verify=False)

    if response.status_code != 200:
        raise Exception(f"Erro ao buscar tabela: {response.status_code}")

    data = response.json()
    classes = data.get("classes", [])

    if not classes:
        raise Exception("Tabela retornou sem classes.")

    # ===== NORMALIZAÇÃO DOS DADOS =====
    dados = []
    for classe in classes:
        dados.append({
            "nivel": classe.get("name"),
            "valores": [v.get("value") for v in classe.get("values", [])]
        })

    # ===== EXCEL =====
    caminho_excel = os.path.join(CaminhoPasta, ARQUIVO)
    wb = obter_workbook(caminho_excel)

    colunas = [
        "Nível", "STEP1", "STEP2", "STEP3", "STEP4",
        "100%", "STEP6", "STEP7", "STEP8", "STEP9", "FILIAL"
    ]

    # ===== ESTILOS =====
    fonteCabecalho = Font(name='Aptos Narrow', size=10, color='FFFFFF')
    fonteCabecalho2 = Font(name='Aptos Narrow', size=10, bold=True, color='FFFFFF')
    fontenivel = Font(name="Segoe UI Variable Small", size=7, bold=True, color="FFFFFF")
    fonteValores = Font(name="Segoe UI Variable Small", size=7, color="000000")

    prenchimento1 = PatternFill("solid", fgColor="E97132")
    prenchimento2 = PatternFill("solid", fgColor="7030A0")
    prenchimento3 = PatternFill("solid", fgColor="FEECEC")

    alinharCentro = Alignment(horizontal="center", vertical="center")

    bordaFina = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # ===== UMA ABA POR FILIAL =====
    for filial in FILIAIS_SP:

        if filial in wb.sheetnames:
            del wb[filial]

        ws = wb.create_sheet(filial)

        # CABEÇALHO
        for col_idx, nome_col in enumerate(colunas, start=1):
            cell = ws.cell(row=1, column=col_idx, value=nome_col)
            cell.font = fonteCabecalho2 if nome_col == "100%" else fonteCabecalho
            cell.fill = prenchimento2 if nome_col == "100%" else prenchimento1
            cell.alignment = alinharCentro
            cell.border = bordaFina

        # DADOS
        for i, linha in enumerate(dados, start=2):

            # NÍVEL
            cell = ws.cell(row=i, column=1, value=linha["nivel"])
            cell.font = fontenivel
            cell.fill = prenchimento2
            cell.alignment = alinharCentro
            cell.border = bordaFina

            # VALORES
            for j in range(2, len(colunas)):
                valor = linha["valores"][j - 2] if j - 2 < len(linha["valores"]) else 0
                if not isinstance(valor, (int, float)):
                    valor = 0

                cell = ws.cell(row=i, column=j, value=valor)
                cell.fill = prenchimento3 if j == 6 else PatternFill("solid", fgColor="FFFFFF")
                cell.number_format = 'R$ #,##0.00'
                cell.font = fonteValores
                cell.alignment = alinharCentro
                cell.border = bordaFina

            # FILIAL
            cell = ws.cell(row=i, column=len(colunas), value=filial)
            cell.font = fontenivel
            cell.fill = prenchimento2
            cell.alignment = alinharCentro
            cell.border = bordaFina

    wb.save(caminho_excel)









#INPUT DE TEXTO
#INPUT DE CAMINHO DE PASTA
#INPUT DE DROPDOWN 
#INPUT fALSE/tRUE
#INPUT NOME aba textO