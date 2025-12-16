import os
import pathlib as pl
from datetime import datetime
import json
import sys
import pathlib as pl
import requests

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

    response = requests.get(url, headers=headers)

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


def request_tabela_salariais_Detalhes(token, tabela_id, keepAlive):
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
    }

    response = requests.post(url, headers=headers, json=payload)

    if response.status_code == 200:
        return response.json()
    else:
        print(f"Erro na requisição: {response.status_code} - {response.text}")
        return None

#_---------------------------_---------------------------------#
# função que concatena e substitui e outra que em uma nova aba


def obter_dados_multiplos_ids(ids):
    """
    Faz UMA request para vários IDs e retorna dados normalizados.
    """
    resultados = []
    token = extrair_token()

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Connection": "keep-alive"
    }

    for id in ids:
        print(id)
        payload = {
            "wageScaleId": id  # lista de IDs
        }

        URL_API = "https://platform.senior.com.br/t/senior.com.br/bridge/1.0/rest/hcm/remuneration/queries/getWageScale"
        response = requests.post(URL_API, json=payload, headers=headers)

        resultados.append(response.json())

    print("Gerado com Sucesso")
    return resultados

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

def criarTabelaNoPadrao(token, tabela_id, keepalive):
    if keepalive:
        tabelasjson = obter_dados_multiplos_ids(tabela_id)  # lista de dicts
    else:
        tabelasjson = request_tabela_salariais_Detalhes(
            token, tabela_id, keepAlive=keepalive
        )
        tabelasjson = [tabelasjson]  # transforma em lista

    if not tabelasjson:
        print("Nenhuma tabela encontrada.")
        return

    for idx, tabela in enumerate(tabelasjson):
        namejson = tabela.get("name")
        classes = tabela.get("classes", [])

        for linhas in classes:
            name = linhas.get("name")
            nomes = []
            valores = []

            for value in linhas.get("values", []):
                valores.append(value.get("value"))
                nomes.append(value.get("name"))

            adicionar_linha(
                name,
                valores,
                nomes,
                namejson,
                idx  # índice da tabela
            )



def obter_workbook(caminho_excel: str):
    from openpyxl import load_workbook, Workbook
    """Retorna um workbook existente, ou cria um novo caso o arquivo não exista."""
    if os.path.exists(caminho_excel):
        print("📄 Abrindo Excel existente…")
        return load_workbook(caminho_excel)
    else:
        print("📄 Criando novo Excel…")
        return Workbook()

def criartabela(CaminhoPasta, tabela_id, nometabela, aba=False, nomeAba=None, onef=False, keepalive=False):
    from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
    from openpyxl import Workbook

    colunas = ["Nível", "STEP1", "STEP2", "STEP3", "STEP4", "100%", "STEP6", "STEP7", "STEP8", "STEP9", "FILIAL"]
    token = extrair_token()

    if keepalive == True:
        criarTabelaNoPadrao(token, tabela_id, True)
    else:
        criarTabelaNoPadrao(token, tabela_id, False)
    #------------------------------------#   
    # CRIAR ARQUIVO EXCEL
    #CRIAR ABA
    caminho_excel = f"{CaminhoPasta}/{nometabela}.xlsx"

    wb = obter_workbook(caminho_excel)

    # Se for criar nova aba
    if aba:
        if keepalive == True:
            for indx in i
        if not nomeAba:
            nomeAba = "ABANOVA"
        ws = wb.create_sheet(nomeAba)
    else:
        ws = wb.active


    # --- ESTILOS DO CABEÇALHO ---

    # Fonte (precisa existir no sistema; se não existir, o Excel ignora)
    fonteCabecalho = Font(
        name='Aptos Narrow',
        size=10,
        bold=False,
        color='FFFFFF'
    )

    fonteCabecalho2 = Font(
        name='Aptos Narrow',
        size=10,
        bold=True,
        color='FFFFFF'
    )

    fontenivel = Font(
        name="Segoe UI Variable Small",
        size=7,
        bold=True,
        color="FFFFFF"
    )

    fonteValores = Font(
        name="Segoe UI Variable Small",
        size=7,
        bold=False,
        color="000000"
    )

    #prenchimento
    prenchimento1 = PatternFill(
        "solid",
        fgColor="E97132"
    )

    prenchimento2 = PatternFill(
        "solid",
        fgColor="7030A0"
    )

    prenchimento3 = PatternFill(
        "solid",
        fgColor="FEECEC"
    )

    # Alinhamento
    alinharCentro = Alignment(
        horizontal="center",
        vertical="center"
    )

    # Bordas
    bordaFina = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # --- APLICAR CABEÇALHO ---
    for col_idx, nome_col in enumerate(colunas, start=1):
        cell = ws.cell(row=1, column=col_idx, value=nome_col)
        if nome_col ==  "100%":
            cell.font = fonteCabecalho2
            cell.fill = prenchimento2
        else:
            cell.font = fonteCabecalho
            cell.fill = prenchimento1

        cell.alignment = alinharCentro
        cell.border = bordaFina
    
    #--- APLICAR DADOS ---#
    #criação da coluna da filial
    for linhaIdx , linhaValue in enumerate(LINHAS, start=2):
        ultima_coluna = ws[2][-1].column
        primeira_vazia = ultima_coluna + 1
        fonte = linhaValue.get("fonte")
        numero = fonte.split()[-1]
        cell = ws.cell(row=linhaIdx, column=ultima_coluna, value=numero)
        cell.font = fontenivel
        cell.fill = prenchimento2
        cell.alignment = alinharCentro
        cell.border = bordaFina

    #criação dos niveis
    for linhaIdx , linhaValue in enumerate(LINHAS, start=2):
        nivel = linhaValue.get("name")
        cell = ws.cell(row=linhaIdx, column=1, value=nivel )
        cell.font = fontenivel
        cell.fill = prenchimento2
        cell.alignment = alinharCentro
        cell.border = bordaFina
    
    #row = linha
    #collum = coluna
    #criaçõo da coluna com os valores - Dinheiro
    for linhaidx, linhaValue in enumerate(LINHAS, start=2):
        Valores = linhaValue.get("valores")
        for valoridx, valor in enumerate(Valores, start=2):
            cell = ws.cell(row=linhaidx, column=valoridx, value=valor)
            if valoridx == 6:
                cell.fill = prenchimento3
            else:
                cell.fill = PatternFill("solid", fgColor="FFFFFF")
            cell.alignment = alinharCentro
            cell.number_format = 'R$ #,##0.00'
            cell.font = fonteValores
            cell.border = bordaFina


    caminho = os.path.join(CaminhoPasta, f"{nometabela}.xlsx")
    wb.save(caminho)



#INPUT DE TEXTO
#INPUT DE CAMINHO DE PASTA
#INPUT DE DROPDOWN 
#INPUT fALSE/tRUE
#INPUT NOME aba textO