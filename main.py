import requests
import pandas as pd
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# =========================================
# CONFIG
# =========================================

BASE_API = "https://services-beta.vipcommerce.com.br"

LOJA_URL = "https://www.spanionline.com.br"

BUSCA = "a"

OUTPUT = "SPANI.xlsx"

# =========================================
# PLAYWRIGHT
# =========================================

print("ABRINDO SPANI...")

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=True
    )

    context = browser.new_context()

    page = context.new_page()

    page.goto(
        LOJA_URL,
        timeout=120000
    )

    page.wait_for_timeout(10000)

    # =====================================
    # SELECIONAR LOJA
    # =====================================

    try:

        print("SELECIONANDO MAUA 1")

        page.locator(
            ".vip-endereco-wrapper"
        ).click()

        page.wait_for_timeout(3000)

        page.locator(
            "text=Spani Mauá 1"
        ).click()

        page.wait_for_timeout(10000)

    except Exception as e:

        print(
            "ERRO LOJA:",
            e
        )

    # =====================================
    # ENDERECO
    # =====================================

    try:

        endereco = page.locator(
            ".vip-endereco-wrapper"
        ).inner_text()

        print("LOJA:", endereco)

    except:

        print("NAO ACHOU ENDERECO")

    # =====================================
    # COOKIES
    # =====================================

    cookies_play = context.cookies()

    session_id = ""

    vip_token = ""

    for c in cookies_play:

        print(c["name"])

        if c["name"] == "sessao-id":

            session_id = c["value"]

        if c["name"] == "vip-token":

            vip_token = c["value"]

    browser.close()

# =========================================
# VALIDAR
# =========================================

print("SESSION:", session_id)

print("VIP TOKEN:", vip_token)

if session_id == "":

    raise Exception(
        "SESSAO VAZIA"
    )

if vip_token == "":

    raise Exception(
        "VIP TOKEN VAZIO"
    )

# =========================================
# HEADERS
# =========================================

headers = {

    "accept": "application/json",

    "origin": "https://www.spanionline.com.br",

    "referer": "https://www.spanionline.com.br/",

    "Authorization": f"Bearer {vip_token}",

    "organizationid": "67",

    "filialid": "1",

    "centrodistribuicaoid": "36",

    "user-agent": "Mozilla/5.0"
}

# =========================================
# COOKIES REQUEST
# =========================================

cookies = {

    "sessao-id": session_id,

    "vip-token": vip_token
}

# =========================================
# LOOP PAGINAS
# =========================================

pagina = 1

todos = []

while True:

    url = (

        f"{BASE_API}"

        f"/api-admin/v1/org/67"

        f"/filial/1"

        f"/centro_distribuicao/36"

        f"/loja/buscas/produtos/termo/{BUSCA}"

        f"?page={pagina}"

        f"&session={session_id}"
    )

    print(f"PAGINA {pagina}")

    print(url)

    r = requests.get(

        url,

        headers=headers,

        cookies=cookies,

        timeout=120
    )

    print("STATUS:", r.status_code)

    if r.status_code != 200:

        print(r.text)

        break

    data = r.json()

    # =====================================
    # DEBUG JSON
    # =====================================

    print("JSON COMPLETO:")
    print(data)

    # =====================================
    # PRODUTOS
    # =====================================

    produtos = data.get(
        "data",
        []
    )

    print("TIPO PRODUTOS:", type(produtos))

    print("PRODUTOS:", len(produtos))

    # =====================================
    # PARAR DEBUG
    # =====================================

    break

print("FINALIZADO DEBUG")
