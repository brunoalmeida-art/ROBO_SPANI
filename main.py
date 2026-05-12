import requests
import pandas as pd
from playwright.sync_api import sync_playwright

# =========================================
# CONFIG
# =========================================

BASE_API = "https://services-beta.vipcommerce.com.br"

LOJA_URL = "https://www.spanionline.com.br"

BUSCA = "a"

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

cookies = {

    "sessao-id": session_id,

    "vip-token": vip_token
}

# =========================================
# API
# =========================================

url = (

    f"{BASE_API}"

    f"/api-admin/v1/org/67"

    f"/filial/1"

    f"/centro_distribuicao/36"

    f"/loja/buscas/produtos/termo/{BUSCA}"

    f"?page=1"

    f"&session={session_id}"
)

print(url)

r = requests.get(

    url,

    headers=headers,

    cookies=cookies,

    timeout=120
)

print("STATUS:", r.status_code)

data = r.json()

# =========================================
# PRODUTOS
# =========================================

produtos = data.get(
    "data",
    {}
).get(
    "produtos",
    []
)

print("TOTAL PRODUTOS:", len(produtos))

# =========================================
# DEBUG PRODUTO
# =========================================

if len(produtos) > 0:

    print("JSON PRODUTO COMPLETO:")
    print(produtos[0])

print("FINALIZADO")
