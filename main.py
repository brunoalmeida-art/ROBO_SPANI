import requests
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
    # SELECIONAR MAUA 1
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

cookies = {

    "sessao-id": session_id,

    "vip-token": vip_token
}

# =========================================
# BUSCA
# =========================================

pagina = 1

categorias = {}

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

    print(f"\nPAGINA {pagina}")

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

    produtos = data.get(
        "data",
        {}
    ).get(
        "produtos",
        []
    )

    print("PRODUTOS:", len(produtos))

    if len(produtos) == 0:

        print("FIM")

        break

    # =====================================
    # MAPEAR CATEGORIAS
    # =====================================

    for p in produtos:

        secao_id = p.get(
            "secao_id",
            "SEM_ID"
        )

        classificacao = p.get(
            "classificacao_mercadologica_id",
            "SEM_CLASSIFICACAO"
        )

        descricao = p.get(
            "descricao",
            ""
        )

        if secao_id not in categorias:

            categorias[secao_id] = {

                "classificacao": classificacao,

                "produto_exemplo": descricao
            }

    pagina += 1

# =========================================
# RESULTADO
# =========================================

print("\n=====================================")
print("CATEGORIAS ENCONTRADAS")
print("=====================================\n")

for secao_id, info in sorted(categorias.items()):

    print(

        f"SECAO ID: {secao_id}"

        f" | CLASSIFICACAO: {info['classificacao']}"

        f" | EXEMPLO: {info['produto_exemplo']}"
    )

print("\nTOTAL CATEGORIAS:", len(categorias))
