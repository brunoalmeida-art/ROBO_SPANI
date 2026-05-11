import requests
import json

# =========================================
# CONFIG
# =========================================

LOJA_ID = 1
CD_ID = 36
ORG_ID = 67

SESSION = "dc9e71d5-6b54-4c28-9f5c-0a0ab5dc316e"

# =========================================
# HEADERS
# =========================================

headers = {

    "accept": "application/json",

    "user-agent": "Mozilla/5.0",

    "OrganizationId": "67",

    "Application": "spanionline.com.br",

    "DomainKey": "spanionline.com.br",

    "session-id": SESSION

}

# =========================================
# TESTE
# =========================================

termo = "ar"

pagina = 1

url = (
    f"https://services-beta.vipcommerce.com.br/"
    f"api-admin/v1/org/{ORG_ID}/"
    f"filial/{LOJA_ID}/"
    f"centro_distribuicao/{CD_ID}/"
    f"loja/buscas/produtos/termo/{termo}"
    f"?page={pagina}"
    f"&&session={SESSION}"
)

print(url)

# =========================================
# REQUEST
# =========================================

r = requests.get(
    url,
    headers=headers,
    timeout=120
)

print(f"STATUS: {r.status_code}")

# =========================================
# JSON
# =========================================

js = r.json()

print(
    json.dumps(
        js,
        indent=2,
        ensure_ascii=False
    )
)

# =========================================
# PRODUTOS
# =========================================

produtos = js.get("produtos", [])

print(f"\nTOTAL PRODUTOS: {len(produtos)}")

# =========================================
# PRIMEIRO PRODUTO
# =========================================

if produtos:

    primeiro = produtos[0]

    print("\n======== PRIMEIRO PRODUTO ========\n")

    print(
        json.dumps(
            primeiro,
            indent=2,
            ensure_ascii=False
        )
    )

    print("\n======== CAMPOS ========\n")

    for chave in primeiro.keys():

        print(chave)
