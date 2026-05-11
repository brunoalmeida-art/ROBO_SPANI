import requests
import json

# =========================================
# TOKEN
# =========================================

TOKEN = "COLE_AQUI_O_BEARER_COMPLETO"

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

    "Authorization": f"Bearer {TOKEN}",

    "Content-Type": "application/json",

    "Application": "spanionline.com.br",

    "DomainKey": "spanionline.com.br",

    "OrganizationId": "67",

    "session-id": SESSION,

    "Origin": "https://www.spanionline.com.br",

    "Referer": "https://www.spanionline.com.br/",

    "User-Agent": "Mozilla/5.0"

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

print(r.text)
