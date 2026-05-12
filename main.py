import requests
import json

# =========================================
# URL
# =========================================

url = (
    "https://services-beta.vipcommerce.com.br/"
    "api-admin/v1/org/67/"
    "filial/1/"
    "centro_distribuicao/36/"
    "loja/buscas/produtos/termo/ar"
    "?page=1"
    "&&session=dc9e71d5-6b54-4c28-9f5c-0a0ab5dc316e"
)

# =========================================
# HEADERS
# =========================================

headers = {

    "accept": "application/json",

    "user-agent": "Mozilla/5.0",

    "OrganizationId": "67",

    "DomainKey": "spanionline.com.br",

    "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJ2aXBjb21tZXJjZSIsImF1ZCI6ImFwaS1hZG1pbiIsInN1YiI6IjZiYzQ4NjdlLWRjYTktMTFlOS04NzQyLTAyMGQ3OTM1OWNhMCIsInZpcGNvbW1lcmNlQ2xpZW50ZUlkIjpudWxsLCJpYXQiOjE3Nzc5MDc4MTMsInZlciI6MSwiY2xpZW50IjpudWxsLCJvcGVyYXRvciI6bnVsbCwib3JnIjoiNjcifQ.mqyEyNRMBcY0rb4kWeNN0-xnEb8kus9i97w3IR6qjCCPdKEyBjUcZkF77_4KtKvHBI2cx25Fd8E9G4Q1cwsADw"
}

# =========================================
# REQUEST
# =========================================

r = requests.get(
    url,
    headers=headers,
    timeout=120
)

# =========================================
# STATUS
# =========================================

print("STATUS:", r.status_code)

# =========================================
# TEXTO BRUTO
# =========================================

print("\n======== RESPOSTA ========\n")

print(r.text[:10000])

# =========================================
# JSON
# =========================================

try:

    js = r.json()

    print("\n======== JSON FORMATADO ========\n")

    print(
        json.dumps(
            js,
            indent=2,
            ensure_ascii=False
        )
    )

    produtos = js.get("produtos", [])

    print(f"\nTOTAL PRODUTOS: {len(produtos)}")

    # =====================================
    # PRIMEIRO PRODUTO
    # =====================================

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

except Exception as e:

    print("\nERRO JSON\n")

    print(e)
