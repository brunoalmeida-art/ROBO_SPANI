import requests
import json

url = "https://services-beta.vipcommerce.com.br/api-admin/v1/org/67/filial/1/centro_distribuicao/36/loja/buscas/produtos/termo/ar?page=1&&session=dc9e71d5-6b54-4c28-9f5c-0a0ab5dc316e"

headers = {
    "accept": "application/json",
    "user-agent": "Mozilla/5.0"
}

r = requests.get(
    url,
    headers=headers,
    timeout=120
)

print("STATUS:", r.status_code)

print(r.text[:5000])
