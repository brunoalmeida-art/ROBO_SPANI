import requests

url = "https://services-beta.vipcommerce.com.br/api-admin/v1/org/67/filial/1/centro_distribuicao/36/loja/produtos/823/detalhes"

headers = {

    "User-Agent": "Mozilla/5.0",

    "Accept": "application/json",

    "organizationid": "67",

    "origin": "https://www.spanionline.com.br",

    "referer": "https://www.spanionline.com.br/"

}

print("TESTANDO API SPANI...")

try:

    response = requests.get(
        url,
        headers=headers,
        timeout=30
    )

    print("STATUS:")
    print(response.status_code)

    print("RESPOSTA:")
    print(response.text)

except Exception as e:

    print("ERRO:")
    print(e)
