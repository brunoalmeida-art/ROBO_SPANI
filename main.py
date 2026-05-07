import requests
import pandas as pd

# =====================================
# HEADERS
# =====================================

headers = {

    "authorization": "COLE_O_TOKEN_AQUI",

    "sessao-id": "340848be2780fc6b67960200ffa5a3fb",

    "organizationid": "67",

    "domainkey": "spanionline.com.br",

    "accept": "application/json",

    "content-type": "application/json",

    "user-agent": "Mozilla/5.0"
}

# =====================================
# URL
# =====================================

url = "https://services-beta.vipcommerce.com.br/api-admin/v1/org/67/filial/1/centro_distribuicao/6/loja/produtos/832/detalhes"

# =====================================
# REQUEST
# =====================================

r = requests.get(url, headers=headers)

print("STATUS:")
print(r.status_code)

print("\nJSON:")
print(r.text)

# =====================================
# EXCEL TESTE
# =====================================

linhas = [[
    "TESTE",
    r.status_code
]]

df = pd.DataFrame(
    linhas,
    columns=["INFO", "STATUS"]
)

df.to_excel(
    "SPANI.xlsx",
    index=False
)

print("\nEXCEL GERADO")
