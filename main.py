import requests
import pandas as pd

# =====================================
# HEADERS
# =====================================

headers = {

    "authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJ2aXBjb21tZXJjZSIsImF1ZCI6ImFwaS1hZG1pbiIsInN1YiI6IjZiYzQ4NjdlLWRjYTktMTFlOS04NzQyLTAyMGQ3OTM1OWNhMCIsInZpcGNvbW1lcmNlQ2xpZW50ZUlkIjpudWxsLCJpYXQiOjE3NzI3MTEyNDMsInZlciI6MSwiY2xpZW50IjpudWxsLCJvcGVyYXRvciI6bnVsbCwib3JnIjoiNjcifQ.5jbsro83AZ-4AG5jJsZKrbgeyocPa6n1vUQclalIR_HgF5FaxEFhJIcC0dggPwzdBzV0nFgPBJkk6ABFH6tDkQ",

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
