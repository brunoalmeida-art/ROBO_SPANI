import requests

url = "https://services-beta.vipcommerce.com.br/api-admin/v1/org/67/filial/1/centro_distribuicao/36/loja/produtos/3560/detalhes"

headers = {

    "Accept": "application/json",

    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",

    "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJ2aXBjb21tZXJjZSIsImF1ZCI6ImFwaS1hZG1pbiIsInN1YiI6IjZiYzQ4NjdlLWRjYTktMTFlOS04NzQyLTAyMGQ3OTM1OWNhMCIsInZpcGNvbW1lcmNlQ2xpZW50ZUlkIjpudWxsLCJpYXQiOjE3NzI3MTEyNDMsInZlciI6MSwiY2xpZW50IjpudWxsLCJvcGVyYXRvciI6bnVsbCwib3JnIjoiNjcifQ.5jbsro83AZ-4AG5jJsZKrbgeyocPa6n1vUQclalIR_HgF5FaxEFhJIcC0dggPwzdBzV0nFgPBJkk6ABFH6tDkQ",

    "Content-Type": "application/json",

    "Domainkey": "spanionline.com.br",

    "Organizationid": "67",

    "Origin": "https://www.spanionline.com.br",

    "Referer": "https://www.spanionline.com.br/",

    "Sec-Ch-Ua": '"Google Chrome";v="147", "Not?A_Brand";v="8", "Chromium";v="147"',

    "Sec-Ch-Ua-Mobile": "?0",

    "Sec-Ch-Ua-Platform": '"Windows"',

    "Sec-Fetch-Dest": "empty",

    "Sec-Fetch-Mode": "cors",

    "Sec-Fetch-Site": "cross-site",

    "User-Agent": "Mozilla/5.0"

}

print("TESTANDO API SPANI...")

response = requests.get(
    url,
    headers=headers,
    timeout=30
)

print("STATUS:")
print(response.status_code)

print("RESPOSTA:")
print(response.text)
