import requests

url = "https://services-beta.vipcommerce.com.br/api-admin/v1/org/67/filial/1/centro_distribuicao/36/loja/produtos/832/detalhes"

headers = {

    "accept": "application/json",

    "accept-encoding": "gzip, deflate, br, zstd",

    "accept-language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",

    "authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJ2aXBjb21tZXJjZSIsImF1ZCI6ImFwaS1hZG1pbiIsInN1YiI6IjZiYzQ4NjdlLWRjYTktMTFlOS04NzQyLTAyMGQ3OTM1OWNhMCIsInZpcGNvbW1lcmNlQ2xpZW50ZUlkIjpudWxsLCJpYXQiOjE3NzI3MTEyNDMsInZlciI6MSwiY2xpZW50IjpudWxsLCJvcGVyYXRvciI6bnVsbCwib3JnIjoiNjcifQ.5jbsro83AZ-4AG5jJsZKrbgeyocPa6n1vUQclalIR_HgF5FaxEFhJIcC0dggPwzdBzV0nFgPBJkk6ABFH6tDkQ",

    "content-type": "application/json",

    "domainkey": "spanionline.com.br",

    "organizationid": "67",

    "origin": "https://www.spanionline.com.br",

    "priority": "u=1, i",

    "referer": "https://www.spanionline.com.br/",

    "sec-ch-ua": '"Google Chrome";v="147", "Not.A/Brand";v="8", "Chromium";v="147"',

    "sec-ch-ua-mobile": "?0",

    "sec-ch-ua-platform": '"Windows"',

    "sec-fetch-dest": "empty",

    "sec-fetch-mode": "cors",

    "sec-fetch-site": "cross-site",

    "sessao-id": "0108d3f7c99faa818e758d1c87e82cd3",

    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36"

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
