import cloudscraper

scraper = cloudscraper.create_scraper()

url = "https://services-beta.vipcommerce.com.br/api-admin/v1/org/67/filial/1/centro_distribuicao/36/loja/produtos/832/detalhes"

headers = {

    "accept": "application/json",

    "authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJ2aXBjb21tZXJjZSIsImF1ZCI6ImFwaS1hZG1pbiIsInN1YiI6IjZiYzQ4NjdlLWRjYTktMTFlOS04NzQyLTAyMGQ3OTM1OWNhMCIsInZpcGNvbW1lcmNlQ2xpZW50ZUlkIjpudWxsLCJpYXQiOjE3NzI3MTEyNDMsInZlciI6MSwiY2xpZW50IjpudWxsLCJvcGVyYXRvciI6bnVsbCwib3JnIjoiNjcifQ.5jbsro83AZ-4AG5jJsZKrbgeyocPa6n1vUQclalIR_HgF5FaxEFhJIcC0dggPwzdBzV0nFgPBJkk6ABFH6tDkQ",

    "organizationid": "67",

    "domainkey": "spanionline.com.br",

    "sessao-id": "0108d3f7c99faa818e758d1c87e82cd3",

    "origin": "https://www.spanionline.com.br",

    "referer": "https://www.spanionline.com.br/"

}

print("TESTANDO API SPANI...")

response = scraper.get(
    url,
    headers=headers
)

print("STATUS:")
print(response.status_code)

print("RESPOSTA:")
print(response.text)
