import requests

url = "https://services-beta.vipcommerce.com.br/api-admin/v1/org/67/filial/1/centro_distribuicao/36/loja/produtos/823/detalhes"

headers = {

    "user-agent": "Mozilla/5.0",

    "accept": "application/json",

    "content-type": "application/json",

    "organizationid": "67",

    "authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJ2aXBjb21tZXJjZSIsImF1ZCI6ImFwaS1hZG1pbiIsInN1YiI6IjZiYzQ4NjdlLWRjYTktMTFlOS04NzQyLTAyMGQ3OTM1OWNhMCIsInZpcGNvbW1lcmNlQ2xpZW50ZUlkIjpudWxsLCJpYXQiOjE3NzI3MTEyNDMsInZlciI6MSwiY2xpZW50IjpudWxsLCJvcGVyYXRvciI6bnVsbCwib3JnIjoiNjcifQ.5jbsro83AZ-4AG5jJsZKrbgeyocPa6n1vUQclalIR_HgF5FaxEFhJIcC0dggPwzdBzV0nFgPBJkk6ABFH6tDkQ"

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
