import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

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

url = "https://services-beta.vipcommerce.com.br/api-admin/v1/org/67/filial/1/centro_distribuicao/36/loja/produtos/832/detalhes"

# =====================================
# REQUEST
# =====================================

r = requests.get(url, headers=headers)

print("STATUS:")
print(r.status_code)

dados = r.json()

print(dados)

# =====================================
# CAMPOS
# =====================================

nome = dados.get("descricao", "")

ean = dados.get("ean", "")

marca = dados.get("marca", "")

preco = dados.get("preco", "")

preco_antigo = dados.get("preco_de", "")

link = "https://www.spanionline.com.br/produto/832/leite-longa-vida-italac-1l-semidesnatado"

print("\nPRODUTO:")
print(nome)

print("\nPRECO:")
print(preco)

# =====================================
# DATAFRAME
# =====================================

linhas = [[
    "LATICINIOS",
    nome,
    preco,
    preco_antigo,
    "",
    "",
    ean,
    marca,
    link
]]

colunas = [
    "SETOR",
    "PRODUTO",
    "PREÇO VAREJO",
    "PREÇO ANTIGO",
    "PREÇO ATACADO",
    "QTD ATACADO",
    "EAN",
    "MARCA",
    "LINK"
]

df = pd.DataFrame(linhas, columns=colunas)

arquivo = "SPANI.xlsx"

with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:

    df.to_excel(
        writer,
        index=False,
        sheet_name="SPANI"
    )

# =====================================
# FORMATAÇÃO
# =====================================

wb = load_workbook(arquivo)

ws = wb["SPANI"]

cor_azul = PatternFill(
    start_color="1F4E78",
    end_color="1F4E78",
    fill_type="solid"
)

fonte_branca = Font(
    color="FFFFFF",
    bold=True
)

for cell in ws[1]:

    cell.fill = cor_azul
    cell.font = fonte_branca

ws.auto_filter.ref = ws.dimensions

for col in ws.columns:

    tamanho = 0

    letra = get_column_letter(col[0].column)

    for cell in col:

        try:

            if len(str(cell.value)) > tamanho:
                tamanho = len(str(cell.value))

        except:
            pass

    ws.column_dimensions[letra].width = tamanho + 5

wb.save(arquivo)

print("\nEXCEL GERADO")
