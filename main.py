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

    "sessao-id": "0108d3f7c99faa818e758d1c87e82cd3",

    "organizationid": "67",

    "domainkey": "spanionline.com.br",

    "accept": "application/json",

    "content-type": "application/json",

    "user-agent": "Mozilla/5.0"
}

# =====================================
# PRODUTO TESTE
# =====================================

produto_id = 803

# =====================================
# URL
# =====================================

url = f"https://services-beta.vipcommerce.com.br/api-admin/v1/org/67/filial/1/centro_distribuicao/36/loja/produtos/{produto_id}/detalhes"

# =====================================
# REQUEST
# =====================================

print("BUSCANDO PRODUTO...")

r = requests.get(
    url,
    headers=headers,
    timeout=10
)

print(f"STATUS: {r.status_code}")

dados = r.json()

produto = dados["data"]["produto"]

# =====================================
# CAMPOS
# =====================================

nome = produto.get("descricao", "")

ean = produto.get("codigo_barras", "")

setor = "MERCEARIA"

# =====================================
# PREÇOS
# =====================================

if produto.get("em_oferta"):

    preco_varejo = produto["oferta"].get("preco_antigo", "")

    preco_oferta = produto["oferta"].get("preco_oferta", "")

    oferta = "SIM"

else:

    preco_varejo = produto.get("preco", "")

    preco_oferta = ""

    oferta = "NÃO"

# =====================================
# LINK
# =====================================

link_slug = produto.get("link", "")

link_real = f"https://www.spanionline.com.br/produto/{produto_id}/{link_slug}"

# =====================================
# LINHAS
# =====================================

linhas = [[
    setor,
    nome,
    preco_varejo,
    preco_oferta,
    "",
    "",
    ean,
    oferta,
    "ABRIR",
    link_real
]]

# =====================================
# DATAFRAME
# =====================================

colunas = [
    "SETOR",
    "PRODUTO",
    "PREÇO VAREJO",
    "PREÇO OFERTA",
    "PREÇO ATACADO",
    "QTD ATACADO",
    "EAN",
    "OFERTA",
    "LINK",
    "LINK_REAL"
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
# OPENPYXL
# =====================================

wb = load_workbook(arquivo)

ws = wb["SPANI"]

# =====================================
# HEADER
# =====================================

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

# =====================================
# HYPERLINK
# =====================================

for row in range(2, ws.max_row + 1):

    link_cell = f"I{row}"

    url_cell = f"J{row}"

    ws[link_cell].hyperlink = ws[url_cell].value

    ws[link_cell].style = "Hyperlink"

# =====================================
# OCULTAR LINK REAL
# =====================================

ws.column_dimensions["J"].hidden = True

# =====================================
# FILTRO
# =====================================

ws.auto_filter.ref = ws.dimensions

# =====================================
# AJUSTAR COLUNAS
# =====================================

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

# =====================================
# SALVAR
# =====================================

wb.save(arquivo)

print("\n======================")
print("EXCEL GERADO")
print(nome)
print(preco_varejo)
print(preco_oferta)
print(oferta)
