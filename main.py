import requests
import pandas as pd
import time

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
# IDS TESTE
# =====================================

produtos_ids = [
    832,
    823,
    9587,
    9584
]

# =====================================
# LISTA FINAL
# =====================================

linhas = []

# =====================================
# LOOP
# =====================================

for produto_id in produtos_ids:

    try:

        print(f"\nBUSCANDO PRODUTO {produto_id}")

        url = f"https://services-beta.vipcommerce.com.br/api-admin/v1/org/67/filial/1/centro_distribuicao/36/loja/produtos/{produto_id}/detalhes"

        r = requests.get(url, headers=headers)

        print(f"STATUS: {r.status_code}")

        if r.status_code != 200:
            continue

        dados = r.json()

        produto = dados["data"]["produto"]

        # =====================================
        # CAMPOS
        # =====================================

        nome = produto.get("descricao", "")

        preco = produto.get("preco", "")

        preco_antigo = produto.get("preco_original", "")

        ean = produto.get("codigo_barras", "")

        marca = produto.get("marca", "")

        oferta = "SIM" if produto.get("em_oferta") else "NAO"

        setor = "MERCEARIA"

        link_slug = produto.get("link", "")

        link = f"https://www.spanionline.com.br/produto/{produto_id}/{link_slug}"

        # =====================================
        # LINHA
        # =====================================

        linhas.append([
            setor,
            nome,
            preco,
            preco_antigo if preco_antigo != preco else "",
            "",
            "",
            ean,
            marca,
            oferta,
            link
        ])

        print(nome)

        time.sleep(1)

    except Exception as e:

        print(f"ERRO PRODUTO {produto_id}")
        print(e)

# =====================================
# DATAFRAME
# =====================================

colunas = [
    "SETOR",
    "PRODUTO",
    "PREÇO VAREJO",
    "PREÇO ANTIGO",
    "PREÇO ATACADO",
    "QTD ATACADO",
    "EAN",
    "MARCA",
    "OFERTA",
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

# COR HEADER

cor_azul = PatternFill(
    start_color="1F4E78",
    end_color="1F4E78",
    fill_type="solid"
)

# FONTE HEADER

fonte_branca = Font(
    color="FFFFFF",
    bold=True
)

# HEADER

for cell in ws[1]:

    cell.fill = cor_azul
    cell.font = fonte_branca

# FILTRO

ws.auto_filter.ref = ws.dimensions

# AJUSTAR COLUNAS

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

# SALVAR

wb.save(arquivo)

print("\n======================")
print("EXCEL GERADO")
print(f"TOTAL PRODUTOS: {len(df)}")
