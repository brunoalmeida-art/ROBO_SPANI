import requests
import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

import smtplib
import os

from email.message import EmailMessage

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
# BUSCAS A-Z
# =====================================

buscas = list("abcdefghijklmnopqrstuvwxyz")

# =====================================
# LISTA FINAL
# =====================================

dados_finais = []

# =====================================
# LOOP BUSCAS
# =====================================

for termo in buscas:

    pagina = 1

    while True:

        print(f"\nBUSCA: {termo} | PAGINA: {pagina}")

        url = f"https://services-beta.vipcommerce.com.br/api-admin/v1/org/67/filial/1/centro_distribuicao/36/loja/buscas/produtos/termo/{termo}?page={pagina}"

        r = requests.get(
            url,
            headers=headers,
            timeout=30
        )

        print(f"STATUS: {r.status_code}")

        if r.status_code != 200:
            break

        dados = r.json()

        produtos = dados.get("data", [])

        if not produtos:
            break

        # =====================================
        # LOOP PRODUTOS
        # =====================================

        for produto in produtos:

            try:

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

                    preco_atacado = produto["oferta"].get("preco_oferta", "")

                    qtd_atacado = produto["oferta"].get("quantidade_minima", "")

                else:

                    preco_varejo = produto.get("preco", "")

                    preco_oferta = ""

                    oferta = "NÃO"

                    preco_atacado = ""

                    qtd_atacado = ""

                # =====================================
                # LINK
                # =====================================

                produto_id = produto.get("produto_id", "")

                link_slug = produto.get("link", "")

                link_real = f"https://www.spanionline.com.br/produto/{produto_id}/{link_slug}"

                # =====================================
                # APPEND
                # =====================================

                dados_finais.append([

                    setor,
                    nome,
                    preco_varejo,
                    preco_oferta,
                    preco_atacado,
                    qtd_atacado,
                    ean,
                    oferta,
                    "ABRIR",
                    link_real

                ])

            except Exception as erro:

                print("ERRO PRODUTO")
                print(erro)

        pagina += 1

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

df = pd.DataFrame(
    dados_finais,
    columns=colunas
)

# =====================================
# REMOVER DUPLICADOS
# =====================================

df = df.drop_duplicates(
    subset=["EAN"]
)

# =====================================
# ORDENAR
# =====================================

df = df.sort_values(
    by="PRODUTO"
)

# =====================================
# EXCEL
# =====================================

arquivo = "SPANI.xlsx"

with pd.ExcelWriter(
    arquivo,
    engine="openpyxl"
) as writer:

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

print("\n===================")
print("EXCEL GERADO")
print(f"TOTAL: {len(df)}")

# =====================================
# EMAIL
# =====================================

EMAIL_USER = os.getenv("EMAIL_USER")

EMAIL_PASS = os.getenv("EMAIL_PASS")

EMAIL_TO = os.getenv("EMAIL_TO")

msg = EmailMessage()

msg["Subject"] = "ROBO SPANI"

msg["From"] = EMAIL_USER

msg["To"] = EMAIL_TO

msg.set_content("""

Bom dia,

Segue em anexo o preços atualizado do Site Spani Spani.

Arquivo gerado automaticamente pelo robô.

Att,
Bruno

""")

# =====================================
# ANEXO
# =====================================

with open("SPANI.xlsx", "rb") as f:

    file_data = f.read()

    msg.add_attachment(
        file_data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="SPANI.xlsx"
    )

# =====================================
# SMTP
# =====================================

with smtplib.SMTP_SSL(
    "smtp.gmail.com",
    465
) as smtp:

    smtp.login(
        EMAIL_USER,
        EMAIL_PASS
    )

    smtp.send_message(msg)

print("EMAIL ENVIADO")
