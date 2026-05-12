import requests
import pandas as pd
import smtplib
import os
import json

from email.message import EmailMessage
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime

# =========================
# DATA
# =========================

HOJE = datetime.now().strftime("%d-%m-%Y_%H-%M")

ARQUIVO_FINAL = f"SPANI_FULL_{HOJE}.xlsx"

# =========================
# EMAIL
# =========================

EMAIL_USER = os.getenv("EMAIL_USER")

EMAIL_PASS = os.getenv("EMAIL_PASS")

DESTINATARIO = "pricing@roldao.com.br"

# =========================
# URL
# =========================

url = (
    "https://services-beta.vipcommerce.com.br/"
    "api-admin/v1/org/67/"
    "filial/1/"
    "centro_distribuicao/36/"
    "loja/buscas/produtos/termo/a"
    "?page=1"
    "&&session=dc9e71d5-6b54-4c28-9f5c-0a0ab5dc316e"
)

# =========================
# HEADERS
# =========================

headers = {

    "accept": "application/json",

    "user-agent": "Mozilla/5.0",

    "OrganizationId": "67",

    "DomainKey": "spanionline.com.br",

    "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJ2aXBjb21tZXJjZSIsImF1ZCI6ImFwaS1hZG1pbiIsInN1YiI6IjZiYzQ4NjdlLWRjYTktMTFlOS04NzQyLTAyMGQ3OTM1OWNhMCIsInZpcGNvbW1lcmNlQ2xpZW50ZUlkIjpudWxsLCJpYXQiOjE3Nzc5MDc4MTMsInZlciI6MSwiY2xpZW50IjpudWxsLCJvcGVyYXRvciI6bnVsbCwib3JnIjoiNjcifQ.mqyEyNRMBcY0rb4kWeNN0-xnEb8kus9i97w3IR6qjCCPdKEyBjUcZkF77_4KtKvHBI2cx25Fd8E9G4Q1cwsADw"
}

# =========================
# REQUEST
# =========================

r = requests.get(
    url,
    headers=headers,
    timeout=120
)

print("STATUS:", r.status_code)

js = r.json()

print(json.dumps(js, indent=2, ensure_ascii=False)[:10000])

# =========================
# PRODUTOS
# =========================

produtos = js.get("data", {}).get("produtos", [])

print("TOTAL PRODUTOS:", len(produtos))

# =========================
# PROCESSAMENTO
# =========================

dados = []

for p in produtos:

    nome = p.get("descricao")

    varejo = p.get("preco")

    varejo = float(varejo) if varejo else None

    ean = p.get("codigo_barras")

    setor = p.get("classificacao_mercadologica_id")

    sku = p.get("sku")

    link_produto = p.get("link")

    url_produto = ""

    if link_produto:

        url_produto = f"https://www.spanionline.com.br/produto/{link_produto}"

    # =====================
    # OFERTA
    # =====================

    oferta = p.get("oferta", {})

    atacado = oferta.get("preco_oferta")

    atacado = float(atacado) if atacado else None

    qtd_atacado = oferta.get("quantidade_minima")

    qtd_atacado = int(qtd_atacado) if qtd_atacado else None

    dados.append({

        "SETOR": setor,

        "PRODUTO": nome,

        "VAREJO": varejo,

        "ATACADO": atacado,

        "QTD ATACADO": qtd_atacado,

        "EAN": ean,

        "SKU": sku,

        "URL": url_produto
    })

# =========================
# DATAFRAME
# =========================

df = pd.DataFrame(dados)

df = df.drop_duplicates(subset=["EAN"])

# =========================
# EXCEL
# =========================

df.to_excel(
    ARQUIVO_FINAL,
    index=False
)

# =========================
# FORMATAR EXCEL
# =========================

wb = load_workbook(ARQUIVO_FINAL)

ws = wb.active

fill = PatternFill(
    start_color="1F4E78",
    end_color="1F4E78",
    fill_type="solid"
)

font_header = Font(
    color="FFFFFF",
    bold=True
)

for cell in ws[1]:

    cell.fill = fill

    cell.font = font_header

# =========================
# LINK
# =========================

ws.insert_cols(8)

ws["H1"] = "LINK"

for row in range(2, ws.max_row + 1):

    url = ws[f"I{row}"].value

    if url:

        cell = ws[f"H{row}"]

        cell.value = "ABRIR"

        cell.hyperlink = url

        cell.font = Font(
            color="0000FF",
            underline="single"
        )

    ws[f"C{row}"].number_format = '0.00'

    ws[f"D{row}"].number_format = '0.00'

    ws[f"E{row}"].number_format = '0'

    ws[f"F{row}"].number_format = '@'

# =========================
# LARGURA
# =========================

colunas = {

    "A": 18,

    "B": 60,

    "C": 12,

    "D": 12,

    "E": 14,

    "F": 20,

    "G": 15,

    "H": 10,

    "I": 70
}

for col, largura in colunas.items():

    ws.column_dimensions[col].width = largura

ws.column_dimensions["I"].hidden = True

ws.freeze_panes = "A2"

wb.save(ARQUIVO_FINAL)

print("🔥 FINALIZADO:", ARQUIVO_FINAL)

# =========================
# EMAIL
# =========================

def enviar_email(arquivo):

    if not EMAIL_USER or not EMAIL_PASS:

        print("❌ EMAIL OU SENHA NÃO CONFIGURADOS")

        return

    msg = EmailMessage()

    msg["Subject"] = f"Relatório Spani {HOJE}"

    msg["From"] = EMAIL_USER

    msg["To"] = DESTINATARIO

    msg.set_content(f"""

Bom dia,

Segue em anexo o relatório atualizado de preços coletados no site do Spani Atacadista.

Arquivo gerado automaticamente pelo robô de monitoramento de preços.

TOTAL PRODUTOS: {len(df)}

Att,
Bruno

""")

    with open(arquivo, "rb") as f:

        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="octet-stream",
            filename=arquivo
        )

    try:

        with smtplib.SMTP_SSL(
            "smtp.gmail.com",
            465
        ) as smtp:

            smtp.login(
                EMAIL_USER,
                EMAIL_PASS
            )

            smtp.send_message(msg)

            print("📧 EMAIL ENVIADO COM SUCESSO!")

    except Exception as e:

        print("❌ ERRO AO ENVIAR EMAIL:", e)

# =========================
# EXECUTAR EMAIL
# =========================

enviar_email(ARQUIVO_FINAL)
