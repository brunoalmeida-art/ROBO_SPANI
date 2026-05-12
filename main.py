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
# CONFIG
# =========================

CEP = "05302-040"

HOJE = datetime.now().strftime("%d-%m-%Y_%H-%M")

ARQUIVO_FINAL = f"SPANI_FULL_{HOJE}.xlsx"

# =========================
# EMAIL
# =========================

EMAIL_USER = os.getenv("EMAIL_USER")

EMAIL_PASS = os.getenv("EMAIL_PASS")

DESTINATARIO = "pricing@roldao.com.br"

# =========================
# SESSION
# =========================

session = requests.Session()

# =========================
# TOKEN
# =========================

TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJ2aXBjb21tZXJjZSIsImF1ZCI6ImFwaS1hZG1pbiIsInN1YiI6IjZiYzQ4NjdlLWRjYTktMTFlOS04NzQyLTAyMGQ3OTM1OWNhMCIsInZpcGNvbW1lcmNlQ2xpZW50ZUlkIjpudWxsLCJpYXQiOjE3Nzc5MDc4MTMsInZlciI6MSwiY2xpZW50IjpudWxsLCJvcGVyYXRvciI6bnVsbCwib3JnIjoiNjcifQ.mqyEyNRMBcY0rb4kWeNN0-xnEb8kus9i97w3IR6qjCCPdKEyBjUcZkF77_4KtKvHBI2cx25Fd8E9G4Q1cwsADw"

# =========================
# HEADERS
# =========================

HEADERS = {

    "accept": "application/json",

    "content-type": "application/json",

    "user-agent": "Mozilla/5.0",

    "Authorization": f"Bearer {TOKEN}",

    "OrganizationId": "67",

    "DomainKey": "spanionline.com.br"
}

# =========================
# DEFINIR CEP
# =========================

try:

    url_cep = (
        "https://services-beta.vipcommerce.com.br/"
        "api/v1/checkout/cliente/cep"
    )

    payload = {
        "cep": CEP
    }

    r_cep = session.post(
        url_cep,
        json=payload,
        headers=HEADERS,
        timeout=60
    )

    print("STATUS CEP:", r_cep.status_code)

except Exception as e:

    print("ERRO CEP:", e)

# =========================
# ÁRVORE DEPARTAMENTOS
# =========================

MAPA_SETORES = {}

try:

    url_departamentos = (
        "https://services-beta.vipcommerce.com.br/"
        "api-admin/v1/org/67/"
        "filial/1/"
        "centro_distribuicao/36/"
        "loja/classificacoes_mercadologicas/"
        "departamentos/arvore"
    )

    r_dep = session.get(
        url_departamentos,
        headers=HEADERS,
        timeout=60
    )

    js_dep = r_dep.json()

    print(
        json.dumps(
            js_dep,
            indent=2,
            ensure_ascii=False
        )[:5000]
    )

    departamentos = js_dep.get(
        "data",
        []
    )

    # =====================
    # LOOP DEPARTAMENTOS
    # =====================

    for dep in departamentos:

        dep_id = dep.get("id")

        dep_nome = (
            dep.get("descricao")
            or dep.get("nome")
            or dep.get("titulo")
        )

        if dep_id and dep_nome:

            MAPA_SETORES[dep_id] = dep_nome

        # =================
        # SUBCATEGORIAS
        # =================

        filhos = dep.get("filhos", [])

        for filho in filhos:

            filho_id = filho.get("id")

            filho_nome = (
                filho.get("descricao")
                or filho.get("nome")
                or filho.get("titulo")
            )

            if filho_id and filho_nome:

                MAPA_SETORES[
                    filho_id
                ] = filho_nome

    print("MAPA SETORES:")

    print(MAPA_SETORES)

except Exception as e:

    print("ERRO DEPARTAMENTOS:", e)

# =========================
# URL BUSCA
# =========================

url = (
    "https://services-beta.vipcommerce.com.br/"
    "api-admin/v1/org/67/"
    "filial/1/"
    "centro_distribuicao/36/"
    "loja/buscas/produtos/termo/a"
    "?page=1"
)

# =========================
# REQUEST
# =========================

try:

    r = session.get(
        url,
        headers=HEADERS,
        timeout=120
    )

    print("STATUS:", r.status_code)

except Exception as e:

    print("ERRO REQUEST:", e)

    raise

# =========================
# JSON
# =========================

try:

    js = r.json()

except Exception as e:

    print("ERRO JSON:", e)

    raise

# =========================
# DEBUG
# =========================

print(json.dumps(js, indent=2, ensure_ascii=False)[:10000])

# =========================
# PRODUTOS
# =========================

produtos = js.get("data", {}).get("produtos", [])

print("TOTAL PRODUTOS API:", len(produtos))

# =========================
# PROCESSAMENTO
# =========================

dados = []

for i, p in enumerate(produtos):

    try:

        print(f"PROCESSANDO {i+1}/{len(produtos)}")

        produto_id = p.get("produto_id")

        nome = p.get("descricao")

        varejo = p.get("preco")

        varejo = float(varejo) if varejo else None

        ean = p.get("codigo_barras")

        # =====================
        # SETOR
        # =====================

        codigo_setor = p.get(
            "classificacao_mercadologica_id"
        )

        setor = MAPA_SETORES.get(
            codigo_setor,
            f"SETOR {codigo_setor}"
        )

        # =====================
        # LINK CORRETO
        # =====================

        link_produto = p.get("link")

        url_produto = ""

        if produto_id and link_produto:

            url_produto = (
                "https://www.spanionline.com.br/produto/"
                f"{produto_id}/"
                f"{link_produto}"
            )

        # =====================
        # VALIDAR PRODUTO
        # =====================

        produto_valido = False

        if url_produto:

            try:

                validar = session.get(
                    url_produto,
                    headers={
                        "User-Agent": "Mozilla/5.0"
                    },
                    timeout=30
                )

                html = validar.text.lower()

                if (
                    "produto indisponível" not in html
                    and
                    "produto indisponivel" not in html
                ):

                    produto_valido = True

            except:

                pass

        # =====================
        # IGNORAR INVÁLIDOS
        # =====================

        if not produto_valido:

            print(
                "❌ PRODUTO INDISPONÍVEL:",
                nome
            )

            continue

        # =====================
        # OFERTA
        # =====================

        oferta = p.get("oferta")

        atacado = None

        qtd_atacado = None

        if oferta:

            preco_oferta = oferta.get(
                "preco_oferta"
            )

            if preco_oferta:

                atacado = float(
                    preco_oferta
                )

            quantidade_minima = oferta.get(
                "quantidade_minima"
            )

            if quantidade_minima:

                qtd_atacado = int(
                    quantidade_minima
                )

        # =====================
        # SALVAR
        # =====================

        dados.append({

            "SETOR": setor,

            "PRODUTO": nome,

            "VAREJO": varejo,

            "ATACADO": atacado,

            "QTD ATACADO": qtd_atacado,

            "EAN": ean,

            "URL": url_produto
        })

        print("✅ OK:", nome)

    except Exception as e:

        print("ERRO PRODUTO:", e)

# =========================
# DATAFRAME
# =========================

df = pd.DataFrame(dados)

if df.empty:

    print("DATAFRAME VAZIO")

    raise Exception("SEM DADOS")

# =========================
# REMOVER DUPLICADOS
# =========================

if "EAN" in df.columns:

    df = df.drop_duplicates(
        subset=["EAN"]
    )

print("TOTAL PRODUTOS VÁLIDOS:", len(df))

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

wb = load_workbook(
    ARQUIVO_FINAL
)

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

ws.insert_cols(7)

ws["G1"] = "LINK"

for row in range(2, ws.max_row + 1):

    try:

        url_link = ws[f"H{row}"].value

        if url_link:

            cell = ws[f"G{row}"]

            cell.value = "ABRIR"

            cell.hyperlink = url_link

            cell.font = Font(
                color="0000FF",
                underline="single"
            )

        ws[f"C{row}"].number_format = '0.00'

        ws[f"D{row}"].number_format = '0.00'

        ws[f"E{row}"].number_format = '0'

        ws[f"F{row}"].number_format = '@'

    except Exception as e:

        print("ERRO FORMATACAO:", e)

# =========================
# LARGURA
# =========================

colunas = {

    "A": 25,

    "B": 60,

    "C": 12,

    "D": 12,

    "E": 14,

    "F": 20,

    "G": 10,

    "H": 70
}

for col, largura in colunas.items():

    ws.column_dimensions[col].width = largura

ws.column_dimensions["H"].hidden = True

ws.freeze_panes = "A2"

wb.save(
    ARQUIVO_FINAL
)

print("🔥 FINALIZADO:", ARQUIVO_FINAL)

# =========================
# EMAIL
# =========================

def enviar_email(arquivo):

    if not EMAIL_USER or not EMAIL_PASS:

        print("❌ EMAIL OU SENHA NÃO CONFIGURADOS")

        return

    msg = EmailMessage()

    msg["Subject"] = (
        f"Relatório Spani {HOJE}"
    )

    msg["From"] = EMAIL_USER

    msg["To"] = DESTINATARIO

    msg.set_content(f"""

Bom dia,

Segue em anexo o relatório atualizado de preços coletados no site do Spani Atacadista.

A coleta foi realizada utilizando o CEP {CEP}.

Arquivo gerado automaticamente pelo robô de monitoramento de preços.

TOTAL PRODUTOS VÁLIDOS: {len(df)}

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

            print(
                "📧 EMAIL ENVIADO COM SUCESSO!"
            )

    except Exception as e:

        print(
            "❌ ERRO AO ENVIAR EMAIL:",
            e
        )

# =========================
# EXECUTAR EMAIL
# =========================

try:

    enviar_email(
        ARQUIVO_FINAL
    )

except Exception as e:

    print("ERRO FINAL:", e)
