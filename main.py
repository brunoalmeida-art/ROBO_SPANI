import requests
import pandas as pd
from playwright.sync_api import sync_playwright
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import smtplib
from email.message import EmailMessage
import os

# =========================================
# CONFIG
# =========================================

BASE_API = "https://services-beta.vipcommerce.com.br"

LOJA_URL = "https://www.spanionline.com.br"

BUSCA = "a"

LIMITE_ITENS = 10

OUTPUT = "SPANI_TESTE.xlsx"

EMAIL_USER = os.getenv("EMAIL_USER")

EMAIL_PASS = os.getenv("EMAIL_PASS")

EMAIL_TO = os.getenv("EMAIL_TO")

# =========================================
# FUNCAO PRECO
# =========================================

def formatar_preco(valor):

    if valor is None or valor == "":

        return ""

    try:

        return (
            f"{float(str(valor).replace(',', '.')):.2f}"
            .replace(".", ",")
        )

    except:

        return str(valor)

# =========================================
# PLAYWRIGHT
# =========================================

print("ABRINDO SPANI...")

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=True
    )

    context = browser.new_context()

    page = context.new_page()

    page.goto(
        LOJA_URL,
        timeout=120000
    )

    page.wait_for_timeout(10000)

    # =====================================
    # SELECIONAR LOJA
    # =====================================

    try:

        print("SELECIONANDO MAUA 1")

        page.locator(
            ".vip-endereco-wrapper"
        ).click()

        page.wait_for_timeout(3000)

        page.locator(
            "text=Spani Mauá 1"
        ).click()

        page.wait_for_timeout(10000)

    except Exception as e:

        print(
            "ERRO LOJA:",
            e
        )

    # =====================================
    # ENDERECO
    # =====================================

    try:

        endereco = page.locator(
            ".vip-endereco-wrapper"
        ).inner_text()

        print("LOJA:", endereco)

    except:

        print("NAO ACHOU ENDERECO")

    # =====================================
    # COOKIES
    # =====================================

    cookies_play = context.cookies()

    session_id = ""

    vip_token = ""

    for c in cookies_play:

        print(c["name"])

        if c["name"] == "sessao-id":

            session_id = c["value"]

        if c["name"] == "vip-token":

            vip_token = c["value"]

    browser.close()

# =========================================
# VALIDAR
# =========================================

print("SESSION:", session_id)

print("VIP TOKEN:", vip_token)

if session_id == "":

    raise Exception(
        "SESSAO VAZIA"
    )

if vip_token == "":

    raise Exception(
        "VIP TOKEN VAZIO"
    )

# =========================================
# HEADERS
# =========================================

headers = {

    "accept": "application/json",

    "origin": "https://www.spanionline.com.br",

    "referer": "https://www.spanionline.com.br/",

    "Authorization": f"Bearer {vip_token}",

    "organizationid": "67",

    "filialid": "1",

    "centrodistribuicaoid": "36",

    "user-agent": "Mozilla/5.0"
}

cookies = {

    "sessao-id": session_id,

    "vip-token": vip_token
}

# =========================================
# BUSCA PRODUTOS
# =========================================

pagina = 1

todos = []

total_coletados = 0

while True:

    url = (

        f"{BASE_API}"

        f"/api-admin/v1/org/67"

        f"/filial/1"

        f"/centro_distribuicao/36"

        f"/loja/buscas/produtos/termo/{BUSCA}"

        f"?page={pagina}"

        f"&session={session_id}"
    )

    print(f"PAGINA {pagina}")

    r = requests.get(

        url,

        headers=headers,

        cookies=cookies,

        timeout=120
    )

    print("STATUS:", r.status_code)

    if r.status_code != 200:

        print(r.text)

        break

    data = r.json()

    produtos = data.get(
        "data",
        {}
    ).get(
        "produtos",
        []
    )

    print("PRODUTOS:", len(produtos))

    if len(produtos) == 0:

        break

    for p in produtos:

        if total_coletados >= LIMITE_ITENS:

            break

        try:

            # =================================
            # SETOR
            # =================================

            setor = (

                p.get("secao_descricao", "")

                or

                p.get("secao_nome", "")

                or

                p.get("secao", "")

            )

            setor = str(setor).strip().upper()

            # =================================
            # PRODUTO
            # =================================

            produto = (

                p.get("descricao", "")

                .strip()

                .upper()
            )

            # =================================
            # EAN
            # =================================

            ean = p.get(
                "codigo_barras",
                ""
            )

            # =================================
            # VAREJO
            # =================================

            varejo_valor = p.get(
                "preco",
                ""
            )

            varejo = formatar_preco(
                varejo_valor
            )

            # =================================
            # ATACADO
            # =================================

            atacado = ""

            qtd_atacado = ""

            oferta = p.get("oferta")

            if oferta:

                preco_oferta = oferta.get(
                    "preco_oferta"
                )

                quantidade_minima = oferta.get(
                    "quantidade_minima"
                )

                try:

                    preco_varejo_float = float(
                        str(varejo_valor).replace(",", ".")
                    )

                    preco_atacado_float = float(
                        str(preco_oferta).replace(",", ".")
                    )

                    if preco_atacado_float < preco_varejo_float:

                        atacado = formatar_preco(
                            preco_oferta
                        )

                        qtd_atacado = quantidade_minima

                except:

                    pass

            # =================================
            # LINK
            # =================================

            slug = p.get(
                "link",
                ""
            )

            produto_id = p.get(
                "produto_id",
                ""
            )

            link = (

                "https://www.spanionline.com.br/produto/"

                f"{produto_id}/"

                f"{slug}"
            )

            # =================================
            # SALVAR
            # =================================

            todos.append({

                "SETOR": setor,

                "PRODUTO": produto,

                "VAREJO": varejo,

                "ATACADO": atacado,

                "QTD ATACADO": qtd_atacado,

                "EAN": ean,

                "LINK": link
            })

            total_coletados += 1

        except Exception as e:

            print(
                "ERRO PRODUTO:",
                e
            )

    if total_coletados >= LIMITE_ITENS:

        print(
            f"LIMITE DE {LIMITE_ITENS} ITENS ATINGIDO"
        )

        break

    pagina += 1

# =========================================
# VALIDAR
# =========================================

if len(todos) == 0:

    raise Exception(
        "SEM DADOS"
    )

# =========================================
# DATAFRAME
# =========================================

df = pd.DataFrame(todos)

df = df.sort_values(
    by="PRODUTO"
).reset_index(drop=True)

print(df.head())

# =========================================
# EXCEL
# =========================================

df.to_excel(
    OUTPUT,
    index=False
)

wb = load_workbook(
    OUTPUT
)

ws = wb.active

# =========================================
# HEADER
# =========================================

fill = PatternFill(

    start_color="16365C",

    end_color="16365C",

    fill_type="solid"
)

font = Font(

    color="FFFFFF",

    bold=True
)

for cell in ws[1]:

    cell.fill = fill

    cell.font = font

# =========================================
# LINK
# =========================================

for row in range(2, ws.max_row + 1):

    cell = ws[f"G{row}"]

    url = cell.value

    cell.value = "ABRIR"

    cell.hyperlink = url

    cell.style = "Hyperlink"

# =========================================
# LARGURA
# =========================================

larguras = {

    1: 20,
    2: 60,
    3: 12,
    4: 12,
    5: 15,
    6: 18,
    7: 12
}

for col, largura in larguras.items():

    ws.column_dimensions[
        get_column_letter(col)
    ].width = largura

# =========================================
# TABELA
# =========================================

tab = Table(

    displayName="TabelaSpani",

    ref=f"A1:G{ws.max_row}"
)

style = TableStyleInfo(

    name="TableStyleMedium2",

    showRowStripes=False,

    showColumnStripes=False
)

tab.tableStyleInfo = style

ws.add_table(tab)

wb.save(OUTPUT)

print("EXCEL FINALIZADO")

# =========================================
# EMAIL
# =========================================

try:

    print("ENVIANDO EMAIL")

    msg = EmailMessage()

    msg["Subject"] = "SPANI - TESTE 10 ITENS"

    msg["From"] = EMAIL_USER

    msg["To"] = EMAIL_TO

    msg.set_content(

        "Segue arquivo teste com 10 itens."
    )

    with open(
        OUTPUT,
        "rb"
    ) as f:

        msg.add_attachment(

            f.read(),

            maintype="application",

            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",

            filename=OUTPUT
        )

    with smtplib.SMTP(
        "smtp.gmail.com",
        587
    ) as smtp:

        smtp.starttls()

        smtp.login(
            EMAIL_USER,
            EMAIL_PASS
        )

        smtp.send_message(msg)

    print("EMAIL ENVIADO")

except Exception as e:

    print(
        "ERRO EMAIL:",
        e
    )

print("FINALIZADO")
