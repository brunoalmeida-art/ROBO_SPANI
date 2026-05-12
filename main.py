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

OUTPUT = "SPANI.xlsx"

EMAIL_USER = os.getenv("EMAIL_USER")

EMAIL_PASS = os.getenv("EMAIL_PASS")

EMAIL_TO = os.getenv("EMAIL_TO")

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

# =========================================
# COOKIES REQUEST
# =========================================

cookies = {

    "sessao-id": session_id,

    "vip-token": vip_token
}

# =========================================
# LOOP PAGINAS
# =========================================

pagina = 1

todos = []

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

    # =====================================
    # PRODUTOS
    # =====================================

    produtos = data.get(
        "data",
        {}
    ).get(
        "produtos",
        []
    )

    print("PRODUTOS:", len(produtos))

    if len(produtos) == 0:

        print("SEM MAIS PRODUTOS")

        break

    # =====================================
    # LOOP PRODUTOS
    # =====================================

    for p in produtos:

        try:

            # =================================
            # PRODUTO
            # =================================

            produto = (

                p.get(
                    "descricao",
                    ""
                )

                .strip()

                .upper()
            )

            # =================================
            # SETOR
            # =================================

            setor = str(

                p.get(
                    "secao_id",
                    "SEM SETOR"
                )

            ).upper()

            # =================================
            # PRECO VAREJO
            # =================================

            varejo = p.get(
                "preco",
                ""
            )

            # =================================
            # PRECO ATACADO
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

                    if (

                        preco_oferta

                        and

                        float(preco_oferta)
                        <
                        float(varejo)

                    ):

                        atacado = preco_oferta

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

                "LINK": link
            })

        except Exception as e:

            print(
                "ERRO PRODUTO:",
                e
            )

    pagina += 1

# =========================================
# VALIDAR
# =========================================

if len(todos) == 0:

    raise Exception(
        "SEM DADOS API"
    )

# =========================================
# DATAFRAME
# =========================================

df = pd.DataFrame(todos)

df = df.sort_values(
    by="PRODUTO"
)

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
# LINKS
# =========================================

for row in range(2, ws.max_row + 1):

    cell = ws[f"F{row}"]

    url = cell.value

    cell.value = "ABRIR"

    cell.hyperlink = url

    cell.style = "Hyperlink"

# =========================================
# LARGURA
# =========================================

larguras = {

    1: 18,

    2: 70,

    3: 12,

    4: 12,

    5: 15,

    6: 12
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

    ref=f"A1:F{ws.max_row}"
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

    msg["Subject"] = "SPANI - ROBÔ PREÇOS"

    msg["From"] = EMAIL_USER

    msg["To"] = EMAIL_TO

    msg.set_content(

        "Arquivo SPANI em anexo."
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
