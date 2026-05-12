from playwright.sync_api import sync_playwright
import pandas as pd
import smtplib
import os
import re

from email.message import EmailMessage
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime

# =========================================
# DATA
# =========================================

HOJE = datetime.now().strftime("%d-%m-%Y")

ARQUIVO = "SPANI.xlsx"

# =========================================
# EMAIL
# =========================================

EMAIL_USER = os.getenv("EMAIL_USER")

EMAIL_PASS = os.getenv("EMAIL_PASS")

DESTINATARIO = "pricing@roldao.com.br"

# =========================================
# DADOS
# =========================================

dados = []

# =========================================
# PLAYWRIGHT
# =========================================

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=True
    )

    context = browser.new_context()

    page = context.new_page()

    # =====================================
    # SITE
    # =====================================

    page.goto(
        "https://www.spanionline.com.br",
        timeout=120000
    )

    page.wait_for_timeout(8000)

    # =====================================
    # DEFINIR LOJA
    # =====================================

    try:

        print("DEFININDO LOJA MAUA")

        page.locator(
            "text=Retirar no endereço"
        ).click()

        page.wait_for_timeout(3000)

        page.locator(
            "input"
        ).nth(1).fill("Mauá")

        page.wait_for_timeout(3000)

        page.locator(
            "text=Spani Mauá 1"
        ).click()

        page.wait_for_timeout(5000)

        print("LOJA DEFINIDA")

    except Exception as e:

        print(
            "ERRO LOJA:",
            e
        )

    # =====================================
    # BUSCA
    # =====================================

    busca = page.locator("input")

    busca.first.fill("a")

    page.keyboard.press("Enter")

    page.wait_for_timeout(8000)

    # =====================================
    # SCROLL
    # =====================================

    for i in range(10):

        page.mouse.wheel(0, 15000)

        print(f"SCROLL {i+1}")

        page.wait_for_timeout(2500)

    # =====================================
    # PEGAR LINKS
    # =====================================

    produtos = page.locator("a")

    total = produtos.count()

    print("TOTAL ELEMENTOS:", total)

    links = []

    for i in range(total):

        try:

            href = produtos.nth(i).get_attribute(
                "href"
            )

            if href and "/produto/" in href:

                if href.startswith("http"):

                    link = href

                else:

                    link = (
                        "https://www.spanionline.com.br"
                        + href
                    )

                if link not in links:

                    links.append(link)

        except:

            pass

    print("TOTAL LINKS:", len(links))

    # =====================================
    # TESTE
    # =====================================

    links = links[:20]

    # =====================================
    # PRODUTOS
    # =====================================

    for i, link in enumerate(links):

        try:

            print(f"{i+1}/{len(links)}")

            page.goto(
                link,
                timeout=120000
            )

            page.wait_for_timeout(5000)

            texto = (
                page.locator("body")
                .inner_text()
            )

            # =================================
            # PRODUTO
            # =================================

            produto = ""

            try:

                h1 = page.locator("h1")

                if h1.count() > 0:

                    produto = (
                        h1.first
                        .inner_text()
                        .strip()
                    )

            except Exception as e:

                print(
                    "ERRO PRODUTO:",
                    e
                )

            # =================================
            # SETOR
            # =================================

            setor = ""

            try:

                breadcrumb = page.locator(
                    ".vip-breadcrumb-label"
                )

                total_breadcrumb = (
                    breadcrumb.count()
                )

                if total_breadcrumb >= 2:

                    setor = (
                        breadcrumb
                        .nth(
                            total_breadcrumb - 2
                        )
                        .inner_text()
                        .strip()
                        .upper()
                    )

            except Exception as e:

                print(
                    "ERRO SETOR:",
                    e
                )

            # =================================
            # PREÇO VAREJO
            # =================================

            varejo = ""

            try:

                varejo_elemento = page.locator(
                    "text=/R\\$\\s?\\d+,\\d+/"
                ).first

                varejo = (
                    varejo_elemento
                    .inner_text()
                    .strip()
                    .replace("/un", "")
                    .strip()
                )

            except Exception as e:

                print(
                    "ERRO VAREJO:",
                    e
                )

            # =================================
            # PREÇO ATACADO
            # =================================

            atacado = ""

            try:

                valores = page.locator(
                    "text=/R\\$\\s?\\d+,\\d+/"
                )

                if valores.count() >= 2:

                    atacado = (
                        valores
                        .nth(1)
                        .inner_text()
                        .strip()
                        .replace("/un", "")
                        .strip()
                    )

            except Exception as e:

                print(
                    "ERRO ATACADO:",
                    e
                )

            # =================================
            # QTD ATACADO
            # =================================

            qtd_atacado = ""

            try:

                qtd_match = re.search(

                    r'a partir da\s*(\d+)',

                    texto,

                    re.IGNORECASE
                )

                if qtd_match:

                    qtd_atacado = (
                        qtd_match.group(1)
                    )

            except Exception as e:

                print(
                    "ERRO QTD:",
                    e
                )

            # =================================
            # SALVAR
            # =================================

            dados.append({

                "SETOR": setor,

                "PRODUTO": produto,

                "VAREJO": varejo,

                "ATACADO": atacado,

                "QTD ATACADO": qtd_atacado,

                "LINK": link
            })

            print(
                "OK:",
                produto
            )

        except Exception as e:

            print(
                "ERRO:",
                e
            )

    browser.close()

# =========================================
# DATAFRAME
# =========================================

df = pd.DataFrame(dados)

df = df.drop_duplicates()

# =========================================
# ORDENAR
# =========================================

df = df.sort_values(
    by=["SETOR", "PRODUTO"]
)

print(df.head())

# =========================================
# EXCEL
# =========================================

df.to_excel(
    ARQUIVO,
    index=False
)

# =========================================
# FORMATAR
# =========================================

wb = load_workbook(
    ARQUIVO
)

ws = wb.active

# =========================================
# HEADER
# =========================================

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

    cell.alignment = Alignment(
        horizontal="center"
    )

# =========================================
# LINKS
# =========================================

for row in range(2, ws.max_row + 1):

    url = ws[f"F{row}"].value

    ws[f"F{row}"] = "ABRIR"

    ws[f"F{row}"].hyperlink = url

    ws[f"F{row}"].font = Font(
        color="0000FF",
        underline="single"
    )

# =========================================
# LARGURA
# =========================================

larguras = {

    "A": 35,

    "B": 70,

    "C": 12,

    "D": 12,

    "E": 15,

    "F": 12
}

for col, largura in larguras.items():

    ws.column_dimensions[col].width = largura

# =========================================
# CONGELAR
# =========================================

ws.freeze_panes = "A2"

wb.save(ARQUIVO)

print("FINALIZADO")

# =========================================
# EMAIL
# =========================================

try:

    if EMAIL_USER and EMAIL_PASS:

        msg = EmailMessage()

        msg["Subject"] = (
            f"Relatório Spani {HOJE}"
        )

        msg["From"] = EMAIL_USER

        msg["To"] = DESTINATARIO

        msg.set_content(f"""

Bom dia,

Segue em anexo o relatório atualizado do Spani Mauá 1.

TOTAL PRODUTOS: {len(df)}

Att,
Bruno

Competitividade – Spani

""")

        with open(ARQUIVO, "rb") as f:

            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="octet-stream",
                filename=ARQUIVO
            )

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

except Exception as e:

    print(
        "ERRO EMAIL:",
        e
    )
