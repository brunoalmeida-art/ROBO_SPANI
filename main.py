from playwright.sync_api import sync_playwright
import pandas as pd
import smtplib
import os

from email.message import EmailMessage
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
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

    page = browser.new_page()

    # =====================================
    # SITE
    # =====================================

    page.goto(
        "https://www.spanionline.com.br",
        timeout=120000
    )

    page.wait_for_timeout(8000)

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
    # SOMENTE 1 PAGINA TESTE
    # =====================================

    links = links[:20]

    # =====================================
    # ABRIR PRODUTOS
    # =====================================

    for i, link in enumerate(links):

        try:

            print(f"{i+1}/{len(links)}")

            page.goto(
                link,
                timeout=120000
            )

            page.wait_for_timeout(4000)

            # =================================
            # NOME PRODUTO
            # =================================

            nome = ""

            try:

                h1 = page.locator("h1")

                if h1.count() > 0:

                    nome = (
                        h1.first
                        .inner_text()
                        .strip()
                    )

            except Exception as e:

                print("ERRO NOME:", e)

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
                    )

            except Exception as e:

                print("ERRO SETOR:", e)

            # =================================
            # PREÇO
            # =================================

            preco = ""

            try:

                spans = page.locator("span")

                total_spans = spans.count()

                for x in range(total_spans):

                    texto = (
                        spans
                        .nth(x)
                        .inner_text()
                    )

                    if "R$" in texto:

                        preco = texto

                        break

            except Exception as e:

                print("ERRO PRECO:", e)

            # =================================
            # SALVAR
            # =================================

            dados.append({

                "SETOR": setor,

                "PRODUTO": nome,

                "PRECO": preco,

                "LINK": link
            })

            print("OK:", nome)

        except Exception as e:

            print("ERRO:", e)

    browser.close()

# =========================================
# DATAFRAME
# =========================================

df = pd.DataFrame(dados)

df = df.drop_duplicates()

print(df.head())

# =========================================
# EXCEL
# =========================================

df.to_excel(
    ARQUIVO,
    index=False
)

# =========================================
# FORMATAR EXCEL
# =========================================

wb = load_workbook(ARQUIVO)

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

# =========================================
# TAMANHO
# =========================================

larguras = {

    "A": 25,

    "B": 60,

    "C": 15,

    "D": 80
}

for col, largura in larguras.items():

    ws.column_dimensions[col].width = largura

# =========================================
# LINK CLICAVEL
# =========================================

for row in range(2, ws.max_row + 1):

    cell = ws[f"D{row}"]

    cell.hyperlink = cell.value

    cell.font = Font(
        color="0000FF",
        underline="single"
    )

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

Segue em anexo o relatório atualizado do Spani.

TOTAL PRODUTOS: {len(df)}

Att,
Bruno

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

    print("ERRO EMAIL:", e)
