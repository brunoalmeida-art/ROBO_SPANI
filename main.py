from playwright.sync_api import sync_playwright
import pandas as pd
import time

dados = []

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=True
    )

    page = browser.new_page()

    # =========================
    # SITE
    # =========================

    page.goto(
        "https://www.spanionline.com.br",
        timeout=120000
    )

    page.wait_for_timeout(8000)

    # =========================
    # BUSCA
    # =========================

    busca = page.locator('input')

    busca.first.fill("a")

    page.keyboard.press("Enter")

    page.wait_for_timeout(8000)

    # =========================
    # SCROLL
    # =========================

    for i in range(30):

        page.mouse.wheel(0, 15000)

        print(f"SCROLL {i+1}")

        page.wait_for_timeout(3000)

    # =========================
    # PRODUTOS
    # =========================

    produtos = page.locator("a")

    total = produtos.count()

    print("TOTAL ELEMENTOS:", total)

    links = []

    for i in range(total):

        try:

            href = produtos.nth(i).get_attribute("href")

            if href and "/produto/" in href:

                link = (
                    "https://www.spanionline.com.br"
                    + href
                )

                if link not in links:

                    links.append(link)

        except:

            pass

    print("TOTAL LINKS:", len(links))

    # =========================
    # ABRIR PRODUTOS
    # =========================

    for i, link in enumerate(links):

        try:

            print(f"{i+1}/{len(links)}")

            page.goto(
                link,
                timeout=120000
            )

            page.wait_for_timeout(4000)

            html = page.content()

            # =====================
            # NOME
            # =====================

            nome = ""

            try:

                nome = (
                    page.locator("h1")
                    .first
                    .inner_text()
                    .strip()
                )

            except:

                pass

            # =====================
            # PREÇO
            # =====================

            preco = ""

            try:

                spans = page.locator("span")

                total_spans = spans.count()

                for x in range(total_spans):

                    texto = spans.nth(x).inner_text()

                    if "R$" in texto:

                        preco = texto

                        break

            except:

                pass

            # =====================
            # SETOR
            # =====================

            setor = ""

            try:

                breadcrumb = page.locator(
                    ".vip-breadcrumb-label.last"
                )

                setor = (
                    breadcrumb.first
                    .inner_text()
                    .strip()
                )

            except:

                pass

            dados.append({

                "SETOR": setor,

                "PRODUTO": nome,

                "PRECO": preco,

                "LINK": link
            })

            print(nome)

        except Exception as e:

            print("ERRO:", e)

    browser.close()

# =========================
# DATAFRAME
# =========================

df = pd.DataFrame(dados)

df = df.drop_duplicates()

print(df.head())

# =========================
# EXCEL
# =========================

df.to_excel(
    "SPANI.xlsx",
    index=False
)

print("FINALIZADO")
