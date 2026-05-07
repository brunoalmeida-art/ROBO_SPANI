from playwright.sync_api import sync_playwright
import pandas as pd

linhas = []

with sync_playwright() as p:

    browser = p.chromium.launch(headless=True)

    page = browser.new_page()

    busca = "skol"

    page.goto(
        f"https://www.spanionline.com.br/busca/{busca}",
        timeout=120000
    )

    print("PAGINA ABERTA")

    page.wait_for_timeout(5000)

    produtos = page.locator("a")

    total = produtos.count()

    print(f"TOTAL LINKS: {total}")

    links_usados = set()

    for i in range(total):

        try:

            texto = produtos.nth(i).inner_text().strip()

            href = produtos.nth(i).get_attribute("href")

            if href and "/produto/" in href:

                link_completo = "https://www.spanionline.com.br" + href

                if link_completo not in links_usados:

                    links_usados.add(link_completo)

                    print("\n===================")
                    print(texto)
                    print(link_completo)

                    linhas.append([
                        "BEBIDAS",
                        texto,
                        "",
                        "",
                        "",
                        "",
                        link_completo
                    ])

        except:
            pass

    browser.close()

# =========================
# DATAFRAME
# =========================

colunas = [
    "SETOR",
    "PRODUTO",
    "PREÇO VAREJO",
    "PREÇO ANTIGO",
    "PREÇO ATACADO",
    "QTD ATACADO",
    "LINK"
]

df = pd.DataFrame(linhas, columns=colunas)

# remove duplicados
df = df.drop_duplicates()

# =========================
# EXPORTA EXCEL
# =========================

arquivo = "SPANI.xlsx"

with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:

    df.to_excel(
        writer,
        index=False,
        sheet_name="SPANI"
    )

print("\n===================")
print("ARQUIVO GERADO")
print(arquivo)
print("TOTAL:", len(df))
