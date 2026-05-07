from playwright.sync_api import sync_playwright
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

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

    page.wait_for_timeout(8000)

    # =========================
    # PEGA CARDS
    # =========================

    cards = page.locator("a[href*='/produto/']")

    total = cards.count()

    print(f"TOTAL PRODUTOS: {total}")

    links_usados = set()

    for i in range(total):

        try:

            href = cards.nth(i).get_attribute("href")

            texto = cards.nth(i).inner_text().strip()

            if href:

                link = "https://www.spanionline.com.br" + href

                if link not in links_usados:

                    links_usados.add(link)

                    print("\n================")
                    print(texto)
                    print(link)

                    linhas.append([
                        "BEBIDAS",
                        texto,
                        "",
                        "",
                        "",
                        "",
                        link
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

df = df.drop_duplicates()

arquivo = "SPANI.xlsx"

with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:

    df.to_excel(
        writer,
        index=False,
        sheet_name="SPANI"
    )

# =========================
# FORMATAÇÃO
# =========================

wb = load_workbook(arquivo)

ws = wb["SPANI"]

cor_azul = PatternFill(
    start_color="1F4E78",
    end_color="1F4E78",
    fill_type="solid"
)

fonte_branca = Font(
    color="FFFFFF",
    bold=True
)

# cabeçalho
for cell in ws[1]:

    cell.fill = cor_azul
    cell.font = fonte_branca

# autofilter
ws.auto_filter.ref = ws.dimensions

# largura automática
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

wb.save(arquivo)

print("\n================")
print("ARQUIVO GERADO")
print(arquivo)

print("TOTAL PRODUTOS:")
print(len(df))
