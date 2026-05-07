from playwright.sync_api import sync_playwright
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

linhas = []

with sync_playwright() as p:

    browser = p.chromium.launch(headless=True)

    page = browser.new_page()

    url = "https://www.spanionline.com.br/produto/832/leite-longa-vida-italac-1l-semidesnatado"

    page.goto(url, timeout=120000)

    print("PAGINA ABERTA")

    page.wait_for_timeout(5000)

    # =========================
    # PRODUTO
    # =========================

    nome = page.locator("h1").first.inner_text()

    print(nome)

    # =========================
    # PREÇO
    # =========================

    preco = page.locator("text=R$").first.inner_text()

    print(preco)

    linhas.append([
        "LATICINIOS",
        nome,
        preco,
        "",
        "",
        "",
        url
    ])

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

for cell in ws[1]:

    cell.fill = cor_azul
    cell.font = fonte_branca

ws.auto_filter.ref = ws.dimensions

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

print("ARQUIVO GERADO")
