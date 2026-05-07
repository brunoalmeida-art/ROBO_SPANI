from playwright.sync_api import sync_playwright
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

linhas = []

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=False,
        args=[
            "--disable-blink-features=AutomationControlled"
        ]
    )

    context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36",
        viewport={"width": 1920, "height": 1080}
    )

    page = context.new_page()

    url = "https://www.spanionline.com.br/produto/832/leite-longa-vida-italac-1l-semidesnatado"

    page.goto(url, timeout=120000)

    print("PAGINA ABERTA")

    page.wait_for_timeout(10000)

    # =========================
    # NOME
    # =========================

    nome = page.locator("h1").first.inner_text()

    print("\nPRODUTO:")
    print(nome)

    # =========================
    # PREÇO
    # =========================

    preco = page.locator("text=R$").first.inner_text()

    print("\nPRECO:")
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

print("\n===================")
print("ARQUIVO GERADO")
print(arquivo)

print("TOTAL:")
print(len(df))
