from playwright.sync_api import sync_playwright

print("ABRINDO SPANI...")

with sync_playwright() as p:

    browser = p.chromium.launch(headless=True)

    page = browser.new_page()

    page.goto(
        "https://www.spanionline.com.br/produto/832/leite-longa-vida-italac-1l-semidesnatado",
        timeout=120000
    )

    print("PAGINA ABERTA")

    titulo = page.title()

    print("TITULO:")
    print(titulo)

    preco = page.locator("text=R$").first.inner_text()

    print("PRECO:")
    print(preco)

    browser.close()
