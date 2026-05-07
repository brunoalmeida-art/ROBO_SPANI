from playwright.sync_api import sync_playwright

with sync_playwright() as p:

    browser = p.chromium.launch(headless=True)

    page = browser.new_page()

    url = "https://www.spanionline.com.br/produto/832/leite-longa-vida-italac-1l-semidesnatado"

    page.goto(url, timeout=120000)

    print("PAGINA ABERTA")

    page.wait_for_timeout(10000)

    print("\nTITULO:")
    print(page.title())

    print("\nHTML:")
    print(page.content())

    browser.close()
