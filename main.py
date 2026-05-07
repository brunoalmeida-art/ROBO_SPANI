from playwright.sync_api import sync_playwright

with sync_playwright() as p:

    browser = p.chromium.launch(headless=True)

    page = browser.new_page()

    page.goto(
        "https://www.spanionline.com.br/busca/skol",
        timeout=120000
    )

    print("PAGINA ABERTA")

    page.wait_for_timeout(5000)

    produtos = page.locator("a")

    total = produtos.count()

    print(f"TOTAL LINKS: {total}")

    for i in range(min(total, 30)):

        try:

            texto = produtos.nth(i).inner_text().strip()

            href = produtos.nth(i).get_attribute("href")

            if href and "/produto/" in href:

                print("\n===================")
                print("NOME:")
                print(texto)

                print("LINK:")
                print(href)

        except:
            pass

    browser.close()
