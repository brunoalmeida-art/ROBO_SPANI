from playwright.sync_api import sync_playwright

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=True,
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--disable-dev-shm-usage"
        ]
    )

    context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36"
    )

    page = context.new_page()

    # =========================
    # MONITORA REQUESTS
    # =========================

    def capturar(response):

        url = response.url

        if "api" in url.lower():

            print("\n====================")
            print("API:")
            print(url)

    page.on("response", capturar)

    # =========================
    # ABRE PRODUTO
    # =========================

    page.goto(
        "https://www.spanionline.com.br/produto/832/leite-longa-vida-italac-1l-semidesnatado",
        wait_until="networkidle",
        timeout=120000
    )

    page.wait_for_timeout(15000)

    browser.close()
