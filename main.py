from playwright.sync_api import sync_playwright

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=True,
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox"
        ]
    )

    context = browser.new_context()

    page = context.new_page()

    # =========================
    # CAPTURA HEADERS
    # =========================

    def capturar(request):

        url = request.url

        if "/produtos/832/detalhes" in url:

            print("\n====================")
            print("API PRODUTO")

            headers = request.headers

            for k, v in headers.items():

                print(f"{k}: {v}")

    page.on("request", capturar)

    page.goto(
        "https://www.spanionline.com.br/produto/832/leite-longa-vida-italac-1l-semidesnatado",
        wait_until="networkidle",
        timeout=120000
    )

    page.wait_for_timeout(15000)

    browser.close()
