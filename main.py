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
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36",
        viewport={"width": 1920, "height": 1080},
        locale="pt-BR"
    )

    page = context.new_page()

    # remove webdriver flag
    page.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined
        })
    """)

    url = "https://www.spanionline.com.br/produto/832/leite-longa-vida-italac-1l-semidesnatado"

    page.goto(
        url,
        wait_until="networkidle",
        timeout=120000
    )

    page.wait_for_timeout(10000)

    print("URL FINAL:")
    print(page.url)

    print("\nTITULO:")
    print(page.title())

    print("\nHTML INICIO:")
    print(page.content()[:5000])

    browser.close()
