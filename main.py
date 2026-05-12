import requests
import json
import pandas as pd
import smtplib
import os

from email.message import EmailMessage

# =========================================
# EMAIL
# =========================================

EMAIL_USER = os.getenv("EMAIL_USER")

EMAIL_PASS = os.getenv("EMAIL_PASS")

EMAIL_TO = "pricing@roldao.com.br"

# =========================================
# URL
# =========================================

url = (
    "https://services-beta.vipcommerce.com.br/"
    "api-admin/v1/org/67/"
    "filial/1/"
    "centro_distribuicao/36/"
    "loja/buscas/produtos/termo/ar"
    "?page=1"
    "&&session=dc9e71d5-6b54-4c28-9f5c-0a0ab5dc316e"
)

# =========================================
# HEADERS
# =========================================

headers = {

    "accept": "application/json",

    "user-agent": "Mozilla/5.0",

    "OrganizationId": "67",

    "DomainKey": "spanionline.com.br",

    "Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJ2aXBjb21tZXJjZSIsImF1ZCI6ImFwaS1hZG1pbiIsInN1YiI6IjZiYzQ4NjdlLWRjYTktMTFlOS04NzQyLTAyMGQ3OTM1OWNhMCIsInZpcGNvbW1lcmNlQ2xpZW50ZUlkIjpudWxsLCJpYXQiOjE3Nzc5MDc4MTMsInZlciI6MSwiY2xpZW50IjpudWxsLCJvcGVyYXRvciI6bnVsbCwib3JnIjoiNjcifQ.mqyEyNRMBcY0rb4kWeNN0-xnEb8kus9i97w3IR6qjCCPdKEyBjUcZkF77_4KtKvHBI2cx25Fd8E9G4Q1cwsADw"
}

# =========================================
# REQUEST
# =========================================

r = requests.get(
    url,
    headers=headers,
    timeout=120
)

# =========================================
# STATUS
# =========================================

print("STATUS:", r.status_code)

# =========================================
# TEXTO BRUTO
# =========================================

print("\n======== RESPOSTA ========\n")

print(r.text[:10000])

# =========================================
# JSON
# =========================================

try:

    js = r.json()

    print("\n======== JSON FORMATADO ========\n")

    print(
        json.dumps(
            js,
            indent=2,
            ensure_ascii=False
        )
    )

    # =====================================
    # PRODUTOS
    # =====================================

    produtos = js.get("data", {}).get("produtos", [])

    print(f"\nTOTAL PRODUTOS: {len(produtos)}")

    # =====================================
    # EXCEL
    # =====================================

    if produtos:

        df = pd.DataFrame(produtos)

        arquivo = "SPANI.xlsx"

        df.to_excel(
            arquivo,
            index=False
        )

        print(f"\nEXCEL GERADO: {arquivo}")

        # =================================
        # EMAIL
        # =================================

        msg = EmailMessage()

        msg["Subject"] = "SPANI - Atualizacao de Produtos"

        msg["From"] = EMAIL_USER

        msg["To"] = EMAIL_TO

        msg.set_content(
            "Arquivo SPANI.xlsx em anexo."
        )

        with open(arquivo, "rb") as f:

            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=arquivo
            )

        with smtplib.SMTP_SSL(
            "smtp.gmail.com",
            465
        ) as smtp:

            smtp.login(
                EMAIL_USER,
                EMAIL_PASS
            )

            smtp.send_message(msg)

        print("\nEMAIL ENVIADO COM SUCESSO")

    else:

        print("\nNENHUM PRODUTO ENCONTRADO")

except Exception as e:

    print("\nERRO JSON\n")

    print(e)
