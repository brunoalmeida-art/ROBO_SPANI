import requests
import pandas as pd
import time
import smtplib
import os

from datetime import datetime
from email.message import EmailMessage

# =========================================
# CONFIG
# =========================================

LOJA_ID = 1
CD_ID = 36
ORG_ID = 67

SESSION = "dc9e71d5-6b54-4c28-9f5c-0a0ab5dc316e"

headers = {
    "accept": "application/json",
    "user-agent": "Mozilla/5.0"
}

# =========================================
# BUSCAS
# =========================================

buscas = list("abcdefghijklmnopqrstuvwxyz")

# =========================================
# RESULTADOS
# =========================================

dados = []

ja_existe = set()

# =========================================
# LOOP BUSCAS
# =========================================

for termo in buscas:

    pagina = 1
    continuar = True

    while continuar:

        print(f"BUSCA: {termo} | PAGINA: {pagina}")

        url = (
            f"https://services-beta.vipcommerce.com.br/"
            f"api-admin/v1/org/{ORG_ID}/"
            f"filial/{LOJA_ID}/"
            f"centro_distribuicao/{CD_ID}/"
            f"loja/buscas/produtos/termo/{termo}"
            f"?page={pagina}"
            f"&&session={SESSION}"
        )

        # =========================================
        # REQUEST
        # =========================================

        try:

            r = requests.get(
                url,
                headers=headers,
                timeout=120
            )

        except Exception as e:

            print(f"ERRO NA URL: {url}")
            print(e)

            break

        # =========================================
        # STATUS
        # =========================================

        if r.status_code != 200:

            print(f"STATUS ERROR: {r.status_code}")

            break

        try:

            js = r.json()

        except:

            print("ERRO JSON")

            break

        produtos = js.get("produtos", [])

        # =========================================
        # SEM PRODUTOS
        # =========================================

        if not produtos:

            continuar = False
            break

        # =========================================
        # LOOP PRODUTOS
        # =========================================

        for prod in produtos:

            try:

                ean = str(
                    prod.get("codigo_barras", "")
                )

                if ean in ja_existe:
                    continue

                ja_existe.add(ean)

                nome = prod.get("descricao", "")

                preco_varejo = (
                    prod.get("preco", "")
                )

                em_oferta = (
                    prod.get("em_oferta", False)
                )

                oferta = prod.get("oferta", {})

                preco_oferta = ""

                preco_antigo = ""

                qtd_oferta = ""

                # =========================================
                # OFERTA
                # =========================================

                if oferta:

                    preco_oferta = (
                        oferta.get(
                            "preco_oferta",
                            ""
                        )
                    )

                    preco_antigo = (
                        oferta.get(
                            "preco_antigo",
                            ""
                        )
                    )

                    qtd_oferta = (
                        oferta.get(
                            "quantidade_minima",
                            ""
                        )
                    )

                # =========================================
                # LINK
                # =========================================

                link = (
                    "https://www.spani.com.br/"
                    + prod.get("link", "")
                )

                # =========================================
                # LINHA
                # =========================================

                dados.append({

                    "PRODUTO": nome,

                    "PREÇO VAREJO": preco_varejo,

                    "PREÇO OFERTA": preco_oferta,

                    "PREÇO ANTIGO": preco_antigo,

                    "QTD OFERTA": qtd_oferta,

                    "OFERTA": (
                        "SIM"
                        if em_oferta
                        else "NÃO"
                    ),

                    "EAN": ean,

                    "LINK": link

                })

            except Exception as erro_prod:

                print("ERRO PRODUTO")
                print(erro_prod)

        pagina += 1

        # =========================================
        # PAUSA ANTI BLOQUEIO
        # =========================================

        time.sleep(0.5)

# =========================================
# DATAFRAME
# =========================================

df = pd.DataFrame(dados)

# =========================================
# ORDENAR
# =========================================

df = df.sort_values(
    by="PRODUTO"
)

# =========================================
# NOME ARQUIVO
# =========================================

agora = datetime.now().strftime(
    "%Y-%m-%d_%H-%M"
)

arquivo = f"SPANI_{agora}.xlsx"

# =========================================
# EXPORTAR
# =========================================

df.to_excel(
    arquivo,
    index=False
)

print(f"ARQUIVO SALVO: {arquivo}")

# =========================================
# EMAIL
# =========================================

EMAIL_FROM = os.getenv("EMAIL_FROM")
EMAIL_TO = os.getenv("EMAIL_TO")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

msg = EmailMessage()

msg["Subject"] = (
    "SPANI - Atualização Automática"
)

msg["From"] = EMAIL_FROM

msg["To"] = EMAIL_TO

msg.set_content("""

Bom dia,

Segue em anexo o preços atualizado do Site Spani.

Arquivo gerado automaticamente pelo robô.

Att,
Bruno

""")

# =========================================
# ANEXO
# =========================================

with open(arquivo, "rb") as f:

    msg.add_attachment(
        f.read(),
        maintype="application",
        subtype=(
            "vnd.openxmlformats-"
            "officedocument."
            "spreadsheetml.sheet"
        ),
        filename=arquivo
    )

# =========================================
# SMTP
# =========================================

with smtplib.SMTP_SSL(
    "smtp.gmail.com",
    465
) as smtp:

    smtp.login(
        EMAIL_FROM,
        EMAIL_PASSWORD
    )

    smtp.send_message(msg)

print("EMAIL ENVIADO")
