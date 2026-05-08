from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from docxtpl import DocxTemplate

from datetime import datetime

import subprocess
import os
import re
import time
import requests

# =========================================================
# APP
# =========================================================

app = FastAPI(
    title="DOMUM PDF API",
    version="3.0.0",
    description="""
API profissional para geração automática de:

- Propostas comerciais
- Contratos empresariais
- Webhook WhatsApp Meta Cloud API

DOMUM Engenharia
"""
)

# =========================================================
# CONFIGURAÇÕES
# =========================================================

PASTA_SAIDA = "pdfs_gerados"

PROPOSTA_MODELO = "proposta_modelo.docx"
CONTRATO_MODELO = "contrato_modelo.docx"

# =========================================================
# URL BASE
# =========================================================

BASE_URL = os.getenv(
    "BASE_URL",
    "https://web-production-a841e.up.railway.app"
)

# =========================================================
# WHATSAPP META CONFIG
# =========================================================

VERIFY_TOKEN = "domum_verify"

WHATSAPP_TOKEN = os.getenv(
    "WHATSAPP_TOKEN",
    ""
)

PHONE_NUMBER_ID = os.getenv(
    "PHONE_NUMBER_ID",
    ""
)

# =========================================================
# CRIAR PASTA AUTOMATICAMENTE
# =========================================================

os.makedirs(
    PASTA_SAIDA,
    exist_ok=True
)

# =========================================================
# SERVIR PDFs PUBLICAMENTE
# =========================================================

app.mount(
    "/arquivos",
    StaticFiles(directory=PASTA_SAIDA),
    name="arquivos"
)

# =========================================================
# UTILIDADES
# =========================================================

def data_atual():

    meses = [
        "janeiro",
        "fevereiro",
        "março",
        "abril",
        "maio",
        "junho",
        "julho",
        "agosto",
        "setembro",
        "outubro",
        "novembro",
        "dezembro"
    ]

    hoje = datetime.now()

    return (
        f"{hoje.day} de "
        f"{meses[hoje.month - 1]} de "
        f"{hoje.year}"
    )


def limpar_nome(texto):

    texto = texto.strip().lower()

    texto = re.sub(
        r"[^\w\s-]",
        "",
        texto
    )

    texto = re.sub(
        r"\s+",
        "_",
        texto
    )

    return texto


def formatar_servicos(servicos):

    texto = ""

    for indice, servico in enumerate(
        servicos,
        start=1
    ):

        texto += (
            f"2.{indice}. "
            f"{servico}\n"
        )

    return texto.strip()


def converter_para_pdf(docx_saida):

    """
    Converte DOCX para PDF usando LibreOffice.
    Compatível com Railway/Linux.
    """

    subprocess.run([

        "libreoffice",

        "--headless",

        "--convert-to",
        "pdf",

        "--outdir",
        PASTA_SAIDA,

        docx_saida

    ])

    time.sleep(2)


def enviar_whatsapp(numero, mensagem):

    """
    Envia mensagem via Meta Cloud API
    """

    if not WHATSAPP_TOKEN:
        print("WHATSAPP_TOKEN não configurado")
        return

    if not PHONE_NUMBER_ID:
        print("PHONE_NUMBER_ID não configurado")
        return

    url = (
        f"https://graph.facebook.com/v21.0/"
        f"{PHONE_NUMBER_ID}/messages"
    )

    headers = {
        "Authorization": f"Bearer {WHATSAPP_TOKEN}",
        "Content-Type": "application/json"
    }

    payload = {
        "messaging_product": "whatsapp",
        "to": numero,
        "type": "text",
        "text": {
            "body": mensagem
        }
    }

    response = requests.post(
        url,
        headers=headers,
        json=payload
    )

    print(response.text)

# =========================================================
# MODELS
# =========================================================

class PropostaRequest(BaseModel):

    cliente: str

    cidade: str

    objeto: str

    servicos: list[str]

    valor: str

    pagamento: str


class ContratoRequest(BaseModel):

    contratante: str

    cpf: str

    rg: str

    profissao: str

    endereco_proprietario: str

    endereco_obra: str

    telefone: str

    email: str

    cidade: str

    objeto: str

    servicos: list[str]

    valor: str

    pagamento: str

# =========================================================
# STATUS API
# =========================================================

@app.get("/")

def home():

    return {

        "status":
        "API DOMUM ONLINE",

        "versao":
        "3.0.0"
    }

# =========================================================
# WEBHOOK META
# =========================================================

from fastapi import Query

 @app.get("/webhook")
async def verify_webhook(
    hub_mode: str = Query(None, alias="hub.mode"),
    hub_verify_token: str = Query(None, alias="hub.verify_token"),
    hub_challenge: str = Query(None, alias="hub.challenge")
):

    print("VALIDANDO WEBHOOK...")
    print(hub_verify_token)

    if hub_verify_token == VERIFY_TOKEN:
        return int(hub_challenge)

    return {
        "error": "Token inválido"
    }


@app.post("/webhook")
async def receive_webhook(request: Request):

    data = await request.json()

    print("\n============================")
    print("MENSAGEM RECEBIDA")
    print("============================")
    print(data)

    try:

        entry = data["entry"][0]

        changes = entry["changes"][0]

        value = changes["value"]

        if "messages" in value:

            mensagem = value["messages"][0]

            numero = mensagem["from"]

            if mensagem["type"] == "text":

                texto = mensagem["text"]["body"]

                print(f"\nMensagem de {numero}")
                print(texto)

                enviar_whatsapp(
                    numero,
                    f"DOMUM recebeu sua mensagem: {texto}"
                )

    except Exception as erro:

        print(f"Erro webhook: {erro}")

    return {
        "status": "ok"
    }

# =========================================================
# GERAR PROPOSTA
# =========================================================

@app.post("/gerar-proposta")

def gerar_proposta(
    dados: PropostaRequest
):

    try:

        contexto = {

            "CLIENTE":
            dados.cliente,

            "CIDADE":
            dados.cidade,

            "OBJETO":
            dados.objeto,

            "SERVICOS":
            formatar_servicos(
                dados.servicos
            ),

            "VALOR":
            dados.valor,

            "PAGAMENTO":
            dados.pagamento,

            "DATA":
            data_atual()
        }

        nome = limpar_nome(
            dados.cliente
        )

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        base = (
            f"proposta_{nome}_{timestamp}"
        )

        docx_saida = os.path.join(
            PASTA_SAIDA,
            f"{base}.docx"
        )

        pdf_saida = os.path.join(
            PASTA_SAIDA,
            f"{base}.pdf"
        )

        doc = DocxTemplate(
            PROPOSTA_MODELO
        )

        doc.render(
            contexto
        )

        doc.save(
            docx_saida
        )

        converter_para_pdf(
            docx_saida
        )

        if not os.path.exists(
            pdf_saida
        ):

            return {

                "status":
                "erro",

                "mensagem":
                "PDF não foi gerado."
            }

        url_pdf = (
            f"{BASE_URL}/arquivos/"
            f"{os.path.basename(pdf_saida)}"
        )

        return {

            "status":
            "ok",

            "tipo":
            "proposta",

            "arquivo":
            os.path.basename(
                pdf_saida
            ),

            "url":
            url_pdf
        }

    except Exception as erro:

        return {

            "status":
            "erro",

            "mensagem":
            str(erro)
        }

# =========================================================
# GERAR CONTRATO
# =========================================================

@app.post("/gerar-contrato")

def gerar_contrato(
    dados: ContratoRequest
):

    try:

        contexto = {

            "CONTRATANTE":
            dados.contratante,

            "CPF":
            dados.cpf,

            "RG":
            dados.rg,

            "PROFISSAO":
            dados.profissao,

            "ENDERECO_PROPRIETARIO":
            dados.endereco_proprietario,

            "ENDERECO_OBRA":
            dados.endereco_obra,

            "TELEFONE":
            dados.telefone,

            "EMAIL":
            dados.email,

            "CIDADE":
            dados.cidade,

            "OBJETO":
            dados.objeto,

            "SERVICOS":
            formatar_servicos(
                dados.servicos
            ),

            "VALOR":
            dados.valor,

            "PAGAMENTO":
            dados.pagamento,

            "DATA":
            data_atual()
        }

        nome = limpar_nome(
            dados.contratante
        )

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        base = (
            f"contrato_{nome}_{timestamp}"
        )

        docx_saida = os.path.join(
            PASTA_SAIDA,
            f"{base}.docx"
        )

        pdf_saida = os.path.join(
            PASTA_SAIDA,
            f"{base}.pdf"
        )

        doc = DocxTemplate(
            CONTRATO_MODELO
        )

        doc.render(
            contexto
        )

        doc.save(
            docx_saida
        )

        converter_para_pdf(
            docx_saida
        )

        if not os.path.exists(
            pdf_saida
        ):

            return {

                "status":
                "erro",

                "mensagem":
                "PDF não foi gerado."
            }

        url_pdf = (
            f"{BASE_URL}/arquivos/"
            f"{os.path.basename(pdf_saida)}"
        )

        return {

            "status":
            "ok",

            "tipo":
            "contrato",

            "arquivo":
            os.path.basename(
                pdf_saida
            ),

            "url":
            url_pdf
        }

    except Exception as erro:

        return {

            "status":
            "erro",

            "mensagem":
            str(erro)
        }