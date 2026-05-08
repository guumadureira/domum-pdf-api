from fastapi import FastAPI, Request, Query
from fastapi.responses import PlainTextResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from pydantic import BaseModel
from typing import List

from docxtpl import DocxTemplate

from openai import OpenAI

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
    version="4.0.0",
    description="""
API profissional para:

- WhatsApp Meta Cloud API
- Atendimento IA com OpenAI
- Geração automática de propostas
- Geração automática de contratos
- PDFs automáticos

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

VERIFY_TOKEN = os.getenv(
    "VERIFY_TOKEN",
    "domum_verify"
)

WHATSAPP_TOKEN = os.getenv(
    "WHATSAPP_TOKEN",
    ""
)

PHONE_NUMBER_ID = os.getenv(
    "PHONE_NUMBER_ID",
    ""
)

# =========================================================
# OPENAI
# =========================================================

OPENAI_API_KEY = os.getenv(
    "OPENAI_API_KEY",
    ""
)

client = OpenAI(
    api_key=OPENAI_API_KEY
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

    try:

        processo = subprocess.run(

            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                PASTA_SAIDA,
                docx_saida
            ],

            capture_output=True,
            text=True
        )

        print(processo.stdout)
        print(processo.stderr)

        time.sleep(2)

    except Exception as erro:

        print(f"Erro ao converter PDF: {erro}")

        raise Exception(
            "LibreOffice não encontrado no servidor."
        )

# =========================================================
# OPENAI IA
# =========================================================

def responder_ia(mensagem_usuario):

    try:

        resposta = client.chat.completions.create(

            model="gpt-4.1-mini",

            messages=[

                {
                    "role": "system",
                    "content": """
Você é a assistente virtual da DOMUM Engenharia.

Seu comportamento:

- profissional
- objetiva
- amigável
- clara
- especialista em engenharia

Você ajuda clientes com:
- propostas
- contratos
- obras
- projetos
- engenharia civil

Nunca invente informações.
"""
                },

                {
                    "role": "user",
                    "content": mensagem_usuario
                }
            ],

            temperature=0.7
        )

        return resposta.choices[0].message.content

    except Exception as erro:

        print(f"Erro OpenAI: {erro}")

        return (
            "Desculpe, ocorreu um erro no atendimento."
        )

# =========================================================
# WHATSAPP
# =========================================================

def enviar_whatsapp(numero, mensagem):

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

    try:

        response = requests.post(
            url,
            headers=headers,
            json=payload,
            timeout=30
        )

        print(response.status_code)
        print(response.text)

    except Exception as erro:

        print(f"Erro WhatsApp: {erro}")

# =========================================================
# MODELS
# =========================================================

class PropostaRequest(BaseModel):

    cliente: str
    cidade: str
    objeto: str
    servicos: List[str]
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
    servicos: List[str]
    valor: str
    pagamento: str

# =========================================================
# STATUS API
# =========================================================

@app.get("/")

def home():

    return {
        "status": "API DOMUM ONLINE",
        "versao": "4.0.0"
    }

# =========================================================
# WEBHOOK META
# =========================================================

@app.get("/webhook")
async def verify_webhook(

    hub_mode: str = Query(None, alias="hub.mode"),
    hub_verify_token: str = Query(None, alias="hub.verify_token"),
    hub_challenge: str = Query(None, alias="hub.challenge")

):

    print("VALIDANDO WEBHOOK...")

    if hub_verify_token == VERIFY_TOKEN:

        return PlainTextResponse(
            content=hub_challenge
        )

    return JSONResponse(

        status_code=403,

        content={
            "error": "Token inválido"
        }
    )

@app.post("/webhook")
async def receive_webhook(request: Request):

    try:

        data = await request.json()

        print("\n============================")
        print("MENSAGEM RECEBIDA")
        print("============================")
        print(data)

        entry = data["entry"][0]

        changes = entry["changes"][0]

        value = changes["value"]

        # Ignora status Meta
        if "statuses" in value:

            return {
                "status": "status_update"
            }

        # Sem mensagens
        if "messages" not in value:

            return {
                "status": "sem_mensagem"
            }

        mensagem = value["messages"][0]

        numero = mensagem["from"]

        tipo = mensagem["type"]

        print(f"\nMensagem de {numero}")
        print(f"Tipo: {tipo}")

        if tipo != "text":

            return {
                "status": "tipo_nao_suportado"
            }

        texto = mensagem["text"]["body"]

        print("\nTEXTO:")
        print(texto)

        # =========================================================
        # IA RESPONDE
        # =========================================================

        resposta_ia = responder_ia(texto)

        print("\nRESPOSTA IA:")
        print(resposta_ia)

        enviar_whatsapp(
            numero,
            resposta_ia
        )

        return {
            "status": "ok"
        }

    except Exception as erro:

        print(f"Erro webhook: {erro}")

        return JSONResponse(

            status_code=500,

            content={
                "status": "erro",
                "mensagem": str(erro)
            }
        )

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

            return JSONResponse(

                status_code=500,

                content={
                    "status": "erro",
                    "mensagem": "PDF não foi gerado."
                }
            )

        url_pdf = (
            f"{BASE_URL}/arquivos/"
            f"{os.path.basename(pdf_saida)}"
        )

        return {

            "status": "ok",

            "tipo": "proposta",

            "arquivo":
            os.path.basename(
                pdf_saida
            ),

            "url":
            url_pdf
        }

    except Exception as erro:

        return JSONResponse(

            status_code=500,

            content={
                "status": "erro",
                "mensagem": str(erro)
            }
        )

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

            return JSONResponse(

                status_code=500,

                content={
                    "status": "erro",
                    "mensagem": "PDF não foi gerado."
                }
            )

        url_pdf = (
            f"{BASE_URL}/arquivos/"
            f"{os.path.basename(pdf_saida)}"
        )

        return {

            "status": "ok",

            "tipo": "contrato",

            "arquivo":
            os.path.basename(
                pdf_saida
            ),

            "url":
            url_pdf
        }

    except Exception as erro:

        return JSONResponse(

            status_code=500,

            content={
                "status": "erro",
                "mensagem": str(erro)
            }
        )