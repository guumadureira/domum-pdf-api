from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from docxtpl import DocxTemplate

from datetime import datetime

import subprocess
import os
import re
import time

# =========================================================
# APP
# =========================================================

app = FastAPI(
    title="DOMUM PDF API",
    version="2.0.0",
    description="""
API profissional para geração automática de:

- Propostas comerciais
- Contratos empresariais

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
# ALTERAR APÓS DEPLOY NO RAILWAY
# =========================================================

BASE_URL = os.getenv(
    "BASE_URL",
    "https://web-production-a841e.up.railway.app"
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
        "2.0.0"
    }

# =========================================================
# GERAR PROPOSTA
# =========================================================

@app.post("/gerar-proposta")

def gerar_proposta(
    dados: PropostaRequest
):

    try:

        # ======================================
        # CONTEXTO
        # ======================================

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

        # ======================================
        # NOMES DOS ARQUIVOS
        # ======================================

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

        # ======================================
        # GERAR DOCX
        # ======================================

        doc = DocxTemplate(
            PROPOSTA_MODELO
        )

        doc.render(
            contexto
        )

        doc.save(
            docx_saida
        )

        # ======================================
        # CONVERTER PDF
        # ======================================

        converter_para_pdf(
            docx_saida
        )

        # ======================================
        # VALIDAR PDF
        # ======================================

        if not os.path.exists(
            pdf_saida
        ):

            return {

                "status":
                "erro",

                "mensagem":
                "PDF não foi gerado."
            }

        # ======================================
        # URL PÚBLICA
        # ======================================

        url_pdf = (
            f"{BASE_URL}/arquivos/"
            f"{os.path.basename(pdf_saida)}"
        )

        # ======================================
        # RETORNO
        # ======================================

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

        # ======================================
        # CONTEXTO
        # ======================================

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

        # ======================================
        # NOMES DOS ARQUIVOS
        # ======================================

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

        # ======================================
        # GERAR DOCX
        # ======================================

        doc = DocxTemplate(
            CONTRATO_MODELO
        )

        doc.render(
            contexto
        )

        doc.save(
            docx_saida
        )

        # ======================================
        # CONVERTER PDF
        # ======================================

        converter_para_pdf(
            docx_saida
        )

        # ======================================
        # VALIDAR PDF
        # ======================================

        if not os.path.exists(
            pdf_saida
        ):

            return {

                "status":
                "erro",

                "mensagem":
                "PDF não foi gerado."
            }

        # ======================================
        # URL PÚBLICA
        # ======================================

        url_pdf = (
            f"{BASE_URL}/arquivos/"
            f"{os.path.basename(pdf_saida)}"
        )

        # ======================================
        # RETORNO
        # ======================================

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

from fastapi import Request

VERIFY_TOKEN = "domum_verify"

@app.get("/webhook")
async def verify_webhook(
    hub_mode: str = None,
    hub_verify_token: str = None,
    hub_challenge: str = None
):
    if hub_verify_token == VERIFY_TOKEN:
        return int(hub_challenge)

    return {"error": "Token inválido"}


@app.post("/webhook")
async def receive_webhook(request: Request):
    data = await request.json()

    print("Mensagem recebida:")
    print(data)

    return {"status": "ok"}