from fastapi import FastAPI
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from pydantic import BaseModel
from typing import List

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
    title="DOMUM Engenharia API",
    version="5.0.0",
    description="""
API profissional para geração automática de:

- Propostas comerciais
- Contratos empresariais
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

BASE_URL = os.getenv(
    "BASE_URL",
    "https://web-production-a841e.up.railway.app"
)

# =========================================================
# PASTA PDFs
# =========================================================

os.makedirs(
    PASTA_SAIDA,
    exist_ok=True
)

# =========================================================
# ARQUIVOS PÚBLICOS
# =========================================================

app.mount(
    "/arquivos",
    StaticFiles(directory=PASTA_SAIDA),
    name="arquivos"
)

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

    # =====================================================
    # CONTROLE COMPLEMENTARES
    # =====================================================

    nao_inclui_complementares: bool = True

    # =====================================================
    # PROJETOS COMPLEMENTARES
    # =====================================================

    projeto_estrutural: bool = False
    projeto_eletrico: bool = False
    projeto_hidraulico: bool = False
    projeto_ar_condicionado: bool = False
    projeto_esgoto: bool = False
    projeto_automacao: bool = False
    projeto_som: bool = False
    projeto_telefonia: bool = False
    cabeamento_estruturado: bool = False

    # =====================================================
    # DETALHAMENTOS
    # =====================================================

    detalhamento_marcenaria: bool = False
    detalhamento_marmoraria: bool = False

    # =====================================================
    # PLANTAS
    # =====================================================

    planta_forro: bool = False
    planta_paginacao_piso: bool = False

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

    return texto or "documento"


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
# STATUS API
# =========================================================

@app.get("/")
def home():

    return {
        "status": "API DOMUM ONLINE",
        "versao": "5.0.0"
    }

# =========================================================
# PRIVACY POLICY
# =========================================================

@app.get("/privacy", response_class=HTMLResponse)
def privacy_policy():

    return """
    <html>

        <head>
            <title>
                Política de Privacidade - DOMUM Engenharia
            </title>
        </head>

        <body style="font-family: Arial; padding: 40px;">

            <h1>
                Política de Privacidade
            </h1>

            <p>
                A DOMUM Engenharia respeita sua privacidade.
            </p>

            <p>
                Os dados enviados para esta aplicação são utilizados
                exclusivamente para geração de propostas comerciais,
                contratos e documentos empresariais.
            </p>

            <p>
                Nenhuma informação é vendida ou compartilhada
                com terceiros.
            </p>

            <p>
                DOMUM Engenharia
            </p>

        </body>

    </html>
    """

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

        # =====================================================
        # GERAR DOCX
        # =====================================================

        doc = DocxTemplate(
            PROPOSTA_MODELO
        )

        doc.render(
            contexto
        )

        doc.save(
            docx_saida
        )

        # =====================================================
        # GERAR PDF
        # =====================================================

        converter_para_pdf(
            docx_saida
        )

        # =====================================================
        # VALIDAR PDF
        # =====================================================

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

        # =====================================================
        # URL PÚBLICA
        # =====================================================

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

        return JSONResponse(

            status_code=500,

            content={

                "status":
                "erro",

                "mensagem":
                str(erro)
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
            data_atual(),

"NAO_INCLUI_COMPLEMENTARES":
dados.nao_inclui_complementares,

"PROJETO_ESTRUTURAL":
dados.projeto_estrutural,

"PROJETO_ELETRICO":
dados.projeto_eletrico,

"PROJETO_HIDRAULICO":
dados.projeto_hidraulico,

"PROJETO_AR_CONDICIONADO":
dados.projeto_ar_condicionado,

"PROJETO_ESGOTO":
dados.projeto_esgoto,

"PROJETO_AUTOMACAO":
dados.projeto_automacao,

"PROJETO_SOM":
dados.projeto_som,

"PROJETO_TELEFONIA":
dados.projeto_telefonia,

"CABEAMENTO_ESTRUTURADO":
dados.cabeamento_estruturado,

"DETALHAMENTO_MARCENARIA":
dados.detalhamento_marcenaria,

"DETALHAMENTO_MARMORARIA":
dados.detalhamento_marmoraria,

"PLANTA_FORRO":
dados.planta_forro,

"PLANTA_PAGINACAO_PISO":
dados.planta_paginacao_piso,
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

        # =====================================================
        # GERAR DOCX
        # =====================================================

        doc = DocxTemplate(
            CONTRATO_MODELO
        )

        doc.render(
            contexto
        )

        doc.save(
            docx_saida
        )

        # =====================================================
        # GERAR PDF
        # =====================================================

        converter_para_pdf(
            docx_saida
        )

        # =====================================================
        # VALIDAR PDF
        # =====================================================

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

        # =====================================================
        # URL PÚBLICA
        # =====================================================

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

        return JSONResponse(

            status_code=500,

            content={

                "status":
                "erro",

                "mensagem":
                str(erro)
            }
        )