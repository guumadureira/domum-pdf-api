# =========================================================
# SISTEMA DOMUM ENGENHARIA
# GERADOR PROFISSIONAL DE DOCUMENTOS
# Python 3.11+
# =========================================================

# INSTALAR:
# pip install docxtpl python-docx docx2pdf fpdf2 pillow

from docxtpl import DocxTemplate
from docx2pdf import convert
from fpdf import FPDF
from datetime import datetime
import os
import re

# =========================================================
# CONFIGURAÇÕES
# =========================================================

PASTA_SAIDA = "pdfs_gerados"

PROPOSTA_MODELO = "proposta_modelo.docx"
CONTRATO_MODELO = "contrato_modelo.docx"

PAPEL_TIMBRADO = "papel_timbrado.png"

os.makedirs(PASTA_SAIDA, exist_ok=True)

# =========================================================
# UTILIDADES
# =========================================================

def data_atual():

    meses = [
        "janeiro", "fevereiro", "março",
        "abril", "maio", "junho",
        "julho", "agosto", "setembro",
        "outubro", "novembro", "dezembro"
    ]

    hoje = datetime.now()

    return f"{hoje.day} de {meses[hoje.month - 1]} de {hoje.year}"


def limpar_nome(texto):

    texto = texto.strip().lower()

    texto = re.sub(r"[^\w\s-]", "", texto)

    texto = re.sub(r"\s+", "_", texto)

    return texto or "documento"


def linha():

    print("=" * 60)


def confirmar():

    resposta = input(
        "\nConfirmar geração? (S/N): "
    ).strip().upper()

    return resposta == "S"


def solicitar_texto(campo):

    while True:

        valor = input(f"{campo}: ").strip()

        if valor:

            return valor

        print("Campo obrigatório.")


def solicitar_lista(titulo):

    print(f"\n{titulo}")

    print("Digite um item por vez.")
    print("Digite FIM para finalizar.\n")

    itens = []

    while True:

        item = input("Item: ").strip()

        if item.upper() == "FIM":

            break

        if item:

            itens.append(item)

    return itens


def validar_arquivo(arquivo):

    if not os.path.exists(arquivo):

        print(f"\nArquivo não encontrado: {arquivo}")

        return False

    return True


def formatar_servicos(servicos):

    texto = ""

    for indice, servico in enumerate(
        servicos,
        start=1
    ):

        texto += f"2.{indice}. {servico}\n"

    return texto.strip()

def converter_docx_para_pdf(docx_saida, pdf_saida):

    try:

        convert(docx_saida, pdf_saida)

        return True

    except Exception as erro:

        print("\nERRO AO CONVERTER PARA PDF")
        print("Verifique se o Microsoft Word está instalado.")

        print(erro)

        return False

# =========================================================
# PROPOSTA
# =========================================================

def gerar_proposta():

    if not validar_arquivo(PROPOSTA_MODELO):

        return

    print("\nGERADOR DE PROPOSTA\n")

    # ==========================================
    # DADOS DA PROPOSTA
    # ==========================================

    dados = {

        # ======================================
        # CLIENTE
        # ======================================

        "CLIENTE": solicitar_texto(
            "Nome do cliente"
        ),

        # ======================================
        # LOCAL
        # ======================================

        "CIDADE": solicitar_texto(
            "Cidade"
        ),

        # ======================================
        # OBJETO
        # ======================================

        "OBJETO": solicitar_texto(
            "Objeto do serviço"
        ),

        # ======================================
        # SERVIÇOS
        # ======================================

        "SERVICOS": formatar_servicos(
            solicitar_lista(
                "Serviços executados"
            )
        ),

        # ======================================
        # FINANCEIRO
        # ======================================

        "VALOR": solicitar_texto(
            "Valor"
        ),

        "PAGAMENTO": solicitar_texto(
            "Forma de pagamento"
        ),

        # ======================================
        # DATA
        # ======================================

        "DATA": data_atual(),
    }

    # ==========================================
    # PRÉVIA
    # ==========================================

    linha()

    print("PRÉVIA DA PROPOSTA")

    linha()

    print(f"Cliente: {dados['CLIENTE']}")
    print(f"Cidade: {dados['CIDADE']}")
    print(f"Objeto: {dados['OBJETO']}")

    print("\nServiços:\n")

    print(dados["SERVICOS"])

    print(f"\nValor: R$ {dados['VALOR']}")
    print(f"Pagamento: {dados['PAGAMENTO']}")
    print(f"Data: {dados['DATA']}")

    linha()

    if not confirmar():

        print("\nOperação cancelada.")

        return

    # ==========================================
    # NOMES DOS ARQUIVOS
    # ==========================================

    nome = limpar_nome(
        dados["CLIENTE"]
    )

    base = (
        f"proposta_{nome}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    )

    docx_saida = os.path.join(
        PASTA_SAIDA,
        f"{base}.docx"
    )

    pdf_saida = os.path.join(
        PASTA_SAIDA,
        f"{base}.pdf"
    )

    # ==========================================
    # GERAR DOCX
    # ==========================================

    doc = DocxTemplate(
        PROPOSTA_MODELO
    )

    doc.render(dados)

    doc.save(docx_saida)

    # ==========================================
    # GERAR PDF
    # ==========================================

    if converter_docx_para_pdf(
        docx_saida,
        pdf_saida
    ):

        print("\nPROPOSTA GERADA COM SUCESSO:")

        print(pdf_saida)

# =========================================================
# CONTRATO
# =========================================================

def gerar_contrato():

    if not validar_arquivo(CONTRATO_MODELO):

        return

    print("\nGERADOR DE CONTRATO\n")

    dados = {

        # ==========================================
        # DADOS DO CONTRATANTE
        # ==========================================

        "CONTRATANTE": solicitar_texto(
            "Nome do contratante"
        ),

        "CPF": solicitar_texto(
            "CPF"
        ),

        "RG": solicitar_texto(
            "RG"
        ),

        "PROFISSAO": solicitar_texto(
            "Profissão"
        ),

        # ==========================================
        # ENDEREÇOS
        # ==========================================

        "ENDERECO_PROPRIETARIO": solicitar_texto(
            "Endereço do proprietário"
        ),

        "ENDERECO_OBRA": solicitar_texto(
            "Endereço da obra"
        ),

        # ==========================================
        # CONTATO
        # ==========================================

        "TELEFONE": solicitar_texto(
            "Telefone"
        ),

        "EMAIL": solicitar_texto(
            "E-mail"
        ),

        # ==========================================
        # LOCAL
        # ==========================================

        "CIDADE": solicitar_texto(
            "Cidade"
        ),

        # ==========================================
        # CONTRATO
        # ==========================================

        "OBJETO": solicitar_texto(
            "Objeto do contrato"
        ),

        # ==========================================
        # SERVIÇOS
        # ==========================================

        "SERVICOS": formatar_servicos(
            solicitar_lista(
                "Serviços contratados"
            )
        ),

        # ==========================================
        # FINANCEIRO
        # ==========================================

        "VALOR": solicitar_texto(
            "Valor"
        ),

        "PAGAMENTO": solicitar_texto(
            "Forma de pagamento"
        ),

        # ==========================================
        # DATA
        # ==========================================

        "DATA": data_atual(),
    }

    # ==========================================
    # PRÉVIA
    # ==========================================

    linha()

    print("PRÉVIA DO CONTRATO")

    linha()

    print(f"Contratante: {dados['CONTRATANTE']}")
    print(f"CPF: {dados['CPF']}")
    print(f"RG: {dados['RG']}")
    print(f"Profissão: {dados['PROFISSAO']}")

    print(
        f"\nEndereço proprietário:"
    )

    print(
        dados['ENDERECO_PROPRIETARIO']
    )

    print(
        f"\nEndereço da obra:"
    )

    print(
        dados['ENDERECO_OBRA']
    )

    print(f"\nTelefone: {dados['TELEFONE']}")
    print(f"E-mail: {dados['EMAIL']}")
    print(f"Cidade: {dados['CIDADE']}")

    print(f"\nObjeto:")

    print(dados['OBJETO'])

    print("\nServiços:\n")

    print(dados["SERVICOS"])

    print(f"\nValor: R$ {dados['VALOR']}")
    print(f"Pagamento: {dados['PAGAMENTO']}")
    print(f"Data: {dados['DATA']}")

    linha()

    if not confirmar():

        print("\nOperação cancelada.")

        return

    # ==========================================
    # NOMES DOS ARQUIVOS
    # ==========================================

    nome = limpar_nome(
        dados["CONTRATANTE"]
    )

    base = (
        f"contrato_{nome}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    )

    docx_saida = os.path.join(
        PASTA_SAIDA,
        f"{base}.docx"
    )

    pdf_saida = os.path.join(
        PASTA_SAIDA,
        f"{base}.pdf"
    )

    # ==========================================
    # GERAR DOCX
    # ==========================================

    doc = DocxTemplate(
        CONTRATO_MODELO
    )

    doc.render(dados)

    doc.save(docx_saida)

    # ==========================================
    # GERAR PDF
    # ==========================================

    if converter_docx_para_pdf(
        docx_saida,
        pdf_saida
    ):

        print("\nCONTRATO GERADO COM SUCESSO:")

        print(pdf_saida)

# =========================================================
# LICITAÇÃO
# =========================================================

class PDFLicitacao(FPDF):

    def header(self):

        if os.path.exists(PAPEL_TIMBRADO):

            self.image(
                PAPEL_TIMBRADO,
                x=0,
                y=0,
                w=210,
                h=297
            )

        self.set_y(42)

    def footer(self):

        pass


def gerar_licitacao():

    print("\nGERADOR DE LICITAÇÃO\n")

    print("Cole o texto completo.")
    print("Digite FIM em uma linha separada.\n")

    linhas = []

    while True:

        linha_digitada = input()

        if linha_digitada.strip().upper() == "FIM":

            break

        linhas.append(linha_digitada)

    texto = "\n".join(linhas)

    nome = solicitar_texto(
        "Nome do arquivo"
    )

    linha()

    print("PRÉVIA DA LICITAÇÃO")

    linha()

    print(texto[:1500])

    if len(texto) > 1500:

        print("\n...[prévia limitada]")

    linha()

    if not confirmar():

        print("\nOperação cancelada.")

        return

    pdf = PDFLicitacao(
        orientation="P",
        unit="mm",
        format="A4"
    )

    pdf.set_auto_page_break(
        auto=True,
        margin=22
    )

    pdf.set_left_margin(20)

    pdf.set_right_margin(20)

    pdf.add_page()

    pdf.set_font(
        "Helvetica",
        "",
        10
    )

    pdf.multi_cell(
        0,
        6,
        texto
    )

    base = (
        f"licitacao_{limpar_nome(nome)}_"
        f"{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    )

    pdf_saida = os.path.join(
        PASTA_SAIDA,
        f"{base}.pdf"
    )

    pdf.output(pdf_saida)

    print("\nLICITAÇÃO GERADA COM SUCESSO:")

    print(pdf_saida)

# =========================================================
# MENU
# =========================================================

def menu():

    while True:

        linha()

        print(
            "DOMUM ENGENHARIA - GERADOR DE DOCUMENTOS"
        )

        linha()

        print("1 - Gerar Proposta")
        print("2 - Gerar Contrato")
        print("3 - Gerar Licitação")
        print("4 - Sair")

        opcao = input(
            "\nEscolha: "
        ).strip()

        try:

            if opcao == "1":

                gerar_proposta()

            elif opcao == "2":

                gerar_contrato()

            elif opcao == "3":

                gerar_licitacao()

            elif opcao == "4":

                print("\nSistema encerrado.")

                break

            else:

                print("\nOpção inválida.")

        except Exception as erro:

            print("\nERRO:")

            print(erro)

# =========================================================
# EXECUÇÃO
# =========================================================

if __name__ == "__main__":

    menu()