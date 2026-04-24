from __future__ import annotations

from pathlib import Path

from pypdf import PdfReader


MAPA_ROTULOS = {
    "COLABORADOR": "NOME",
    "FUNCAO": "FUNCAO",
    "TREINAMENTO": "TREINAMENTO",
    "CPF": "CPF",
    "EMPRESA": "EMPRESA",
    "CARGA HORARIA": "CARGA_HORARIA",
    "INSTRUTOR": "INSTRUTOR",
    "EMISSAO": "DATA_EMISSAO",
    "VALIDADE": "VALIDADE",
}


def _texto_pdf(caminho_pdf: str) -> str:
    leitor = PdfReader(caminho_pdf)
    partes = []
    for pagina in leitor.pages:
        partes.append(pagina.extract_text() or "")
    return "\n".join(partes)


def _normalizar_linhas(texto: str) -> list[str]:
    return [linha.strip() for linha in texto.splitlines() if linha and linha.strip()]


def _valor_apos_rotulo(linhas: list[str], indice: int) -> str:
    if indice + 1 >= len(linhas):
        return ""
    return linhas[indice + 1].strip()


def _extrair_codigo_responsavel(linhas: list[str]) -> tuple[str, str]:
    codigo = ""
    responsavel = ""
    for linha in linhas:
        if linha.upper().startswith("RESPONSAVEL:"):
            responsavel = linha.split(":", 1)[1].strip()
        if linha.upper().startswith("CODIGO:"):
            codigo = linha.split(":", 1)[1].strip()
    return codigo, responsavel


def extrair_dados_carteirinha_pdf(caminho_pdf: str) -> dict:
    caminho = Path(caminho_pdf)
    texto = _texto_pdf(str(caminho))
    linhas = _normalizar_linhas(texto)
    dados = {
        "CODIGO": "",
        "NOME": "",
        "CPF": "",
        "EMPRESA": "",
        "FUNCAO": "",
        "TREINAMENTO": "",
        "CARGA_HORARIA": "",
        "INSTRUTOR": "",
        "DATA_EMISSAO": "",
        "VALIDADE": "",
        "RESPONSAVEL": "",
        "OBS": "",
        "CERTIFICADO_RUA": "",
        "CERTIFICADO_NUMERO": "",
        "CERTIFICADO_BAIRRO": "",
        "CERTIFICADO_CIDADE": "",
        "CERTIFICADO_UF": "",
        "CERTIFICADO_CEP": "",
        "CERTIFICADO_IMPRIMIR_CPF": "SIM",
        "CERTIFICADO_MODELO": "TESOURA",
        "DOCUMENTOS_PASTA": str(caminho.parent),
        "CERTIFICADO_PDF_CAMINHO": "",
        "CERTIFICADO_WORD_CAMINHO": "",
        "PDF_CAMINHO": str(caminho),
    }

    indice = 0
    while indice < len(linhas):
        linha = linhas[indice]
        chave = MAPA_ROTULOS.get(linha.upper())
        if chave:
            dados[chave] = _valor_apos_rotulo(linhas, indice)
            indice += 2
            continue
        indice += 1

    codigo, responsavel = _extrair_codigo_responsavel(linhas)
    if codigo:
        dados["CODIGO"] = codigo
    if responsavel:
        dados["RESPONSAVEL"] = responsavel

    # Estrutura esperada: raiz/empresa/colaborador/arquivo.pdf
    if not dados["EMPRESA"] and caminho.parent.parent != caminho.parent:
        dados["EMPRESA"] = caminho.parent.parent.name
    if not dados["NOME"]:
        dados["NOME"] = caminho.parent.name
    if not dados["CODIGO"]:
        dados["CODIGO"] = caminho.stem

    return dados


def listar_pdfs_carteirinhas(pasta_raiz: str) -> list[str]:
    raiz = Path(pasta_raiz)
    if not raiz.exists():
        return []
    return [str(arquivo) for arquivo in raiz.rglob("*.pdf") if arquivo.is_file()]
