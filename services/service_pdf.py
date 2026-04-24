from __future__ import annotations

import os
import shutil
from pathlib import Path
from tempfile import TemporaryDirectory

from PIL import Image as PILImage, ImageOps
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from utils.caminhos import caminho_recurso

LINK_EMPRESA = "https://acessoequipamentos.com.br/"
LOGO_ARQUIVOS = ("Logo.png", "assets/Logo.png")
TAMANHO_MAXIMO_PDF = 5 * 1024 * 1024
PERFIS_COMPRESSAO = [
    (1600, 55),
    (1280, 45),
    (1024, 35),
    (900, 28),
    (768, 22),
    (640, 18),
]


def _localizar_logo() -> str | None:
    for arquivo in LOGO_ARQUIVOS:
        caminho = caminho_recurso(arquivo)
        if os.path.exists(caminho):
            return caminho
    return None


def _criar_qr_flowable(link: str, tamanho: float) -> Drawing:
    widget = qr.QrCodeWidget(link)
    x1, y1, x2, y2 = widget.getBounds()
    largura = max(x2 - x1, 1)
    altura = max(y2 - y1, 1)
    desenho = Drawing(
        tamanho,
        tamanho,
        transform=[tamanho / largura, 0, 0, tamanho / altura, 0, 0],
    )
    desenho.add(widget)
    return desenho


def _humanizar_titulo(texto: str) -> str:
    valor = str(texto or "").strip()
    if not valor:
        return "Imagem"
    return valor.replace("_", " ").replace("-", " ").strip().title()


def _normalizar_item_imagem(item) -> dict | None:
    if isinstance(item, str):
        caminho = item
        titulo = _humanizar_titulo(Path(item).stem)
    elif isinstance(item, dict):
        caminho = item.get("caminho") or item.get("path") or item.get("arquivo")
        titulo = item.get("titulo") or item.get("nome") or _humanizar_titulo(Path(str(caminho or "")).stem)
    else:
        return None

    if not caminho:
        return None

    return {"caminho": str(caminho), "titulo": str(titulo or "Imagem").strip()}


def desenhar_fundo(canvas_obj, _doc):
    caminho_logo = _localizar_logo()
    if not caminho_logo:
        return

    try:
        canvas_obj.saveState()
        canvas_obj.setFillAlpha(0.08)
        canvas_obj.drawImage(
            caminho_logo,
            110,
            180,
            width=380,
            height=380,
            preserveAspectRatio=True,
            mask="auto",
        )
        canvas_obj.restoreState()
    except Exception:
        pass


def _montar_estilos():
    styles = getSampleStyleSheet()
    return {
        "normal": styles["Normal"],
        "titulo": ParagraphStyle(
            name="TituloCentral",
            parent=styles["Title"],
            alignment=TA_CENTER,
            textColor=colors.HexColor("#8B1E1E"),
        ),
        "rodape": ParagraphStyle(
            name="Rodape",
            parent=styles["Normal"],
            alignment=TA_CENTER,
            textColor=colors.grey,
            fontSize=9,
        ),
    }


def _adicionar_cabecalho(elementos, estilos):
    caminho_logo = _localizar_logo()
    if caminho_logo:
        logo = Image(caminho_logo, width=165, height=82)
        logo.hAlign = "CENTER"
        elementos.append(logo)

    elementos.append(Spacer(1, 12))
    elementos.append(Paragraph("Relatorio Fotografico", estilos["titulo"]))
    elementos.append(Spacer(1, 16))


def _adicionar_dados(elementos, dados, estilo_normal):
    for campo, valor in dados.items():
        texto = f"<font color='#8B1E1E'><b>{campo}:</b></font> {valor or '-'}"
        elementos.append(Paragraph(texto, estilo_normal))
        elementos.append(Spacer(1, 6))


def _flowable_imagem(item_imagem: dict):
    caminho = Path(item_imagem["caminho"])
    if not caminho.exists():
        return None

    try:
        flowable = Image(str(caminho), width=248, height=186)
        flowable.hAlign = "CENTER"
        return flowable
    except Exception:
        return None


def _adicionar_imagens(elementos, imagens, estilos):
    linhas = []
    linha_atual = []

    for item in imagens:
        item_imagem = _normalizar_item_imagem(item)
        if not item_imagem:
            continue

        bloco = _flowable_imagem(item_imagem)
        if not bloco:
            continue
        linha_atual.append(bloco)

        if len(linha_atual) == 2:
            linhas.append(linha_atual)
            linha_atual = []

    if linha_atual:
        while len(linha_atual) < 2:
            linha_atual.append("")
        linhas.append(linha_atual)

    if not linhas:
        return

    tabela = Table(linhas, colWidths=[250, 250], hAlign="CENTER")
    tabela.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    elementos.append(Spacer(1, 10))
    elementos.append(tabela)


def _comprimir_imagem_windows(caminho_origem: str, caminho_destino: str, max_lado: int, qualidade: int) -> bool:
    try:
        with PILImage.open(caminho_origem) as imagem_original:
            # Respeita a orientacao original gravada no EXIF antes de redimensionar.
            imagem = ImageOps.exif_transpose(imagem_original)
            imagem = imagem.convert("RGB")

            largura, altura = imagem.size
            maior_lado = max(largura, altura) or 1
            escala = min(1.0, max_lado / maior_lado)
            novo_tamanho = (
                max(1, int(largura * escala)),
                max(1, int(altura * escala)),
            )

            if novo_tamanho != imagem.size:
                imagem = imagem.resize(novo_tamanho, PILImage.Resampling.LANCZOS)

            imagem.save(
                caminho_destino,
                format="JPEG",
                quality=qualidade,
                optimize=True,
            )
        return os.path.exists(caminho_destino)
    except Exception:
        return False


def _preparar_imagens_para_pdf(imagens, pasta_temp: str, max_lado: int, qualidade: int) -> list[dict]:
    resultado = []
    for indice, item in enumerate(imagens):
        item_imagem = _normalizar_item_imagem(item)
        if not item_imagem:
            continue

        caminho = Path(item_imagem["caminho"])
        if not caminho.exists():
            continue

        destino = Path(pasta_temp) / f"img_{indice:03d}.jpg"
        if _comprimir_imagem_windows(str(caminho), str(destino), max_lado=max_lado, qualidade=qualidade):
            resultado.append({"caminho": str(destino), "titulo": item_imagem["titulo"]})
        else:
            resultado.append(item_imagem)
    return resultado


def _adicionar_rodape(elementos, estilos):
    img_qr = _criar_qr_flowable(LINK_EMPRESA, 80)
    tabela_rodape = Table(
        [[img_qr, Paragraph(LINK_EMPRESA, estilos["rodape"])]],
        colWidths=[90, 390],
    )
    tabela_rodape.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (0, 0), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    elementos.append(Spacer(1, 18))
    elementos.append(tabela_rodape)


def _gerar_relatorio_bruto(dados, imagens, caminho_pdf):
    estilos = _montar_estilos()
    elementos = []

    _adicionar_cabecalho(elementos, estilos)
    _adicionar_dados(elementos, dados, estilos["normal"])
    _adicionar_imagens(elementos, imagens, estilos)
    _adicionar_rodape(elementos, estilos)

    doc = SimpleDocTemplate(
        caminho_pdf,
        pagesize=A4,
        leftMargin=28,
        rightMargin=28,
        topMargin=28,
        bottomMargin=28,
    )
    doc.build(
        elementos,
        onFirstPage=desenhar_fundo,
        onLaterPages=desenhar_fundo,
    )
    return caminho_pdf


def gerar_relatorio(dados, imagens, caminho_pdf):
    caminho_destino = Path(caminho_pdf)
    caminho_destino.parent.mkdir(parents=True, exist_ok=True)

    with TemporaryDirectory() as pasta_temp:
        for max_lado, qualidade in PERFIS_COMPRESSAO:
            imagens_preparadas = _preparar_imagens_para_pdf(imagens, pasta_temp, max_lado=max_lado, qualidade=qualidade)
            caminho_teste = Path(pasta_temp) / "relatorio_comprimido.pdf"
            if caminho_teste.exists():
                caminho_teste.unlink()

            _gerar_relatorio_bruto(dados, imagens_preparadas, str(caminho_teste))
            if caminho_teste.exists() and caminho_teste.stat().st_size <= TAMANHO_MAXIMO_PDF:
                shutil.copyfile(caminho_teste, caminho_destino)
                return str(caminho_destino)

        raise ValueError(
            "Nao foi possivel gerar o relatorio com ate 5 MB, mesmo apos compressao maxima das imagens."
        )
