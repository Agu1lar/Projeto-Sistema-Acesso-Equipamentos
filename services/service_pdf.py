from __future__ import annotations

from pathlib import Path
from tempfile import NamedTemporaryFile

import os
import qrcode
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from utils.caminhos import caminho_recurso

LINK_EMPRESA = "https://acessoequipamentos.com.br/"
LOGO_ARQUIVOS = ("Logo.png", "assets/Logo.png")


def _localizar_logo() -> str | None:
    for arquivo in LOGO_ARQUIVOS:
        caminho = caminho_recurso(arquivo)
        if os.path.exists(caminho):
            return caminho
    return None


def _gerar_qr_temporario(link: str) -> str:
    with NamedTemporaryFile(delete=False, suffix=".png") as arquivo_temp:
        qrcode.make(link).save(arquivo_temp.name)
        return arquivo_temp.name


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
        logo = Image(caminho_logo, width=150, height=75)
        logo.hAlign = "CENTER"
        elementos.append(logo)

    elementos.append(Spacer(1, 16))
    elementos.append(Paragraph("Relatorio Fotografico", estilos["titulo"]))
    elementos.append(Spacer(1, 20))


def _adicionar_dados(elementos, dados, estilo_normal):
    for campo, valor in dados.items():
        texto = f"<font color='#8B1E1E'><b>{campo}:</b></font> {valor or '-'}"
        elementos.append(Paragraph(texto, estilo_normal))
        elementos.append(Spacer(1, 8))


def _adicionar_imagens(elementos, imagens):
    linhas = []
    linha_atual = []

    for imagem in imagens:
        caminho = Path(imagem)
        if not caminho.exists():
            continue

        try:
            flowable = Image(str(caminho), width=240, height=170)
            flowable.hAlign = "CENTER"
            linha_atual.append(flowable)
        except Exception:
            continue

        if len(linha_atual) == 2:
            linhas.append(linha_atual)
            linha_atual = []

    if linha_atual:
        while len(linha_atual) < 2:
            linha_atual.append("")
        linhas.append(linha_atual)

    if not linhas:
        return

    tabela = Table(linhas, colWidths=[260, 260], hAlign="CENTER")
    tabela.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    elementos.append(Spacer(1, 14))
    elementos.append(tabela)


def _adicionar_rodape(elementos, estilos, caminho_qr):
    img_qr = Image(caminho_qr, width=80, height=80)
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
    elementos.append(Spacer(1, 24))
    elementos.append(tabela_rodape)


def gerar_relatorio(dados, imagens, caminho_pdf):
    estilos = _montar_estilos()
    elementos = []
    caminho_qr = _gerar_qr_temporario(LINK_EMPRESA)

    try:
        _adicionar_cabecalho(elementos, estilos)
        _adicionar_dados(elementos, dados, estilos["normal"])
        _adicionar_imagens(elementos, imagens)
        _adicionar_rodape(elementos, estilos, caminho_qr)

        doc = SimpleDocTemplate(
            caminho_pdf,
            pagesize=A4,
            leftMargin=36,
            rightMargin=36,
            topMargin=36,
            bottomMargin=36,
        )
        doc.build(
            elementos,
            onFirstPage=desenhar_fundo,
            onLaterPages=desenhar_fundo,
        )
        return caminho_pdf
    finally:
        if os.path.exists(caminho_qr):
            os.remove(caminho_qr)
