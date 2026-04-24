from __future__ import annotations

import os
from pathlib import Path

from reportlab.graphics import renderPDF
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

from utils.caminhos import caminho_recurso

LINK_EMPRESA = "https://acessoequipamentos.com.br/"
LOGO_ARQUIVOS = ("Logo.png", "assets/Logo.png")
MODELO_CARTEIRINHA_VERSAO = "2026.03.24.5"


def _localizar_logo() -> str | None:
    for arquivo in LOGO_ARQUIVOS:
        caminho = caminho_recurso(arquivo)
        if os.path.exists(caminho):
            return caminho
    return None


def _texto(valor: str | None, fallback: str = "-") -> str:
    texto = str(valor or "").strip()
    return texto if texto else fallback


def _criar_qr_drawing(link: str, tamanho: float) -> Drawing:
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


def _ajustar_texto(pdf: canvas.Canvas, texto: str, fonte: str, tamanho: int, largura_max: float) -> str:
    valor = _texto(texto)
    if pdf.stringWidth(valor, fonte, tamanho) <= largura_max:
        return valor

    sufixo = "..."
    while valor and pdf.stringWidth(valor + sufixo, fonte, tamanho) > largura_max:
        valor = valor[:-1]
    return (valor + sufixo) if valor else sufixo


def _quebrar_linhas(
    pdf: canvas.Canvas,
    texto: str,
    fonte: str,
    tamanho: int,
    largura_max: float,
    max_linhas: int,
) -> list[str]:
    palavras = _texto(texto).split()
    if not palavras:
        return ["-"]

    linhas = []
    atual = palavras[0]

    for palavra in palavras[1:]:
        candidato = f"{atual} {palavra}"
        if pdf.stringWidth(candidato, fonte, tamanho) <= largura_max:
            atual = candidato
        else:
            linhas.append(atual)
            atual = palavra
            if len(linhas) == max_linhas - 1:
                break

    if len(linhas) < max_linhas:
        linhas.append(atual)

    restante = palavras[len(" ".join(linhas).split()):]
    if restante:
        linhas[-1] = _ajustar_texto(pdf, f"{linhas[-1]} {' '.join(restante)}", fonte, tamanho, largura_max)

    return linhas[:max_linhas]


def gerar_carteirinha_treinamento(dados: dict, caminho_pdf: str) -> str:
    arquivo = Path(caminho_pdf)
    arquivo.parent.mkdir(parents=True, exist_ok=True)

    pagina_w, pagina_h = landscape(A4)
    pdf = canvas.Canvas(str(arquivo), pagesize=landscape(A4))

    card_w = 170 * mm
    card_h = 100 * mm
    x = (pagina_w - card_w) / 2
    y = (pagina_h - card_h) / 2

    cor_primaria = colors.HexColor("#0F172A")
    cor_borda = colors.HexColor("#CBD5E1")
    cor_texto = colors.HexColor("#0F172A")
    cor_muted = colors.HexColor("#475569")
    cor_fundo = colors.HexColor("#F8FAFC")

    qr_drawing = _criar_qr_drawing(LINK_EMPRESA, 8 * mm)
    caminho_logo = _localizar_logo()

    pdf.setTitle(f"Carteirinha de Treinamento - {_texto(dados.get('NOME'))}")

    pdf.setFillColor(colors.white)
    pdf.rect(0, 0, pagina_w, pagina_h, stroke=0, fill=1)

    pdf.setFillColor(cor_fundo)
    pdf.roundRect(x, y, card_w, card_h, 7 * mm, stroke=0, fill=1)

    pdf.setStrokeColor(cor_borda)
    pdf.setLineWidth(1)
    pdf.roundRect(x, y, card_w, card_h, 7 * mm, stroke=1, fill=0)

    header_h = 18 * mm
    pdf.setFillColor(cor_primaria)
    pdf.roundRect(x, y + card_h - header_h, card_w, header_h, 7 * mm, stroke=0, fill=1)

    if caminho_logo:
        try:
            pdf.drawImage(
                caminho_logo,
                x + 8 * mm,
                y + card_h - 13.5 * mm,
                width=28 * mm,
                height=8 * mm,
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass

    pdf.setFillColor(colors.white)
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawRightString(x + card_w - 8 * mm, y + card_h - 7.2 * mm, "CARTEIRINHA DE TREINAMENTO")
    pdf.setFont("Helvetica", 6)
    pdf.drawRightString(x + card_w - 8 * mm, y + card_h - 11.2 * mm, "Documento de identificacao de capacitado")

    foto_w = 24 * mm
    foto_h = 30 * mm
    foto_x = x + card_w - foto_w - 10 * mm
    foto_y = y + 40 * mm
    pdf.setFillColor(colors.white)
    pdf.roundRect(foto_x, foto_y, foto_w, foto_h, 3 * mm, stroke=0, fill=1)
    pdf.setStrokeColor(colors.HexColor("#94A3B8"))
    pdf.setDash(4, 3)
    pdf.roundRect(foto_x, foto_y, foto_w, foto_h, 3 * mm, stroke=1, fill=0)
    pdf.setDash()
    pdf.setFillColor(cor_muted)
    pdf.setFont("Helvetica-Bold", 8)
    pdf.drawCentredString(foto_x + foto_w / 2, foto_y + foto_h / 2 + 3 * mm, "FOTO")
    pdf.setFont("Helvetica", 7)
    pdf.drawCentredString(foto_x + foto_w / 2, foto_y + foto_h / 2 - 1 * mm, "opcional")
    pdf.drawCentredString(foto_x + foto_w / 2, foto_y + foto_h / 2 - 5 * mm, "3x4")

    conteudo_x = x + 9 * mm
    largura_texto = 100 * mm

    pdf.setFillColor(cor_muted)
    pdf.setFont("Helvetica-Bold", 6)
    pdf.drawString(conteudo_x, y + 73 * mm, "COLABORADOR")
    pdf.setFillColor(cor_texto)
    pdf.setFont("Helvetica-Bold", 13)
    pdf.drawString(
        conteudo_x,
        y + 66 * mm,
        _ajustar_texto(pdf, _texto(dados.get("NOME")), "Helvetica-Bold", 13, largura_texto),
    )

    pdf.setFillColor(cor_muted)
    pdf.setFont("Helvetica-Bold", 6)
    pdf.drawString(conteudo_x, y + 56 * mm, "FUNCAO")
    pdf.setFillColor(cor_texto)
    pdf.setFont("Helvetica", 8.5)
    for indice, linha in enumerate(_quebrar_linhas(pdf, _texto(dados.get("FUNCAO")), "Helvetica", 8.5, largura_texto, 2)):
        pdf.drawString(conteudo_x, y + 50 * mm - (indice * 4.2 * mm), linha)

    pdf.setFillColor(cor_muted)
    pdf.setFont("Helvetica-Bold", 6)
    pdf.drawString(conteudo_x, y + 40 * mm, "TREINAMENTO")
    pdf.setFillColor(cor_texto)
    pdf.setFont("Helvetica-Bold", 9.5)
    linhas_treinamento = _quebrar_linhas(
        pdf, _texto(dados.get("TREINAMENTO")), "Helvetica-Bold", 9.5, largura_texto, 2
    )
    for indice, linha in enumerate(linhas_treinamento):
        pdf.drawString(conteudo_x, y + 34 * mm - (indice * 4.4 * mm), linha)

    bloco_y = y + 10 * mm
    bloco_h = 20 * mm
    bloco_x = x + 8 * mm
    bloco_w = card_w - 46 * mm
    pdf.setStrokeColor(colors.HexColor("#E2E8F0"))
    pdf.setLineWidth(0.8)
    pdf.roundRect(bloco_x, bloco_y, bloco_w, bloco_h, 3 * mm, stroke=1, fill=0)

    col1_x = bloco_x + 3 * mm
    col2_x = bloco_x + 55 * mm
    col3_x = bloco_x + 106 * mm
    linha1_y = bloco_y + bloco_h - 4 * mm
    linha2_y = bloco_y + bloco_h - 12 * mm

    def desenhar_campo(rotulo: str, valor: str, pos_x: float, pos_y: float, largura: float):
        pdf.setFillColor(cor_muted)
        pdf.setFont("Helvetica-Bold", 5.6)
        pdf.drawString(pos_x, pos_y, rotulo)
        pdf.setFillColor(cor_texto)
        pdf.setFont("Helvetica", 7.5)
        pdf.drawString(pos_x, pos_y - 3.2 * mm, _ajustar_texto(pdf, valor, "Helvetica", 7.5, largura))

    desenhar_campo("CPF", _texto(dados.get("CPF")), col1_x, linha1_y, 42 * mm)
    desenhar_campo("EMPRESA", _texto(dados.get("EMPRESA")), col2_x, linha1_y, 38 * mm)
    desenhar_campo("CARGA HORARIA", _texto(dados.get("CARGA_HORARIA")), col3_x, linha1_y, 24 * mm)
    desenhar_campo("INSTRUTOR", _texto(dados.get("INSTRUTOR")), col1_x, linha2_y, 42 * mm)
    desenhar_campo("EMISSAO", _texto(dados.get("DATA_EMISSAO")), col2_x, linha2_y, 38 * mm)
    desenhar_campo("VALIDADE", _texto(dados.get("VALIDADE")), col3_x, linha2_y, 24 * mm)

    responsavel = _texto(dados.get("RESPONSAVEL"))
    codigo = _texto(dados.get("CODIGO"))
    pdf.setFillColor(cor_texto)
    pdf.setFont("Helvetica-Bold", 5.8)
    pdf.drawString(
        x + 9 * mm,
        y + 4.6 * mm,
        _ajustar_texto(pdf, f"Responsavel: {responsavel}", "Helvetica-Bold", 5.8, 110 * mm),
    )
    pdf.setFont("Helvetica", 5.4)
    pdf.drawString(
        x + 9 * mm,
        y + 1.7 * mm,
        _ajustar_texto(pdf, f"Codigo: {codigo}", "Helvetica", 5.4, 110 * mm),
    )

    qr_size = 8 * mm
    renderPDF.draw(
        qr_drawing,
        pdf,
        x + card_w - 9 * mm - qr_size,
        y + 2 * mm,
    )

    pdf.save()
    return str(arquivo)
