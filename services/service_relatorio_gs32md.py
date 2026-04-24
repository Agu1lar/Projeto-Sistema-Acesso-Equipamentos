from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

from PIL import Image as PILImage, ImageOps
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from services.service_mobile_sync import FOTO_CAMPOS
from utils.caminhos import caminho_recurso

LOGO_ARQUIVOS = ("Logo.png", "assets/Logo.png")
FOTOS_POR_LINHA = 2
LARGURA_FOTO = 350
ALTURA_FOTO = 215
ALTURA_RODAPE = 36
ESPACO_RODAPE = 24
RODAPE_LINHAS = (
    "Praça Chui, 100 - Bairro João Pinheiro, Belo Horizonte - MG, CEP: 30.530.120",
    "31 3376-3377 // 31 99470-0201",
)


def _localizar_logo() -> str | None:
    for arquivo in LOGO_ARQUIVOS:
        caminho = caminho_recurso(arquivo)
        if Path(caminho).exists():
            return caminho
    return None


def _desenhar_rodape(canvas, doc) -> None:
    largura_pagina = doc.pagesize[0]
    y_linha = ALTURA_RODAPE + 10
    canvas.saveState()
    canvas.setStrokeColor(colors.HexColor("#cbd5e1"))
    canvas.setLineWidth(0.6)
    canvas.line(doc.leftMargin, y_linha, largura_pagina - doc.rightMargin, y_linha)
    canvas.setFillColor(colors.HexColor("#475569"))
    canvas.setFont("Helvetica", 8)
    canvas.drawCentredString(largura_pagina / 2, 22, RODAPE_LINHAS[0])
    canvas.drawCentredString(largura_pagina / 2, 12, RODAPE_LINHAS[1])
    canvas.restoreState()


def _montar_estilos():
    estilos = getSampleStyleSheet()
    return {
        "titulo": ParagraphStyle(
            name="TituloGS32MD",
            parent=estilos["Title"],
            fontSize=20,
            leading=24,
            textColor=colors.HexColor("#7f1d1d"),
            alignment=1,
        ),
        "cabecalho": ParagraphStyle(
            name="CabecalhoGS32MD",
            parent=estilos["BodyText"],
            fontSize=10,
            leading=13,
            textColor=colors.HexColor("#0f172a"),
        ),
        "horimetro": ParagraphStyle(
            name="HorimetroGS32MD",
            parent=estilos["BodyText"],
            fontSize=11,
            leading=14,
            textColor=colors.HexColor("#7f1d1d"),
        ),
        "legenda": ParagraphStyle(
            name="LegendaGS32MD",
            parent=estilos["BodyText"],
            fontSize=9,
            leading=11,
            textColor=colors.HexColor("#334155"),
            alignment=1,
        ),
        "vazio": ParagraphStyle(
            name="VazioGS32MD",
            parent=estilos["BodyText"],
            fontSize=10,
            leading=13,
            textColor=colors.HexColor("#64748b"),
            alignment=1,
        ),
    }


def _resolver_caminho_foto(payload: dict, foto: dict) -> str | None:
    caminho = str(foto.get("caminho", "")).strip()
    if not caminho:
        return None

    caminho_path = Path(caminho)
    if caminho_path.is_absolute() and caminho_path.exists():
        return str(caminho_path)

    base = Path(payload.get("pasta", ""))
    candidato = base / caminho
    if candidato.exists():
        return str(candidato)
    return None


def _normalizar_foto(origem: str, destino: Path) -> str:
    with PILImage.open(origem) as imagem_original:
        imagem = ImageOps.exif_transpose(imagem_original).convert("RGB")
        imagem.thumbnail((LARGURA_FOTO * 2, ALTURA_FOTO * 2), PILImage.Resampling.LANCZOS)
        imagem.save(destino, format="JPEG", quality=88, optimize=True)
    return str(destino)


def _coletar_fotos_ordenadas(payload: dict, item: dict) -> list[dict]:
    grupos = {grupo.get("campo"): grupo for grupo in item.get("fotos_nomeadas", [])}
    fotos = []
    indice_figura = 1

    for campo, titulo_padrao in FOTO_CAMPOS:
        grupo = grupos.get(campo)
        if not grupo:
            continue
        for foto in grupo.get("arquivos", []):
            caminho = _resolver_caminho_foto(payload, foto)
            if not caminho:
                continue
            fotos.append(
                {
                    "indice": indice_figura,
                    "titulo": grupo.get("titulo", titulo_padrao),
                    "caminho": caminho,
                }
            )
            indice_figura += 1

    campos_mapeados = {campo for campo, _titulo in FOTO_CAMPOS}
    for foto in item.get("fotos", []):
        campo = foto.get("campo")
        if campo in campos_mapeados:
            continue
        caminho = _resolver_caminho_foto(payload, foto)
        if not caminho:
            continue
        fotos.append(
            {
                "indice": indice_figura,
                "titulo": foto.get("titulo", "Foto complementar"),
                "caminho": caminho,
            }
        )
        indice_figura += 1

    return fotos


def _bloco_cabecalho(item: dict, estilos: dict):
    dados = item.get("dados", {})
    linhas = [
        ("Cliente", dados.get("cliente", "")),
        ("Obra", dados.get("obra", "")),
        ("Contrato", dados.get("contrato", "")),
        ("Equipamento", dados.get("equipamento", "")),
        ("Modelo", dados.get("modelo", "")),
        ("Serie", dados.get("serie", "")),
        ("Patrimonio", dados.get("patrimonio", "")),
        ("Data", dados.get("data_fim", "") or dados.get("data_inicio", "") or item.get("data", "")),
    ]

    tabela_dados = []
    for esquerda, direita in zip(linhas[::2], linhas[1::2]):
        tabela_dados.append(
            [
                Paragraph(f"<b>{esquerda[0]}:</b> {esquerda[1] or '-'}", estilos["cabecalho"]),
                Paragraph(f"<b>{direita[0]}:</b> {direita[1] or '-'}", estilos["cabecalho"]),
            ]
        )

    tabela = Table(tabela_dados, colWidths=[360, 360], hAlign="LEFT")
    tabela.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#fff7ed")),
                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#fdba74")),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#fed7aa")),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    return tabela


def _bloco_fotos(payload: dict, item: dict, pasta_temp: str, estilos: dict):
    fotos = _coletar_fotos_ordenadas(payload, item)
    if not fotos:
        return [Paragraph("Nenhuma foto foi recebida para este item.", estilos["vazio"])]

    linhas = []
    linha_atual = []
    for indice, foto in enumerate(fotos, start=1):
        destino = Path(pasta_temp) / f"foto_{indice:03d}.jpg"
        caminho_normalizado = _normalizar_foto(foto["caminho"], destino)
        imagem = Image(caminho_normalizado, width=LARGURA_FOTO, height=ALTURA_FOTO)
        imagem.hAlign = "CENTER"
        legenda = Paragraph(f"Figura {foto['indice']} / {foto['titulo']}", estilos["legenda"])
        linha_atual.append([legenda, Spacer(1, 4), imagem])

        if len(linha_atual) == FOTOS_POR_LINHA:
            linhas.append(linha_atual)
            linha_atual = []

    if linha_atual:
        while len(linha_atual) < FOTOS_POR_LINHA:
            linha_atual.append("")
        linhas.append(linha_atual)

    tabela = Table(linhas, colWidths=[365, 365], hAlign="CENTER")
    tabela.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ]
        )
    )
    return [tabela]


def sugerir_nome_relatorio_gs32md(payload: dict) -> str:
    itens = payload.get("itens", [])
    dados = itens[0].get("dados", {}) if itens else {}
    contrato = str(dados.get("contrato", "")).strip() or "sem_contrato"
    modelo = str(dados.get("modelo", "")).strip() or "GS32MD"
    return f"Relatorio_{modelo}_{contrato}.pdf".replace(" ", "_")


def gerar_relatorio_gs32md_pdf(payload: dict, caminho_pdf: str) -> str:
    itens = payload.get("itens", [])
    if not itens:
        raise ValueError("A pendencia selecionada nao possui itens para gerar o relatorio.")

    estilos = _montar_estilos()
    elementos = []
    logo = _localizar_logo()

    with TemporaryDirectory() as pasta_temp:
        for indice_item, item in enumerate(itens, start=1):
            dados = item.get("dados", {})
            if indice_item > 1:
                elementos.append(PageBreak())

            if logo:
                imagem_logo = Image(logo, width=145, height=72)
                imagem_logo.hAlign = "CENTER"
                elementos.append(imagem_logo)
                elementos.append(Spacer(1, 10))

            elementos.append(Paragraph("Relatório Fotográfico", estilos["titulo"]))
            elementos.append(Spacer(1, 10))
            elementos.append(_bloco_cabecalho(item, estilos))
            elementos.append(Spacer(1, 10))
            elementos.append(
                Paragraph(
                    f"<b>HORIMETRO:</b> {dados.get('horimetro_atual', '') or '-'}",
                    estilos["horimetro"],
                )
            )
            elementos.append(Spacer(1, 12))
            elementos.extend(_bloco_fotos(payload, item, pasta_temp, estilos))

        doc = SimpleDocTemplate(
            caminho_pdf,
            pagesize=landscape(A4),
            leftMargin=24,
            rightMargin=24,
            topMargin=20,
            bottomMargin=ALTURA_RODAPE + ESPACO_RODAPE,
        )
        doc.build(
            elementos,
            onFirstPage=lambda canvas, doc: _desenhar_rodape(canvas, doc),
            onLaterPages=lambda canvas, doc: _desenhar_rodape(canvas, doc),
        )

    return caminho_pdf
