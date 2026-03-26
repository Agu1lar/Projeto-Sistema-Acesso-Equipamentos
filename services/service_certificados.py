from __future__ import annotations

import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import Image, LongTable, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from utils.caminhos import caminho_base, caminho_recurso


STATUS_EM_DIA = "EM DIA"
STATUS_PENDENTE = "PENDENTE RENOVACAO"
STATUS_VENCIDO = "VENCIDO"
STATUS_DATA_INVALIDA = "DATA INVALIDA"
LOGO_ARQUIVOS = ("Logo.png", "assets/Logo.png")
COLUNAS_RELATORIO = [
    "TIPO_CERTIFICADO",
    "NOME",
    "CPF",
    "CARGA_HORARIA",
    "DATA_EMISSAO",
    "DATA_VENCIMENTO",
    "STATUS",
    "DIAS_RESTANTES",
]
COR_TITULO = colors.HexColor("#8b1e1e")
COR_CABECALHO = colors.HexColor("#8b1e1e")
COR_LINHA_1 = colors.HexColor("#f8fafc")
COR_LINHA_2 = colors.HexColor("#eef2f7")
COR_BORDA = colors.HexColor("#cbd5e1")
COR_TEXTO = colors.HexColor("#0f172a")
COR_PENDENTE_BG = colors.HexColor("#fff7ed")
COR_PENDENTE_TX = colors.HexColor("#c2410c")
COR_VENCIDO_BG = colors.HexColor("#fef2f2")
COR_VENCIDO_TX = colors.HexColor("#b91c1c")


def _texto(valor: str | None, fallback: str = "-") -> str:
    texto = str(valor or "").strip()
    return texto if texto else fallback


def _bool_sim(valor) -> bool:
    return str(valor or "").strip().upper() in {"1", "SIM", "S", "TRUE", "YES", "X"}


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
    palavras = str(texto or "").strip().split()
    if not palavras:
        return ["-"]

    linhas: list[str] = []
    atual = palavras[0]

    for palavra in palavras[1:]:
        candidato = f"{atual} {palavra}"
        if pdf.stringWidth(candidato, fonte, tamanho) <= largura_max:
            atual = candidato
            continue

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


def _desenhar_marca_dagua(pdf: canvas.Canvas, logo: str | None, pagina_w: float, pagina_h: float):
    if not logo:
        return
    try:
        pdf.saveState()
        if hasattr(pdf, "setFillAlpha"):
            pdf.setFillAlpha(0.08)
        pdf.drawImage(
            logo,
            (pagina_w - 120 * mm) / 2,
            (pagina_h - 120 * mm) / 2 - 4 * mm,
            width=120 * mm,
            height=120 * mm,
            preserveAspectRatio=True,
            mask="auto",
        )
        pdf.restoreState()
    except Exception:
        pass


def _localizar_modelo_certificado() -> str | None:
    for configuracao in MODELOS_CERTIFICADO.values():
        candidatos = [
            Path(caminho_base()) / configuracao["arquivo"],
            Path.cwd() / configuracao["arquivo"],
        ]
        for candidato in candidatos:
            if candidato.exists():
                return str(candidato)
    for pasta in {Path(caminho_base()), Path.cwd()}:
        for arquivo in pasta.glob("*CERTIFICADO*.docx"):
            if arquivo.is_file():
                return str(arquivo)
    return None


def _obter_modelo_certificado(chave_modelo: str | None) -> tuple[str, dict]:
    chave = str(chave_modelo or "").strip().upper()
    if chave not in MODELOS_CERTIFICADO:
        chave = "TESOURA"

    configuracao = MODELOS_CERTIFICADO[chave]
    candidatos = [
        Path(caminho_base()) / configuracao["arquivo"],
        Path.cwd() / configuracao["arquivo"],
    ]
    for candidato in candidatos:
        if candidato.exists():
            return str(candidato), configuracao
    raise FileNotFoundError(f"Modelo de certificado Word nao encontrado para {chave}.")


def _formatar_data_extenso(valor: str) -> str:
    meses = {
        1: "Janeiro",
        2: "Fevereiro",
        3: "Março",
        4: "Abril",
        5: "Maio",
        6: "Junho",
        7: "Julho",
        8: "Agosto",
        9: "Setembro",
        10: "Outubro",
        11: "Novembro",
        12: "Dezembro",
    }
    data = _parse_data_br(valor)
    if not data:
        return f"Belo Horizonte, {_texto(valor)}"
    return f"Belo Horizonte, {data.day:02d} de {meses[data.month]} {data.year}"


def _distribuir_texto(texto: str, comprimentos: list[int]) -> list[str]:
    if not comprimentos:
        return []
    partes = []
    cursor = 0
    for indice, comprimento in enumerate(comprimentos):
        if indice == len(comprimentos) - 1:
            partes.append(texto[cursor:])
        else:
            partes.append(texto[cursor:cursor + comprimento])
            cursor += comprimento
    return partes


def _substituir_textos_iguais(textos: list[ET.Element], alvo: str, novo: str):
    for node in textos:
        if node.text == alvo:
            node.text = novo


def _substituir_intervalos_por_texto(textos: list[ET.Element], inicio: str, fim: str, novo_texto: str):
    total = len(textos)
    indice = 0
    while indice < total:
        if textos[indice].text != inicio:
            indice += 1
            continue

        fim_indice = None
        for candidato in range(indice, total):
            if textos[candidato].text == fim:
                fim_indice = candidato
                break
        if fim_indice is None:
            break

        faixa = textos[indice:fim_indice + 1]
        comprimentos = [len(node.text or "") for node in faixa]
        partes = _distribuir_texto(novo_texto, comprimentos)
        for node, parte in zip(faixa, partes):
            node.text = parte
        indice = fim_indice + 1


MODELOS_CERTIFICADO = {
    "TESOURA": {
        "arquivo": "2.CERTIFICADO PADRÃO TESOURA ATUALIZADO.docx",
        "nome_base": "HUDSON RAFAEL DOS SANTOS",
        "cpf_base": "091.828.396-54",
        "endereco_base": "PRAÇA CHUI, 100, JOÃO PINHEIRO – BELO HORIZONTE/MG, 30530-120",
        "instrutor_base": "JULIO MOREIRA RODRIGUES",
        "responsavel_base": "Flaviano Silveira Queiroz",
        "data_extenso_base": "Belo Horizonte, 24 de Março 2026",
        "inicio_intro": "concluiu com êxito o treinamento em Plataforma Elevatória Móvel de Trabalho (PEMT), g",
        "fim_intro": "PRAÇA CHUI, 100, JOÃO PINHEIRO – BELO HORIZONTE/MG, 30530-120",
        "prefixo_intro": "concluiu com êxito o treinamento em ",
    },
    "ARTICULADA": {
        "arquivo": "2.CERTIFICADO PADRÃO ARTICULADA ATUALIZADO(1).docx",
        "nome_base": "Lucas Andrade De Laia",
        "cpf_base": "108.251.236-23",
        "endereco_base": "Av. Assis Chateaubriand, 889 - Floresta, Belo Horizonte – MG - 30150-101,",
        "instrutor_base": "Ederson Diniz Carneiro",
        "responsavel_base": "Flaviano Silveira Queiroz",
        "data_extenso_base": "Belo Horizonte, 03 de Março 2026",
        "inicio_intro": " concluiu com êxito o treinamento em Plataforma Elevatória M",
        "fim_intro": "30150-101,",
        "prefixo_intro": " concluiu com êxito o treinamento em ",
    },
}
NS_W = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def gerar_certificado_treinamento(dados: dict, caminho_pdf: str) -> str:
    arquivo = Path(caminho_pdf)
    arquivo.parent.mkdir(parents=True, exist_ok=True)

    pagina_w, pagina_h = landscape(A4)
    pdf = canvas.Canvas(str(arquivo), pagesize=landscape(A4))
    logo = _localizar_logo()

    nome = _texto(dados.get("NOME"))
    cpf = _texto(dados.get("CPF"))
    treinamento = _texto(dados.get("TREINAMENTO"))
    carga_horaria = _texto(dados.get("CARGA_HORARIA"))
    data_realizacao = _texto(dados.get("DATA_EMISSAO"))
    validade = str(dados.get("VALIDADE", "") or "").strip()
    rua = _texto(dados.get("CERTIFICADO_RUA"))
    numero = _texto(dados.get("CERTIFICADO_NUMERO"))
    bairro = _texto(dados.get("CERTIFICADO_BAIRRO"))
    cep = _texto(dados.get("CERTIFICADO_CEP"))
    instrutor = _texto(dados.get("INSTRUTOR"))
    responsavel = _texto(dados.get("RESPONSAVEL"))
    imprimir_cpf = _bool_sim(dados.get("CERTIFICADO_IMPRIMIR_CPF", "SIM"))

    pdf.setTitle(f"Certificado de Treinamento - {nome}")
    pdf.setFillColor(colors.white)
    pdf.rect(0, 0, pagina_w, pagina_h, stroke=0, fill=1)

    pdf.setFillColor(colors.HexColor("#b30000"))
    pdf.wedge(-8 * mm, pagina_h - 56 * mm, 56 * mm, pagina_h + 8 * mm, 180, 270, stroke=0, fill=1)
    pdf.wedge(pagina_w - 56 * mm, -8 * mm, pagina_w + 8 * mm, 56 * mm, 0, 90, stroke=0, fill=1)

    _desenhar_marca_dagua(pdf, logo, pagina_w, pagina_h)

    if logo:
        try:
            pdf.drawImage(
                logo,
                pagina_w - 56 * mm,
                pagina_h - 22 * mm,
                width=48 * mm,
                height=16 * mm,
                preserveAspectRatio=True,
                mask="auto",
            )
        except Exception:
            pass

    pdf.setFillColor(colors.black)
    pdf.setFont("Times-Bold", 22)
    pdf.drawCentredString(pagina_w / 2, pagina_h - 18 * mm, "CERTIFICADO DE TREINAMENTO")

    largura_texto = pagina_w - 34 * mm
    inicio_x = 16 * mm
    linha_y = pagina_h - 34 * mm

    parte_cpf = f", portador do CPF {cpf}," if imprimir_cpf and cpf != "-" else ","
    endereco = f"{rua}, {numero} - {bairro}, Belo Horizonte/MG - {cep}"
    texto_abertura = (
        f"{nome}{parte_cpf} concluiu com exito o treinamento em {treinamento}, "
        f"com carga horaria de {carga_horaria}, realizado no dia {data_realizacao}, {endereco}."
    )

    pdf.setFont("Helvetica", 12)
    for indice, linha in enumerate(_quebrar_linhas(pdf, texto_abertura, "Helvetica", 12, largura_texto, 4)):
        pdf.drawString(inicio_x, linha_y - (indice * 7.2 * mm), linha)

    conteudo_y = linha_y - 38 * mm
    pdf.setFont("Times-Roman", 14)
    pdf.drawString(inicio_x + 10 * mm, conteudo_y, "Conteudo do Treinamento:")
    pdf.setFont("Times-Roman", 12.5)
    pdf.drawString(
        inicio_x + 10 * mm,
        conteudo_y - 7 * mm,
        "O treinamento, em conformidade com a Norma Regulamentadora (NR-18), abordou os seguintes topicos:",
    )

    topicos = [
        ("Apresentacao da PEMT", "Identificacao das pecas, componentes e funcionamento do equipamento."),
        ("Procedimentos de Seguranca", "Uso correto de Equipamentos de Protecao Individual (EPIs), inspecao pre-operacional e check list de seguranca."),
        ("Operacao Segura", "Manobras, posicionamento da plataforma, limites de carga e altura, e procedimentos em caso de emergencia."),
        ("Manutencao", "Verificacao diaria das condicoes da maquina, incluindo baterias e sistemas eletricos e hidraulicos."),
        ("Condicoes de Trabalho", "Analise dos riscos no local de trabalho, isolamento da area de operacao e interacoes com outros equipamentos."),
    ]

    y_topico = conteudo_y - 18 * mm
    largura_topico = pagina_w - 44 * mm
    for titulo, descricao in topicos:
        linhas = _quebrar_linhas(pdf, f"{titulo}: {descricao}", "Helvetica", 11.5, largura_topico, 3)
        for indice_linha, linha in enumerate(linhas):
            if indice_linha == 0:
                marcador = "- "
                negrito = f"{titulo}:"
                resto = linha[len(f"{titulo}:"):].lstrip()
                pdf.setFont("Helvetica", 11.5)
                pdf.drawString(inicio_x + 8 * mm, y_topico, marcador)
                pdf.setFont("Helvetica-Bold", 11.5)
                pdf.drawString(inicio_x + 12 * mm, y_topico, negrito)
                pdf.setFont("Helvetica", 11.5)
                pdf.drawString(
                    inicio_x + 12 * mm + pdf.stringWidth(negrito + " ", "Helvetica-Bold", 11.5),
                    y_topico,
                    resto,
                )
            else:
                pdf.setFont("Helvetica", 11.5)
                pdf.drawString(inicio_x + 12 * mm, y_topico, linha)
            y_topico -= 6.8 * mm
        y_topico -= 1.5 * mm

    cidade_data = f"Belo Horizonte, {data_realizacao}"
    pdf.setFont("Helvetica-Bold", 12.5)
    pdf.drawCentredString(pagina_w / 2, 50 * mm, cidade_data)

    assinaturas_y = 26 * mm
    linha_w = 62 * mm
    posicoes = [26 * mm, (pagina_w - linha_w) / 2, pagina_w - 26 * mm - linha_w]
    textos = [
        [responsavel, "Coordenador de Manutencao e", "Operacional"],
        [instrutor, "Instrutor Tecnico de Treinamento"],
        [nome, "Operador"],
    ]

    for pos_x, linhas in zip(posicoes, textos):
        pdf.line(pos_x, assinaturas_y + 10 * mm, pos_x + linha_w, assinaturas_y + 10 * mm)
        cursor_y = assinaturas_y + 6.8 * mm
        for linha in linhas:
            pdf.setFont("Helvetica-Bold", 10 if linha == linhas[0] else 9.2)
            pdf.drawCentredString(pos_x + (linha_w / 2), cursor_y, _ajustar_texto(pdf, linha, "Helvetica-Bold", 10, linha_w))
            cursor_y -= 5.4 * mm

    rodape_validade = "VALIDADE DE 24 MESES"
    if validade:
        rodape_validade = f"VALIDADE ATE {validade}"
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawRightString(pagina_w - 18 * mm, 14 * mm, rodape_validade)

    pdf.save()
    return str(arquivo)


def gerar_certificado_word(dados: dict, caminho_arquivo: str) -> str:
    arquivo = Path(caminho_arquivo)
    arquivo.parent.mkdir(parents=True, exist_ok=True)
    modelo, configuracao_modelo = _obter_modelo_certificado(dados.get("CERTIFICADO_MODELO"))

    nome = _texto(dados.get("NOME")).upper()
    cpf = _texto(dados.get("CPF"))
    treinamento = _texto(dados.get("TREINAMENTO"))
    carga_horaria = _texto(dados.get("CARGA_HORARIA"))
    data_realizacao = _texto(dados.get("DATA_EMISSAO"))
    rua = _texto(dados.get("CERTIFICADO_RUA"))
    numero = _texto(dados.get("CERTIFICADO_NUMERO"))
    bairro = _texto(dados.get("CERTIFICADO_BAIRRO"))
    cep = _texto(dados.get("CERTIFICADO_CEP"))
    instrutor = _texto(dados.get("INSTRUTOR")).upper()
    responsavel = _texto(dados.get("RESPONSAVEL"))
    imprimir_cpf = _bool_sim(dados.get("CERTIFICADO_IMPRIMIR_CPF", "SIM"))

    endereco = f"{rua}, {numero}, {bairro} – BELO HORIZONTE/MG, {cep}".upper()
    intro = (
        f"{configuracao_modelo['prefixo_intro']}{treinamento}, "
        f"com carga horária de {carga_horaria}, realizado no dia {data_realizacao}, {endereco}"
    )

    shutil.copyfile(modelo, arquivo)

    with zipfile.ZipFile(arquivo, "r") as docx:
        conteudos = {info.filename: docx.read(info.filename) for info in docx.infolist()}
        raiz = ET.fromstring(conteudos["word/document.xml"])
        textos = raiz.findall(".//w:t", NS_W)

        _substituir_textos_iguais(textos, configuracao_modelo["nome_base"], nome)
        _substituir_textos_iguais(textos, configuracao_modelo["cpf_base"], cpf if imprimir_cpf else "")
        _substituir_textos_iguais(
            textos,
            configuracao_modelo["endereco_base"],
            endereco,
        )
        _substituir_textos_iguais(textos, configuracao_modelo["instrutor_base"], instrutor)
        _substituir_textos_iguais(textos, configuracao_modelo["responsavel_base"], responsavel)
        _substituir_textos_iguais(textos, configuracao_modelo["data_extenso_base"], _formatar_data_extenso(data_realizacao))

        _substituir_intervalos_por_texto(
            textos,
            configuracao_modelo["inicio_intro"],
            configuracao_modelo["fim_intro"],
            intro,
        )

        if not imprimir_cpf:
            for alvo in ("portador do CP", "F", "portador do ", "CPF "):
                _substituir_textos_iguais(textos, alvo, "")

        conteudos["word/document.xml"] = ET.tostring(raiz, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(arquivo, "w", compression=zipfile.ZIP_DEFLATED) as docx:
        for nome_arquivo, dados_arquivo in conteudos.items():
            docx.writestr(nome_arquivo, dados_arquivo)

    return str(arquivo)


def _parse_data_br(valor: str):
    texto = str(valor or "").strip()
    if not texto:
        return None
    try:
        return datetime.strptime(texto, "%d/%m/%Y")
    except ValueError:
        return None


def calcular_status_certificado(data_vencimento: str, referencia: datetime | None = None) -> tuple[str, str]:
    vencimento = _parse_data_br(data_vencimento)
    if vencimento is None:
        return STATUS_DATA_INVALIDA, ""

    base = referencia or datetime.now()
    dias_restantes = (vencimento.date() - base.date()).days
    if dias_restantes < 0:
        return STATUS_VENCIDO, str(dias_restantes)
    if dias_restantes <= 30:
        return STATUS_PENDENTE, str(dias_restantes)
    return STATUS_EM_DIA, str(dias_restantes)


def preparar_certificado(dados: dict, referencia: datetime | None = None) -> dict:
    registro = {chave: str(valor or "").strip() for chave, valor in dados.items()}
    status, dias_restantes = calcular_status_certificado(registro.get("DATA_VENCIMENTO", ""), referencia=referencia)
    registro["STATUS"] = status
    registro["DIAS_RESTANTES"] = dias_restantes
    registro.setdefault("DATA_CADASTRO", datetime.now().strftime("%d/%m/%Y %H:%M"))
    return registro


def atualizar_status_certificados_df(df: pd.DataFrame, referencia: datetime | None = None) -> pd.DataFrame:
    base = df.fillna("").copy()
    if base.empty:
        return base

    for coluna in ("STATUS", "DIAS_RESTANTES", "DATA_CADASTRO"):
        if coluna not in base.columns:
            base[coluna] = ""

    for indice in base.index:
        atualizado = preparar_certificado(base.loc[indice].to_dict(), referencia=referencia)
        for chave in ("STATUS", "DIAS_RESTANTES"):
            base.loc[indice, chave] = atualizado.get(chave, "")
        if not str(base.loc[indice, "DATA_CADASTRO"]).strip():
            base.loc[indice, "DATA_CADASTRO"] = atualizado.get("DATA_CADASTRO", "")
    return base


def listar_certificados_pendentes(banco) -> pd.DataFrame:
    df = atualizar_status_certificados_df(banco.carregar_dataframe("CERTIFICADOS"))
    return df[df["STATUS"].isin([STATUS_PENDENTE, STATUS_VENCIDO, STATUS_DATA_INVALIDA])].copy()


def _localizar_logo() -> str | None:
    for arquivo in LOGO_ARQUIVOS:
        caminho = caminho_recurso(arquivo)
        if Path(caminho).exists():
            return caminho
    return None


def gerar_relatorio_certificados(df: pd.DataFrame, caminho_pdf: str, titulo: str = "Relatorio de Certificados") -> str:
    base = atualizar_status_certificados_df(df.fillna("").copy())
    estilos = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        name="CertificadosTitulo",
        parent=estilos["Title"],
        textColor=COR_TITULO,
        alignment=1,
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
    )
    estilo_info = ParagraphStyle(
        name="CertificadosInfo",
        parent=estilos["BodyText"],
        fontSize=8,
        leading=10,
        textColor=COR_TEXTO,
        wordWrap="CJK",
    )
    estilo_cabecalho = ParagraphStyle(
        name="CertificadosCabecalho",
        parent=estilo_info,
        textColor=colors.white,
        alignment=1,
        fontName="Helvetica-Bold",
    )

    elementos = []
    logo = _localizar_logo()
    if logo:
        elementos.append(Image(logo, width=130, height=58))
    elementos.append(Spacer(1, 8))
    elementos.append(Paragraph(titulo, estilo_titulo))
    elementos.append(Spacer(1, 6))
    elementos.append(
        Paragraph(
            f"Gerado em {datetime.now():%d/%m/%Y %H:%M} | Total: {len(base)} | "
            f"Pendentes: {len(base[base['STATUS'] == STATUS_PENDENTE])} | "
            f"Vencidos: {len(base[base['STATUS'] == STATUS_VENCIDO])}",
            estilo_info,
        )
    )
    elementos.append(Spacer(1, 14))

    dados = [[Paragraph(coluna.replace("_", " ").title(), estilo_cabecalho) for coluna in COLUNAS_RELATORIO]]
    for _, row in base.iterrows():
        dados.append([Paragraph(str(row.get(coluna, "") or "-"), estilo_info) for coluna in COLUNAS_RELATORIO])

    larguras = [100, 155, 82, 70, 82, 82, 90, 62]
    tabela = LongTable(dados, colWidths=larguras, repeatRows=1)
    estilo_tabela = [
        ("BACKGROUND", (0, 0), (-1, 0), COR_CABECALHO),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TEXTCOLOR", (0, 1), (-1, -1), COR_TEXTO),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
        ("GRID", (0, 0), (-1, -1), 0.35, COR_BORDA),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [COR_LINHA_1, COR_LINHA_2]),
    ]

    for linha_idx, (_, row) in enumerate(base.iterrows(), start=1):
        status = str(row.get("STATUS", "")).strip().upper()
        if status == STATUS_PENDENTE:
            estilo_tabela.append(("BACKGROUND", (0, linha_idx), (-1, linha_idx), COR_PENDENTE_BG))
            estilo_tabela.append(("TEXTCOLOR", (0, linha_idx), (-1, linha_idx), COR_PENDENTE_TX))
            estilo_tabela.append(("FONTNAME", (0, linha_idx), (-1, linha_idx), "Helvetica-Bold"))
        elif status in {STATUS_VENCIDO, STATUS_DATA_INVALIDA}:
            estilo_tabela.append(("BACKGROUND", (0, linha_idx), (-1, linha_idx), COR_VENCIDO_BG))
            estilo_tabela.append(("TEXTCOLOR", (0, linha_idx), (-1, linha_idx), COR_VENCIDO_TX))
            estilo_tabela.append(("FONTNAME", (0, linha_idx), (-1, linha_idx), "Helvetica-Bold"))

    tabela.setStyle(TableStyle(estilo_tabela))
    elementos.append(tabela)
    elementos.append(Spacer(1, 12))

    legenda = Table(
        [[
            Paragraph("Em dia: texto padrao", estilo_info),
            Paragraph("Pendente renovacao: destaque laranja/vermelho", estilo_info),
            Paragraph("Vencido/Data invalida: vermelho forte", estilo_info),
        ]],
        colWidths=[180, 240, 190],
    )
    legenda.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), COR_LINHA_1),
                ("BOX", (0, 0), (-1, -1), 0.35, COR_BORDA),
                ("INNERGRID", (0, 0), (-1, -1), 0.35, COR_BORDA),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    elementos.append(legenda)

    doc = SimpleDocTemplate(
        caminho_pdf,
        pagesize=landscape(A4),
        leftMargin=22,
        rightMargin=22,
        topMargin=20,
        bottomMargin=20,
    )
    doc.build(elementos)
    return caminho_pdf
