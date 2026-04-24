from __future__ import annotations

import copy
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from utils.caminhos import caminho_base


NS_W = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
MODELO_RELATORIO_TECNICO = "modelo.docx"

IDENTIFICACAO_PADRAO = {
    "data_inspecao": "11 de fevereiro de 2026",
    "horario": "14:00h",
    "local": "Belo Horizonte - MG",
    "equipamento": "Plataforma Elevatória Móvel de Trabalho",
    "tipo": "tipo mastro vertical",
    "sigla": "PEMT",
    "fabricante": "Genie",
    "modelo": "AWP 30 S",
    "numero_serie": "99-10758",
    "patrimonio": "130-007",
    "alimentacao": "Elétrico (Bateria)",
}

OBJETIVO_PADRAO = (
    "Avaliar as condições técnicas e operacionais da PEMT, verificando conformidade com requisitos de "
    "segurança estabelecidos pela NR-18 (Segurança e Saúde no Trabalho na Indústria da Construção) e "
    "normas técnicas aplicáveis, visando atestar a liberação do equipamento para uso operacional."
)

METODOLOGIA_PADRAO = (
    "A inspeção foi realizada mediante exame visual detalhado, testes funcionais dos sistemas elétricos, "
    "hidráulicos e mecânicos, verificação de dispositivos de segurança e consulta aos manuais técnicos do "
    "fabricante. Todos os itens críticos para operação segura foram avaliados conforme procedimentos estabelecidos."
)

ITENS_INSPECIONADOS_PADRAO = [
    {
        "titulo": "1. SISTEMA ELÉTRICO",
        "itens": [
            "Lâmpadas de sinalização e alarme sonoro (buzina)",
            "Motores elétricos de acionamento",
            "Recarregador de baterias (110V e 220V - 60Hz)",
            "Caixa de controle da plataforma",
            "Chaves, joystick e teclado de membrana",
            "Banco e compartimento de baterias",
            "Cabos elétricos de comando e tomadas",
            "Placas e adesivos informativos",
            "Adesivos indicadores de direção",
        ],
    },
    {
        "titulo": "2. SISTEMA HIDRÁULICO",
        "itens": [
            "Conjunto motor e bomba hidráulica",
            "Componentes elétricos/eletrônicos integrados",
            "Válvulas de controle e filtro hidráulico",
            "Mangueiras e tubulações (sem vazamentos)",
            "Nível e qualidade do óleo hidráulico",
            "Tanque hidráulico e sistema de respiro",
            "Cilindros de elevação (vedação e curso)",
            "Válvulas de contrabalanço",
        ],
    },
    {
        "titulo": "3. ESTRUTURA E CHASSI",
        "itens": [
            "Integridade estrutural do chassi",
            "Guarda-corpo e deck da plataforma",
            "Pintura e proteção anticorrosiva",
            "Placas de identificação do chassi",
        ],
    },
    {
        "titulo": "4. SISTEMA DE RODAGEM E DESLOCAMENTO",
        "itens": [
            "Pneus",
            "Rodas e cubos",
            "Rolamentos das rodas",
            "Sistema de direção",
            "Conjunto de frenagem",
        ],
    },
    {
        "titulo": "5. SISTEMAS DE CONTROLE",
        "itens": [
            "Controle de solo (comandos terrestres)",
            "Joystick da plataforma",
            "Sistema de descida de emergência",
            "Alarmes sonoros de segurança",
        ],
    },
    {
        "titulo": "6. SINALIZAÇÃO E DOCUMENTAÇÃO",
        "itens": [
            "Sinalização de segurança e operação",
            "Adesivos de instruções e alertas",
            "Manual de operação disponível",
            "Identificação de capacidade de carga",
        ],
    },
    {
        "titulo": "7. CONFORMIDADE NR-18",
        "itens": [
            "Requisitos de segurança atendidos",
            "Dispositivos de proteção operantes",
            "Sistemas de alarme funcionais",
            "Sinalização de segurança adequada",
        ],
    },
]

ESPECIFICACOES_PADRAO = {
    "deslocamento": (
        "O deslocamento deve ser realizado com cautela, observando possíveis declives e irregularidades do solo, "
        "visto que o equipamento não possui tração motorizada."
    ),
    "limite_inclinacao": (
        "A operação do equipamento deve ocorrer em superfície firme, estável e nivelada, respeitando o limite máximo "
        "de inclinação de 4° (quatro graus), conforme parâmetros técnicos de segurança aplicáveis ao modelo."
    ),
    "riscos_operacionais": (
        "A utilização do equipamento em condições superiores ao limite estabelecido caracteriza condição operacional "
        "insegura, podendo resultar em perda de estabilidade estrutural, comprometimento da segurança operacional e "
        "risco iminente de tombamento, com potencial para danos materiais e riscos à integridade física dos envolvidos."
    ),
    "indicacao_nivelamento": (
        "O equipamento não dispõe de sistema eletrônico de intertravamento ou bloqueio automático que impeça a "
        "elevação em condições de desnível. O controle de nivelamento é realizado exclusivamente por meio de nível "
        "de bolha instalado no chassi, com função meramente indicativa, não possuindo qualquer função de bloqueio "
        "operacional. Dessa forma, a verificação das condições de nivelamento e a decisão segura de operação são de "
        "responsabilidade direta do operador habilitado, devendo ser realizadas previamente ao início e durante toda "
        "a execução das atividades."
    ),
    "responsabilidades_operacionais": (
        "O cumprimento das recomendações técnicas, limites operacionais, normas regulamentadoras aplicáveis e "
        "orientações de segurança constantes neste laudo é de responsabilidade do operador, do locatário e de seus "
        "respectivos responsáveis técnicos e hierárquicos, cabendo-lhes assegurar a adequada utilização do equipamento."
    ),
    "advertencia_formal": (
        "O descumprimento das condições estabelecidas, bem como a operação em desacordo com as especificações do "
        "fabricante e boas práticas de segurança, poderá resultar em condições inseguras de trabalho, sendo que "
        "eventuais incidentes, acidentes, danos operacionais ou prejuízos decorrentes da utilização inadequada serão "
        "de responsabilidade dos envolvidos na operação e gestão do equipamento."
    ),
}

RESPONSABILIDADES_USUARIO_PADRAO = [
    "Executar diariamente a lista de verificação operacional antes do início das atividades",
    "Interromper imediatamente o uso do equipamento em caso de identificação de defeito ou anormalidade",
    "Acionar a locadora para reparo sempre que necessário",
    "Solicitar manutenção preventiva a cada 3 meses de uso (realizada pela locadora conforme cronograma de manutenção). Conservar o equipamento conforme orientações do fabricante e boas práticas operacionais",
]

VERIFICACAO_DIARIA_PADRAO = [
    {
        "titulo": "INSPEÇÃO VISUAL",
        "itens": [
            "Limpeza geral do equipamento",
            "Verificação de rodas, pneus e chassi",
            "Inspeção do pantógrafo, pinos, travões e buchas",
            "Verificação de vazamentos (hidráulico/elétrico)",
        ],
    },
    {
        "titulo": "SISTEMA DE BATERIAS",
        "itens": [
            "Aperto e limpeza dos terminais das baterias",
            "Teste do recarregador (110V e 220V - 60Hz)",
        ],
    },
    {
        "titulo": "TESTES FUNCIONAIS",
        "itens": [
            "Teste do sistema de descida de emergência",
            "Teste dos comandos do painel de solo",
            "Teste dos comandos da plataforma (joystick/botoeira)",
            "Verificação de alarmes sonoros",
            "Teste de movimentação (elevação/descida/deslocamento)",
        ],
    },
]

OBSERVACOES_IMPORTANTES_PADRAO = [
    "Este laudo aplica-se exclusivamente ao equipamento e modelo especificados",
    "A temperatura ambiente influencia diretamente o desempenho operacional e autonomia das baterias",
    "O não cumprimento das especificações técnicas pode resultar em acidentes graves e responsabilização civil e criminal",
]

RESPONSAVEL_TECNICO_PADRAO = {
    "nome": "FLAVIANO SILVEIRA QUEIROZ",
    "registro": "CFT MG 05986962613",
    "assinatura": "_________________________________",
    "local": "Belo Horizonte - MG",
    "data": "11 de fevereiro de 2026",
}


def criar_dados_padrao_relatorio_tecnico() -> dict:
    itens_inspecionados = [
        {"titulo": secao["titulo"], "itens": [{"texto": item, "ok": True} for item in secao["itens"]]}
        for secao in ITENS_INSPECIONADOS_PADRAO
    ]

    verificacao_diaria = [
        {"titulo": secao["titulo"], "itens": [{"texto": item, "ok": True} for item in secao["itens"]]}
        for secao in VERIFICACAO_DIARIA_PADRAO
    ]

    return {
        "identificacao": copy.deepcopy(IDENTIFICACAO_PADRAO),
        "objetivo": OBJETIVO_PADRAO,
        "metodologia": METODOLOGIA_PADRAO,
        "itens_inspecionados": itens_inspecionados,
        "especificacoes": copy.deepcopy(ESPECIFICACOES_PADRAO),
        "responsabilidades_usuario": list(RESPONSABILIDADES_USUARIO_PADRAO),
        "verificacao_diaria": verificacao_diaria,
        "observacoes_importantes": list(OBSERVACOES_IMPORTANTES_PADRAO),
        "responsavel_tecnico": copy.deepcopy(RESPONSAVEL_TECNICO_PADRAO),
    }


def _localizar_modelo_relatorio_tecnico() -> str:
    candidatos = [
        Path(caminho_base()) / MODELO_RELATORIO_TECNICO,
        Path.cwd() / MODELO_RELATORIO_TECNICO,
    ]
    for candidato in candidatos:
        if candidato.exists():
            return str(candidato)
    raise FileNotFoundError("Modelo do relatório técnico não encontrado.")


def _paragrafos_documento(raiz: ET.Element) -> list[ET.Element]:
    body = raiz.find("w:body", NS_W)
    if body is None:
        return []
    return body.findall("w:p", NS_W)


def _paragrafo_ou_erro(paragrafos: list[ET.Element], indice: int, contexto: str) -> ET.Element:
    try:
        return paragrafos[indice]
    except IndexError as exc:
        raise ValueError(
            f"O modelo Word está diferente do esperado. Parágrafo ausente em {contexto} (índice {indice})."
        ) from exc


def _runs_texto(paragrafo: ET.Element) -> list[ET.Element]:
    return [node for node in paragrafo.findall(".//w:t", NS_W)]


def _definir_texto_no_run(paragrafo: ET.Element, indice_run: int, novo_texto: str) -> None:
    runs = _runs_texto(paragrafo)
    if indice_run >= len(runs):
        raise ValueError(f"O modelo Word não possui runs suficientes para o trecho solicitado (run {indice_run}).")
    runs[indice_run].text = novo_texto


def _definir_texto_paragrafo(paragrafo: ET.Element, novo_texto: str) -> None:
    runs = _runs_texto(paragrafo)
    if not runs:
        raise ValueError("O modelo Word possui um parágrafo sem bloco de texto editável.")
    runs[0].text = novo_texto
    for run in runs[1:]:
        run.text = ""


def _texto_resultado(secao: dict) -> str:
    return "CONFORME" if all(item.get("ok", True) for item in secao["itens"]) else "NÃO CONFORME"


def _prefixo_item(ok: bool) -> str:
    return "✓" if ok else "X"


def _texto_identificacao_conclusao(identificacao: dict) -> str:
    return (
        f" que a {identificacao['equipamento']} ({identificacao['sigla']}), {identificacao['tipo']}, "
        f"marca {identificacao['fabricante']}, modelo {identificacao['modelo']}, número de série "
        f"{identificacao['numero_serie']}, patrimônio {identificacao['patrimonio']}, "
    )


def gerar_relatorio_tecnico_word(dados: dict, caminho_saida: str) -> str:
    modelo = _localizar_modelo_relatorio_tecnico()
    destino = Path(caminho_saida)
    destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(modelo, destino)

    identificacao = dados["identificacao"]
    responsavel = dados["responsavel_tecnico"]

    with zipfile.ZipFile(destino, "r") as docx:
        conteudos = {info.filename: docx.read(info.filename) for info in docx.infolist()}

    raiz = ET.fromstring(conteudos["word/document.xml"])
    body = raiz.find("w:body", NS_W)
    if body is None:
        raise ValueError("O modelo Word não possui corpo de documento válido.")

    paragrafos = _paragrafos_documento(raiz)

    _definir_texto_paragrafo(
        _paragrafo_ou_erro(paragrafos, 2, "identificação inicial"),
        (
            f"Data da Inspeção: {identificacao['data_inspecao']}\n"
            f"Horário: {identificacao['horario']}\n"
            f"Local: {identificacao['local']}"
        ),
    )

    _definir_texto_paragrafo(
        _paragrafo_ou_erro(paragrafos, 3, "dados do equipamento"),
        (
            f"Equipamento: {identificacao['equipamento']}\n"
            f"Tipo: {identificacao['tipo']} ({identificacao['sigla']})\n"
            f"Fabricante: {identificacao['fabricante']}\n"
            f"Modelo: {identificacao['modelo']}\n"
            f"Número de Série: {identificacao['numero_serie']}\n"
            f"Patrimônio: {identificacao['patrimonio']}\n"
            f"Tipo de Alimentação: {identificacao['alimentacao']}"
        ),
    )

    _definir_texto_paragrafo(_paragrafo_ou_erro(paragrafos, 6, "objetivo"), dados["objetivo"])
    _definir_texto_paragrafo(_paragrafo_ou_erro(paragrafos, 9, "metodologia"), dados["metodologia"])

    mapa_secoes = [
        (12, 13, 22),
        (23, 24, 32),
        (33, 34, 38),
        (39, 40, 45),
        (46, 47, 51),
        (52, 53, 57),
        (58, 59, 63),
    ]
    for secao_dados, (idx_titulo, idx_inicio, idx_resultado) in zip(dados["itens_inspecionados"], mapa_secoes):
        _definir_texto_paragrafo(_paragrafo_ou_erro(paragrafos, idx_titulo, secao_dados["titulo"]), secao_dados["titulo"])
        for deslocamento, item in enumerate(secao_dados["itens"]):
            _definir_texto_paragrafo(
                _paragrafo_ou_erro(paragrafos, idx_inicio + deslocamento, f"item da seção {secao_dados['titulo']}"),
                f"{_prefixo_item(item['ok'])} {item['texto']}",
            )
        _definir_texto_paragrafo(
            _paragrafo_ou_erro(paragrafos, idx_resultado, f"resultado da seção {secao_dados['titulo']}"),
            f"Resultado: {_texto_resultado(secao_dados)}",
        )

    _definir_texto_paragrafo(
        _paragrafo_ou_erro(paragrafos, 67, "deslocamento manual"),
        f"Deslocamento manual: {dados['especificacoes']['deslocamento']}",
    )
    _definir_texto_no_run(
        _paragrafo_ou_erro(paragrafos, 70, "limite de inclinação"),
        1,
        dados["especificacoes"]["limite_inclinacao"],
    )
    _definir_texto_no_run(_paragrafo_ou_erro(paragrafos, 70, "limite de inclinação"), 2, "")
    _definir_texto_no_run(_paragrafo_ou_erro(paragrafos, 70, "limite de inclinação"), 3, "")
    _definir_texto_paragrafo(
        _paragrafo_ou_erro(paragrafos, 71, "riscos operacionais"),
        f"Riscos operacionais associados: {dados['especificacoes']['riscos_operacionais']}",
    )
    _definir_texto_paragrafo(
        _paragrafo_ou_erro(paragrafos, 72, "indicação de nivelamento"),
        f"Sistema de indicação de nivelamento: {dados['especificacoes']['indicacao_nivelamento']}",
    )
    _definir_texto_paragrafo(
        _paragrafo_ou_erro(paragrafos, 73, "responsabilidades operacionais"),
        f"Responsabilidades operacionais e contratuais: {dados['especificacoes']['responsabilidades_operacionais']}",
    )
    _definir_texto_paragrafo(
        _paragrafo_ou_erro(paragrafos, 74, "advertência formal"),
        f"Advertência técnica formal: {dados['especificacoes']['advertencia_formal']}",
    )

    for indice, texto in enumerate(dados["responsabilidades_usuario"][:4]):
        _definir_texto_paragrafo(
            _paragrafo_ou_erro(paragrafos, 78 + indice, "responsabilidades do usuário"),
            texto,
        )

    checklist_indices = [
        (85, "Limpeza geral do equipamento"),
        (86, "Verificação de rodas, pneus e chassi"),
        (87, "Inspeção do pantógrafo, pinos, travões e buchas"),
        (88, "Verificação de vazamentos (hidráulico/elétrico)"),
        (90, "Aperto e limpeza dos terminais das baterias"),
        (91, "Teste do recarregador (110V e 220V - 60Hz)"),
        (93, "Teste do sistema de descida de emergência"),
        (94, "Teste dos comandos do painel de solo"),
        (95, "Teste dos comandos da plataforma (joystick/botoeira)"),
        (96, "Verificação de alarmes sonoros"),
        (97, "Teste de movimentação (elevação/descida/deslocamento)"),
    ]
    checklist_marcado = {
        item["texto"]
        for secao in dados["verificacao_diaria"]
        for item in secao["itens"]
        if item.get("ok", True)
    }
    for indice, texto in checklist_indices:
        if texto not in checklist_marcado:
            body.remove(_paragrafo_ou_erro(paragrafos, indice, f"checklist diário '{texto}'"))

    _definir_texto_no_run(
        _paragrafo_ou_erro(paragrafos, 100, "conclusão"),
        0,
        f"Após inspeção técnica detalhada realizada em {identificacao['data_inspecao']}, ",
    )
    _definir_texto_no_run(
        _paragrafo_ou_erro(paragrafos, 100, "conclusão"),
        2,
        _texto_identificacao_conclusao(identificacao),
    )

    _definir_texto_no_run(_paragrafo_ou_erro(paragrafos, 117, "data do responsável técnico"), 2, responsavel["data"])
    _definir_texto_paragrafo(
        _paragrafo_ou_erro(paragrafos, 118, "local do responsável técnico"),
        f"Local: {responsavel['local']}",
    )

    conteudos["word/document.xml"] = ET.tostring(raiz, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(destino, "w", compression=zipfile.ZIP_DEFLATED) as docx:
        for nome_arquivo, dados_arquivo in conteudos.items():
            docx.writestr(nome_arquivo, dados_arquivo)

    return str(destino)
