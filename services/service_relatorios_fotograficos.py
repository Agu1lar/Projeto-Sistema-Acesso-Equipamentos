from __future__ import annotations

import json
import os
import re
from datetime import datetime
from pathlib import Path

from utils.caminhos import caminho_config_relatorios, caminho_relatorios_fotograficos


def _slug(texto: str, fallback: str) -> str:
    valor = re.sub(r"[^\w\-\.]+", "_", (texto or "").strip(), flags=re.UNICODE)
    valor = valor.strip("._")
    return valor or fallback


def carregar_configuracao() -> dict:
    caminho = Path(caminho_config_relatorios())
    if not caminho.exists():
        return {}

    try:
        return json.loads(caminho.read_text(encoding="utf-8"))
    except Exception:
        return {}


def salvar_pasta_base(pasta_base: str) -> None:
    caminho = Path(caminho_config_relatorios())
    caminho.parent.mkdir(parents=True, exist_ok=True)
    payload = {"pasta_base": pasta_base}
    caminho.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def obter_pasta_base() -> str:
    configuracao = carregar_configuracao()
    pasta_base = configuracao.get("pasta_base")

    if pasta_base:
        Path(pasta_base).mkdir(parents=True, exist_ok=True)
        return pasta_base

    return caminho_relatorios_fotograficos()


def construir_nome_pdf(contrato: str, modelo: str, data_referencia: datetime | None = None) -> str:
    data_base = data_referencia or datetime.now()
    nome_modelo = _slug(modelo, "modelo")
    nome_contrato = _slug(contrato, "contrato")
    return f"Relatorio_{nome_modelo}_{nome_contrato}_{data_base:%d-%m-%Y_%H-%M}.pdf"


def caminho_pdf_modelo(modelo: str, contrato: str, data_referencia: datetime | None = None) -> str:
    modelo_limpo = _slug(modelo, "sem_modelo")
    pasta_modelo = Path(obter_pasta_base()) / modelo_limpo
    pasta_modelo.mkdir(parents=True, exist_ok=True)
    return str(pasta_modelo / construir_nome_pdf(contrato, modelo, data_referencia))


def listar_relatorios(modelo: str = "", data: str = "") -> list[dict]:
    pasta_base = Path(obter_pasta_base())
    if not pasta_base.exists():
        return []

    filtro_modelo = _slug(modelo, "") if modelo.strip() else ""
    data_filtro = None

    if data.strip():
        try:
            data_filtro = datetime.strptime(data.strip(), "%d/%m/%Y").date()
        except ValueError as exc:
            raise ValueError("Use a data no formato dd/mm/aaaa.") from exc

    resultados = []
    for arquivo in pasta_base.rglob("*.pdf"):
        if filtro_modelo and filtro_modelo.lower() not in arquivo.parent.name.lower():
            continue

        data_arquivo = datetime.fromtimestamp(arquivo.stat().st_mtime)
        if data_filtro and data_arquivo.date() != data_filtro:
            continue

        resultados.append(
            {
                "modelo": arquivo.parent.name,
                "arquivo": arquivo.name,
                "data": data_arquivo.strftime("%d/%m/%Y %H:%M"),
                "timestamp": data_arquivo.timestamp(),
                "caminho": str(arquivo),
            }
        )

    resultados.sort(key=lambda item: item["timestamp"], reverse=True)
    return resultados


def abrir_relatorio(caminho: str) -> None:
    arquivo = Path(caminho)
    if not arquivo.exists():
        raise FileNotFoundError("O relatório selecionado não existe mais.")
    os.startfile(str(arquivo))
