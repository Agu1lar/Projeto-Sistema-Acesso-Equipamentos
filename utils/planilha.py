from __future__ import annotations

import shutil
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

from utils.caminhos import caminho_dados

ABAS_ALIAS = {
    "VEICULOS": {"VEICULOS", "VEICULO", "VEÍCULOS", "VEÍCULO"},
    "MANUTENCOES": {"MANUTENCOES", "MANUTENCAO", "MANUTENÇÕES", "MANUTENÇÃO"},
    "AGENDAMENTOS": {"AGENDAMENTOS", "AGENDAMENTO"},
}

COLUNAS_ALIAS = {
    "VEICULOS": {
        "PATRIMONIO": {"PATRIMONIO", "PLACA", "VEICULO"},
        "DATA_ATUALIZACAO": {"DATA_ATUALIZACAO", "ATUALIZADO_EM", "DATA"},
    },
    "MANUTENCOES": {
        "PATRIMONIO": {"PATRIMONIO", "PLACA", "VEICULO"},
        "DESCRICAO": {"DESCRICAO", "DESCRIÇÃO"},
        "CATEGORIA": {"CATEGORIA", "TIPO", "TIPO_MANUTENCAO"},
        "DETALHE": {"DETALHE", "DETALHES"},
        "DATA_INICIO": {"DATA_INICIO", "INICIO", "DATA_INICIAL"},
        "DATA_FIM": {"DATA_FIM", "FIM", "DATA_FINAL"},
        "HORIMETRO": {"HORIMETRO", "HORÍMETRO"},
        "HORIMETRO_ATUAL": {"HORIMETRO_ATUAL", "HORÍMETRO_ATUAL"},
        "HORIMETRO_TROCA": {"HORIMETRO_TROCA", "HORÍMETRO_TROCA"},
        "SITUACAO_HORIMETRO": {"SITUACAO_HORIMETRO", "SIT_HORIMETRO"},
        "SITUACAO_DATA": {"SITUACAO_DATA", "STATUS_DATA"},
    },
    "AGENDAMENTOS": {
        "PATRIMONIO": {"PATRIMONIO", "PLACA", "VEICULO"},
        "DESCRICAO": {"DESCRICAO", "DESCRIÇÃO"},
        "HORIMETRO_ATUAL": {"HORIMETRO_ATUAL", "HORÍMETRO_ATUAL"},
        "HORIMETRO_TROCA": {"HORIMETRO_TROCA", "HORÍMETRO_TROCA"},
    },
}


def normalizar_identificador(texto) -> str:
    valor = str(texto or "").strip().upper()
    valor = unicodedata.normalize("NFKD", valor).encode("ASCII", "ignore").decode("ASCII")
    return valor.replace(" ", "_")


def normalizar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame()

    df_normalizado = df.copy()
    df_normalizado = df_normalizado.loc[:, ~df_normalizado.columns.astype(str).str.contains("^UNNAMED", case=False)]
    df_normalizado.columns = [normalizar_identificador(coluna) for coluna in df_normalizado.columns]
    return df_normalizado.fillna("")


def _mapa_alias_colunas(aba: str) -> dict[str, str]:
    mapa = {}
    for coluna_destino, aliases in COLUNAS_ALIAS.get(aba, {}).items():
        for alias in aliases:
            mapa[normalizar_identificador(alias)] = coluna_destino
    return mapa


def padronizar_dataframe_aba(df: pd.DataFrame, aba: str, colunas_esperadas: list[str]) -> pd.DataFrame:
    df_padronizado = normalizar_dataframe(df)
    mapa_alias = _mapa_alias_colunas(aba)

    renomear = {}
    for coluna in df_padronizado.columns:
        destino = mapa_alias.get(coluna)
        if destino and destino not in df_padronizado.columns:
            renomear[coluna] = destino

    if renomear:
        df_padronizado = df_padronizado.rename(columns=renomear)

    for coluna in colunas_esperadas:
        if coluna not in df_padronizado.columns:
            df_padronizado[coluna] = ""

    extras = [coluna for coluna in df_padronizado.columns if coluna not in colunas_esperadas]
    return df_padronizado[colunas_esperadas + extras]


def resolver_nome_aba(sheet_names: list[str], aba_destino: str) -> str | None:
    alvo_normalizado = normalizar_identificador(aba_destino)
    aliases = {normalizar_identificador(valor) for valor in ABAS_ALIAS.get(aba_destino, {aba_destino})}

    for nome in sheet_names:
        nome_normalizado = normalizar_identificador(nome)
        if nome_normalizado == alvo_normalizado or nome_normalizado in aliases:
            return nome

    return None


def carregar_todas_abas_seguras(arquivo: str, estrutura_abas: dict[str, list[str]]) -> dict[str, pd.DataFrame]:
    abas = {aba: pd.DataFrame(columns=colunas) for aba, colunas in estrutura_abas.items()}
    caminho = Path(arquivo)

    if not caminho.exists():
        return abas

    try:
        planilha = pd.ExcelFile(caminho)
    except Exception:
        return abas

    for aba_destino, colunas in estrutura_abas.items():
        nome_real = resolver_nome_aba(planilha.sheet_names, aba_destino)
        if not nome_real:
            continue

        try:
            df = pd.read_excel(caminho, sheet_name=nome_real, dtype=str)
        except Exception:
            df = pd.DataFrame()

        abas[aba_destino] = padronizar_dataframe_aba(df, aba_destino, colunas)

    return abas


def criar_backup_planilha(arquivo: str) -> str | None:
    origem = Path(arquivo)
    if not origem.exists():
        return None

    pasta_backup = Path(caminho_dados()) / "backups"
    pasta_backup.mkdir(parents=True, exist_ok=True)

    destino = pasta_backup / f"{origem.stem}_{datetime.now():%Y%m%d_%H%M%S}{origem.suffix}"
    shutil.copy2(origem, destino)
    return str(destino)
