from __future__ import annotations

from pathlib import Path

import pandas as pd


VEICULOS_COLUNAS = [
    "PATRIMONIO",
    "HORIMETRO",
    "DATA_ATUALIZACAO",
    "MARCA",
    "ANO",
    "OBS",
]

MANUTENCOES_COLUNAS = [
    "DATA",
    "PATRIMONIO",
    "HORIMETRO",
    "DESCRICAO",
    "CATEGORIA",
    "VALOR",
    "HORIMETRO_ATUAL",
    "HORIMETRO_TROCA",
    "SITUACAO_HORIMETRO",
    "SITUACAO_DATA",
    "DETALHE",
    "DATA_INICIO",
    "DATA_FIM",
    "MES",
    "ANO",
]

AGENDAMENTOS_COLUNAS = [
    "PATRIMONIO",
    "DESCRICAO",
    "HORIMETRO_ATUAL",
    "HORIMETRO_TROCA",
    "HORIMETRO_PROX",
    "DATA_PROX",
    "STATUS",
]

ESTRUTURA_ABAS = {
    "VEICULOS": VEICULOS_COLUNAS,
    "MANUTENCOES": MANUTENCOES_COLUNAS,
    "AGENDAMENTOS": AGENDAMENTOS_COLUNAS,
}


class BancoDados:
    def __init__(self, arquivo: str):
        self.arquivo = str(Path(arquivo))
        self.inicializar_arquivo()

    def salvar(self, aba: str, dados: dict) -> None:
        df_existente = self.carregar_dataframe(aba)
        novo_df = pd.concat([df_existente, pd.DataFrame([dados])], ignore_index=True)
        self.escrever_aba(aba, novo_df)

    def carregar_veiculos(self) -> list[str]:
        df = self.carregar_dataframe("VEICULOS")
        if "PATRIMONIO" not in df.columns:
            return []
        return df["PATRIMONIO"].dropna().astype(str).tolist()

    def ler_aba(self, aba: str) -> str:
        df = self.carregar_dataframe(aba)
        if df.empty:
            return "Nenhum dado encontrado."
        return df.to_string(index=False)

    def carregar_dataframe(self, aba: str) -> pd.DataFrame:
        colunas = ESTRUTURA_ABAS.get(aba, [])
        try:
            df = pd.read_excel(self.arquivo, sheet_name=aba)
        except Exception:
            return pd.DataFrame(columns=colunas)

        for coluna in colunas:
            if coluna not in df.columns:
                df[coluna] = ""

        if colunas:
            extras = [col for col in df.columns if col not in colunas]
            df = df[colunas + extras]

        return df

    def escrever_aba(self, aba: str, dataframe: pd.DataFrame) -> None:
        abas = self._carregar_todas_as_abas()
        abas[aba] = dataframe.copy()

        with pd.ExcelWriter(self.arquivo, engine="openpyxl") as writer:
            for nome_aba, df in abas.items():
                df.to_excel(writer, sheet_name=nome_aba, index=False)

    def inicializar_arquivo(self) -> None:
        if not Path(self.arquivo).exists():
            self.criar_estrutura()

    def criar_estrutura(self) -> None:
        with pd.ExcelWriter(self.arquivo, engine="openpyxl") as writer:
            for aba, colunas in ESTRUTURA_ABAS.items():
                pd.DataFrame(columns=colunas).to_excel(
                    writer,
                    sheet_name=aba,
                    index=False,
                )

    def _carregar_todas_as_abas(self) -> dict[str, pd.DataFrame]:
        abas = {
            nome_aba: pd.DataFrame(columns=colunas)
            for nome_aba, colunas in ESTRUTURA_ABAS.items()
        }

        if not Path(self.arquivo).exists():
            return abas

        try:
            arquivo_excel = pd.ExcelFile(self.arquivo)
        except Exception:
            return abas

        for nome_aba in arquivo_excel.sheet_names:
            abas[nome_aba] = pd.read_excel(self.arquivo, sheet_name=nome_aba)

        return abas
