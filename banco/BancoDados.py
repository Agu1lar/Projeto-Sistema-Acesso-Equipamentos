from __future__ import annotations

from pathlib import Path

import pandas as pd

from utils.planilha import carregar_todas_abas_seguras, criar_backup_planilha, padronizar_dataframe_aba


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

TREINAMENTOS_COLUNAS = [
    "CODIGO",
    "NOME",
    "CPF",
    "EMPRESA",
    "FUNCAO",
    "TREINAMENTO",
    "CARGA_HORARIA",
    "INSTRUTOR",
    "DATA_EMISSAO",
    "VALIDADE",
    "RESPONSAVEL",
    "OBS",
    "CERTIFICADO_RUA",
    "CERTIFICADO_NUMERO",
    "CERTIFICADO_BAIRRO",
    "CERTIFICADO_CIDADE",
    "CERTIFICADO_UF",
    "CERTIFICADO_CEP",
    "CERTIFICADO_IMPRIMIR_CPF",
    "CERTIFICADO_MODELO",
    "DOCUMENTOS_PASTA",
    "CERTIFICADO_PDF_CAMINHO",
    "CERTIFICADO_WORD_CAMINHO",
    "PDF_CAMINHO",
    "MODELO_VERSAO",
    "DATA_CADASTRO",
    "DATA_ATUALIZACAO",
]

CERTIFICADOS_COLUNAS = [
    "TIPO_CERTIFICADO",
    "NOME",
    "CPF",
    "CARGA_HORARIA",
    "DATA_EMISSAO",
    "DATA_VENCIMENTO",
    "STATUS",
    "DIAS_RESTANTES",
    "OBS",
    "DATA_CADASTRO",
]

ESTRUTURA_ABAS = {
    "VEICULOS": VEICULOS_COLUNAS,
    "MANUTENCOES": MANUTENCOES_COLUNAS,
    "AGENDAMENTOS": AGENDAMENTOS_COLUNAS,
    "TREINAMENTOS": TREINAMENTOS_COLUNAS,
    "CERTIFICADOS": CERTIFICADOS_COLUNAS,
}


class BancoDados:
    def __init__(self, arquivo: str):
        self.arquivo = str(Path(arquivo))
        self.inicializar_arquivo()
        self.reparar_estrutura()

    def salvar(self, aba: str, dados: dict) -> None:
        df_existente = self.carregar_dataframe(aba)
        novo_df = pd.concat([df_existente, pd.DataFrame([dados])], ignore_index=True)
        self.escrever_aba(aba, novo_df)

    def carregar_veiculos(self) -> list[str]:
        df = self.carregar_dataframe("VEICULOS")
        if "PATRIMONIO" not in df.columns:
            return []
        return sorted({valor for valor in df["PATRIMONIO"].dropna().astype(str) if valor.strip()})

    def ler_aba(self, aba: str) -> str:
        df = self.carregar_dataframe(aba)
        if df.empty:
            return "Nenhum dado encontrado."
        return df.to_string(index=False)

    def carregar_dataframe(self, aba: str) -> pd.DataFrame:
        abas = carregar_todas_abas_seguras(self.arquivo, ESTRUTURA_ABAS)
        return abas.get(aba, pd.DataFrame(columns=ESTRUTURA_ABAS.get(aba, [])))

    def escrever_aba(self, aba: str, dataframe: pd.DataFrame) -> None:
        abas = self._carregar_todas_as_abas()
        colunas_esperadas = ESTRUTURA_ABAS.get(aba, [])
        abas[aba] = padronizar_dataframe_aba(dataframe, aba, colunas_esperadas)

        criar_backup_planilha(self.arquivo)
        self._escrever_abas(abas)

    def inicializar_arquivo(self) -> None:
        if not Path(self.arquivo).exists():
            self.criar_estrutura()

    def criar_estrutura(self) -> None:
        with pd.ExcelWriter(self.arquivo, engine="openpyxl") as writer:
            for aba, colunas in ESTRUTURA_ABAS.items():
                pd.DataFrame(columns=colunas).to_excel(writer, sheet_name=aba, index=False)

    def reparar_estrutura(self) -> None:
        abas = self._carregar_todas_as_abas()
        precisa_reescrever = False

        for aba, colunas in ESTRUTURA_ABAS.items():
            df_atual = abas.get(aba, pd.DataFrame())
            df_padronizado = padronizar_dataframe_aba(df_atual, aba, colunas)
            if list(df_atual.columns) != list(df_padronizado.columns):
                precisa_reescrever = True
            abas[aba] = df_padronizado

        if precisa_reescrever:
            criar_backup_planilha(self.arquivo)
            self._escrever_abas(abas)

    def _carregar_todas_as_abas(self) -> dict[str, pd.DataFrame]:
        return carregar_todas_abas_seguras(self.arquivo, ESTRUTURA_ABAS)

    def _escrever_abas(self, abas: dict[str, pd.DataFrame]) -> None:
        try:
            with pd.ExcelWriter(self.arquivo, engine="openpyxl") as writer:
                for nome_aba, colunas in ESTRUTURA_ABAS.items():
                    df = abas.get(nome_aba, pd.DataFrame(columns=colunas))
                    df.to_excel(writer, sheet_name=nome_aba, index=False)
        except PermissionError as exc:
            raise PermissionError(
                "Não foi possível salvar a planilha. Feche o Excel e tente novamente."
            ) from exc
