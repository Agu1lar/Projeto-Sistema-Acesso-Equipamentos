from __future__ import annotations

from datetime import datetime

import pandas as pd


def normalizar_patrimonio(valor: str) -> str:
    return str(valor or "").strip().upper()


def localizar_veiculo(df_veiculos: pd.DataFrame, patrimonio: str) -> pd.Series | None:
    patrimonio_norm = normalizar_patrimonio(patrimonio)
    if not patrimonio_norm or df_veiculos.empty or "PATRIMONIO" not in df_veiculos.columns:
        return None

    filtro = df_veiculos["PATRIMONIO"].astype(str).str.upper() == patrimonio_norm
    if not filtro.any():
        return None
    return df_veiculos.loc[filtro].iloc[-1]


def resumo_veiculo(banco, patrimonio: str) -> dict | None:
    patrimonio_norm = normalizar_patrimonio(patrimonio)
    if not patrimonio_norm:
        return None

    df_veiculos = banco.carregar_dataframe("VEICULOS").fillna("")
    df_manut = banco.carregar_dataframe("MANUTENCOES").fillna("")
    veiculo = localizar_veiculo(df_veiculos, patrimonio_norm)

    manut = df_manut[df_manut["PATRIMONIO"].astype(str).str.upper() == patrimonio_norm] if not df_manut.empty else pd.DataFrame()
    valores = pd.to_numeric(manut.get("VALOR"), errors="coerce").fillna(0) if not manut.empty else pd.Series(dtype=float)

    return {
        "patrimonio": patrimonio_norm,
        "marca": str(veiculo.get("MARCA", "")).strip() if veiculo is not None else "",
        "ano": str(veiculo.get("ANO", "")).strip() if veiculo is not None else "",
        "horimetro_base": str(veiculo.get("HORIMETRO", "")).strip() if veiculo is not None else "",
        "tem_cadastro": veiculo is not None,
        "manutencoes": int(len(manut)),
        "gasto_total": float(valores.sum()) if not manut.empty else 0.0,
        "ultima_categoria": str(manut.iloc[-1].get("CATEGORIA", "")).strip() if not manut.empty else "",
        "ultimo_horimetro": str(manut.iloc[-1].get("HORIMETRO_ATUAL", "")).strip() if not manut.empty else "",
    }


def atualizar_veiculo_por_manutencao(banco, patrimonio: str, horimetro_atual: float) -> None:
    patrimonio_norm = normalizar_patrimonio(patrimonio)
    if not patrimonio_norm:
        return

    df_veiculos = banco.carregar_dataframe("VEICULOS").fillna("").copy()
    filtro = df_veiculos["PATRIMONIO"].astype(str).str.upper() == patrimonio_norm
    if not filtro.any():
        return

    indice = df_veiculos.index[filtro][-1]
    df_veiculos.loc[indice, "HORIMETRO"] = horimetro_atual
    df_veiculos.loc[indice, "DATA_ATUALIZACAO"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    banco.escrever_aba("VEICULOS", df_veiculos)


def cadastrar_veiculo_automatico(banco, patrimonio: str, horimetro_atual: float, observacao: str = "") -> None:
    patrimonio_norm = normalizar_patrimonio(patrimonio)
    if not patrimonio_norm:
        return

    df_veiculos = banco.carregar_dataframe("VEICULOS").fillna("")
    if not df_veiculos.empty and (df_veiculos["PATRIMONIO"].astype(str).str.upper() == patrimonio_norm).any():
        return

    banco.salvar(
        "VEICULOS",
        {
            "PATRIMONIO": patrimonio_norm,
            "HORIMETRO": horimetro_atual,
            "DATA_ATUALIZACAO": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "MARCA": "",
            "ANO": "",
            "OBS": observacao.strip(),
        },
    )


def listar_inconsistencias(banco) -> list[dict]:
    inconsistencias = []
    df_veiculos = banco.carregar_dataframe("VEICULOS").fillna("")
    df_manut = banco.carregar_dataframe("MANUTENCOES").fillna("")

    if not df_veiculos.empty:
        duplicados = df_veiculos["PATRIMONIO"].astype(str).str.upper()
        for patrimonio in sorted({valor for valor in duplicados[duplicados.duplicated()] if valor.strip()}):
            inconsistencias.append({"tipo": "Veiculo duplicado", "referencia": patrimonio, "origem": "VEICULOS"})

        for _, row in df_veiculos.iterrows():
            patrimonio = normalizar_patrimonio(row.get("PATRIMONIO", ""))
            if patrimonio and (not str(row.get("MARCA", "")).strip() or not str(row.get("ANO", "")).strip()):
                inconsistencias.append({"tipo": "Cadastro incompleto", "referencia": patrimonio, "origem": "VEICULOS"})

    if not df_manut.empty:
        patrimonios_veiculo = set(df_veiculos["PATRIMONIO"].astype(str).str.upper()) if not df_veiculos.empty else set()
        for _, row in df_manut.iterrows():
            patrimonio = normalizar_patrimonio(row.get("PATRIMONIO", ""))
            if patrimonio and patrimonio not in patrimonios_veiculo:
                inconsistencias.append({"tipo": "Manutencao sem veiculo", "referencia": patrimonio, "origem": "MANUTENCOES"})

            try:
                horimetro_atual = float(str(row.get("HORIMETRO_ATUAL", "")).replace(",", ".") or 0)
                horimetro_troca = float(str(row.get("HORIMETRO_TROCA", "")).replace(",", ".") or 0)
                if horimetro_troca and horimetro_atual > horimetro_troca + 500:
                    inconsistencias.append(
                        {"tipo": "Horimetro fora do esperado", "referencia": patrimonio, "origem": "MANUTENCOES"}
                    )
            except Exception:
                inconsistencias.append({"tipo": "Valor invalido", "referencia": patrimonio, "origem": "MANUTENCOES"})

    return inconsistencias
