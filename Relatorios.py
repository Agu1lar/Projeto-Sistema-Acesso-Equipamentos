from __future__ import annotations

import datetime
import os
from pathlib import Path
from tkinter import filedialog, messagebox

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import LongTable, Paragraph, SimpleDocTemplate, Spacer, TableStyle


class Relatorios:
    def __init__(self, arquivo):
        self.arquivo = arquivo
        self.estilos = getSampleStyleSheet()

    def relatorio_por_categoria(self, categoria, aba="MANUTENCOES"):
        try:
            df = self._carregar_planilha(aba)
            df = df[df["CATEGORIA"].astype(str).str.upper() == categoria.upper()]
            if df.empty:
                messagebox.showwarning("Aviso", f"Sem dados para {categoria}")
                return
            self.relatorio_filtrado(df, categoria)
        except Exception as erro:
            messagebox.showerror("Erro", str(erro))

    def relatorio_manutencoes(self):
        try:
            df = self._carregar_planilha("MANUTENCOES")
            if df.empty:
                messagebox.showwarning("Aviso", "Sem dados")
                return
            self._gerar_relatorio_tabulado(df, "GERAL")
        except Exception as erro:
            messagebox.showerror("Erro", str(erro))

    def relatorio_por_veiculo(self, patrimonio):
        try:
            df = self._carregar_planilha("MANUTENCOES")
            df = df[df["PATRIMONIO"].astype(str) == str(patrimonio)]
            if df.empty:
                messagebox.showwarning("Aviso", "Sem dados")
                return
            self._gerar_relatorio_tabulado(df, patrimonio)
        except Exception as erro:
            messagebox.showerror("Erro", str(erro))

    def resumo_por_veiculo(self, patrimonio):
        try:
            df = self._carregar_planilha("MANUTENCOES")
        except Exception:
            return None

        df = df[df["PATRIMONIO"].astype(str) == str(patrimonio)]
        if df.empty:
            return None

        valores = pd.to_numeric(df.get("VALOR"), errors="coerce").fillna(0)
        return {"total": len(df), "gasto": float(valores.sum())}

    def relatorio_filtrado(self, df, aba):
        try:
            if df.empty:
                messagebox.showwarning("Aviso", "Sem dados com esse filtro")
                return

            df_exportacao = df.copy()
            total = self._calcular_total(df_exportacao)
            nome_padrao = f"relatorio_{aba}_{datetime.datetime.now():%d-%m-%Y}.pdf"
            caminho = filedialog.asksaveasfilename(
                initialfile=nome_padrao,
                defaultextension=".pdf",
                filetypes=[("Arquivos PDF", "*.pdf")],
                title="Salvar relatório como",
            )
            if not caminho:
                return

            self._build_pdf(caminho, f"RELATÓRIO - {aba}", total, df_exportacao)
            self._abrir_pdf(caminho)
            messagebox.showinfo("Sucesso", f"Relatório gerado.\n\nArquivo: {caminho}")
        except Exception as erro:
            messagebox.showerror("Erro", f"Erro ao gerar relatório:\n{erro}")

    def _carregar_planilha(self, aba):
        return pd.read_excel(self.arquivo, sheet_name=aba, dtype=str).fillna("")

    def _calcular_total(self, df):
        if "VALOR" not in df.columns:
            return 0.0
        return float(pd.to_numeric(df["VALOR"], errors="coerce").fillna(0).sum())

    def _gerar_relatorio_tabulado(self, df, nome_base):
        nome = f"relatorio_{nome_base}_{datetime.datetime.now():%d-%m-%Y}.pdf"
        total = self._calcular_total(df)
        self._build_pdf(nome, f"RELATÓRIO - {nome_base}", total, df)
        self._abrir_pdf(nome)

    def _build_pdf(self, caminho, titulo, total, df):
        doc = SimpleDocTemplate(
            caminho,
            pagesize=landscape(A4),
            leftMargin=20,
            rightMargin=20,
            topMargin=20,
            bottomMargin=20,
        )

        elementos = [
            Paragraph(titulo, self.estilos["Title"]),
            Spacer(1, 10),
            Paragraph(f"Total: R$ {total:.2f}", self.estilos["Normal"]),
            Spacer(1, 10),
        ]

        tabela = self._montar_tabela(df, doc.width)
        elementos.append(tabela)
        doc.build(elementos)

    def _montar_tabela(self, df, largura_total):
        df_texto = df.fillna("").astype(str)
        dados = [[Paragraph(str(coluna), self.estilos["BodyText"]) for coluna in df_texto.columns]]

        for _, row in df_texto.iterrows():
            dados.append([Paragraph(valor, self.estilos["BodyText"]) for valor in row])

        pesos = []
        for coluna in df_texto.columns:
            maior_texto = df_texto[coluna].astype(str).str.len().max() if not df_texto.empty else 10
            pesos.append(max(len(str(coluna)), maior_texto, 10))

        soma_pesos = sum(pesos) or len(df_texto.columns) or 1
        larguras = [(peso / soma_pesos) * largura_total for peso in pesos]

        tabela = LongTable(dados, colWidths=larguras, repeatRows=1)
        tabela.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 7),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]),
                ]
            )
        )
        return tabela

    @staticmethod
    def _abrir_pdf(caminho):
        caminho_pdf = Path(caminho)
        if caminho_pdf.exists():
            os.startfile(str(caminho_pdf))
