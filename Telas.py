from __future__ import annotations

import unicodedata
from datetime import datetime
from tkinter import messagebox, ttk

import customtkinter as ctk
import pandas as pd

from Relatorios import Relatorios

CATEGORIAS_MANUTENCAO = ["GERAL", "OLEO", "PNEUS", "MECANICA", "ELETRICA"]

COLUNAS_EXIBICAO = {
    "VEICULOS": ["PATRIMONIO", "MARCA", "ANO", "HORIMETRO", "DATA_ATUALIZACAO"],
    "MANUTENCOES": [
        "PATRIMONIO",
        "CATEGORIA",
        "ANO",
        "DATA_INICIO",
        "DATA_FIM",
        "VALOR",
        "HORIMETRO_ATUAL",
        "HORIMETRO_TROCA",
    ],
}

NOMES_COLUNAS = {
    "PATRIMONIO": "Placa",
    "MARCA": "Marca",
    "ANO": "Ano",
    "HORIMETRO": "Horímetro",
    "DATA_ATUALIZACAO": "Atualizado em",
    "CATEGORIA": "Tipo",
    "DATA_INICIO": "Data Início",
    "DATA_FIM": "Data Fim",
    "VALOR": "Valor (R$)",
    "HORIMETRO_ATUAL": "Horímetro Atual",
    "HORIMETRO_TROCA": "Horímetro Troca",
}


def limpar_campos(campos):
    for campo in campos:
        if isinstance(campo, ctk.CTkEntry):
            campo.delete(0, "end")
        elif isinstance(campo, ctk.CTkComboBox):
            campo.set("")


def normalizar_coluna(coluna):
    texto = str(coluna).strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    return texto.replace(" ", "_")


def carregar_df_normalizado(arquivo, aba):
    try:
        df = pd.read_excel(arquivo, sheet_name=aba, dtype=str).fillna("")
    except Exception:
        return pd.DataFrame()

    df = df.loc[:, ~df.columns.astype(str).str.contains("^UNNAMED", case=False)]
    df.columns = [normalizar_coluna(coluna) for coluna in df.columns]
    return df.reset_index(drop=True)


class Telas:
    def __init__(self, app, frame, banco, voltar_menu):
        self.app = app
        self.frame = frame
        self.banco = banco
        self.voltar_menu = voltar_menu
        self.relatorios = Relatorios(self.banco.arquivo)
        self.df_filtrado = pd.DataFrame()

    def limpar(self):
        for widget in self.frame.winfo_children():
            widget.destroy()

    def criar_card(self, titulo):
        container = ctk.CTkFrame(self.frame, fg_color="transparent")
        container.pack(expand=True)

        card = ctk.CTkFrame(container, corner_radius=15)
        card.pack(padx=20, pady=20)

        ctk.CTkLabel(card, text=titulo, font=("Arial", 24, "bold")).pack(pady=20)
        return card

    def _colunas_visiveis(self, aba, df):
        desejadas = COLUNAS_EXIBICAO.get(aba, list(df.columns))
        existentes = [coluna for coluna in desejadas if coluna in df.columns]
        return existentes or list(df.columns)

    def visualizar(self, aba):
        item_selecionado = None
        self.limpar()

        container = ctk.CTkFrame(self.frame)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(container, text=f"DADOS - {aba}", font=("Arial", 22, "bold")).pack(pady=10)

        df = carregar_df_normalizado(self.banco.arquivo, aba)
        if "MES" in df.columns:
            df = df.drop(columns=["MES"])

        self.df_filtrado = df.copy()
        colunas_existentes = self._colunas_visiveis(aba, df)

        edicao_frame = ctk.CTkFrame(container)
        edicao_frame.pack(fill="x", pady=10)

        linhas_campos = [
            ["PATRIMONIO", "CATEGORIA", "ANO"],
            ["VALOR", "HORIMETRO_ATUAL", "HORIMETRO_TROCA"],
            ["SITUACAO_HORIMETRO", "SITUACAO_DATA"],
            ["DATA_INICIO", "DATA_FIM"],
            ["DATA"],
        ]
        campos = {}

        for row_index, linha in enumerate(linhas_campos):
            for col_index, coluna in enumerate(linha):
                if coluna not in df.columns:
                    continue
                entry = ctk.CTkEntry(edicao_frame, placeholder_text=coluna, width=180)
                entry.grid(row=row_index, column=col_index, padx=10, pady=8)
                campos[coluna] = entry

        tabela_frame = ctk.CTkFrame(container)
        tabela_frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(tabela_frame, columns=colunas_existentes, show="headings")
        scroll_y = ttk.Scrollbar(tabela_frame, orient="vertical", command=tree.yview)
        scroll_x = ttk.Scrollbar(tabela_frame, orient="horizontal", command=tree.xview)

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        tree.pack(side="left", fill="both", expand=True)
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        for coluna in colunas_existentes:
            tree.heading(coluna, text=NOMES_COLUNAS.get(coluna, coluna))
            tree.column(coluna, anchor="center", width=140)

        def atualizar(df_filtrado):
            tree.delete(*tree.get_children())
            for indice, row in df_filtrado.iterrows():
                valores = [str(row.get(coluna, "")) for coluna in colunas_existentes]
                tree.insert("", "end", iid=str(indice), values=valores)

        atualizar(df)

        def gerar_pdf():
            self.relatorios.relatorio_filtrado(self.df_filtrado, aba)

        def selecionar_item(_event):
            nonlocal item_selecionado
            selecionado = tree.focus()
            if not selecionado:
                return

            item_selecionado = int(selecionado)
            linha = df.iloc[item_selecionado]

            for coluna, campo in campos.items():
                campo.delete(0, "end")
                campo.insert(0, str(linha.get(coluna, "")))

        def mostrar_detalhes(_event):
            selecionado = tree.focus()
            if not selecionado:
                return

            try:
                linha = df.loc[int(selecionado)]
                descricao = linha.get("DESCRICAO", "")
                detalhe = linha.get("DETALHE", "")

                janela = ctk.CTkToplevel(self.app)
                janela.title("Detalhes da Manutenção")
                janela.geometry("400x300")
                janela.transient(self.app)
                janela.grab_set()

                ctk.CTkLabel(janela, text="Descrição", font=("Arial", 16, "bold")).pack(pady=5)
                ctk.CTkLabel(janela, text=descricao, wraplength=350).pack(pady=5)

                ctk.CTkLabel(janela, text="Detalhe", font=("Arial", 16, "bold")).pack(pady=5)
                ctk.CTkLabel(janela, text=detalhe, wraplength=350).pack(pady=5)

                ctk.CTkButton(janela, text="Fechar", command=janela.destroy).pack(pady=10)
            except Exception as erro:
                messagebox.showerror("Erro", str(erro))

        tree.bind("<<TreeviewSelect>>", selecionar_item)
        tree.bind("<Double-1>", mostrar_detalhes)

        def abrir_filtro():
            janela = ctk.CTkToplevel(self.app)
            janela.title("Filtro Avançado")
            janela.geometry("420x500")
            janela.transient(self.app)
            janela.grab_set()

            ctk.CTkLabel(janela, text="FILTRO AVANÇADO", font=("Arial", 18, "bold")).pack(pady=10)

            usar_patrimonio = ctk.BooleanVar()
            usar_categoria = ctk.BooleanVar()
            usar_valor = ctk.BooleanVar()
            usar_data = ctk.BooleanVar()

            frame_patrimonio = ctk.CTkFrame(janela)
            frame_patrimonio.pack(fill="x", pady=5, padx=10)
            ctk.CTkCheckBox(frame_patrimonio, text="Filtrar por Patrimônio", variable=usar_patrimonio).pack(anchor="w")
            patrimonio = ctk.CTkEntry(frame_patrimonio, placeholder_text="Digite a placa")
            patrimonio.pack(fill="x", pady=5)

            frame_categoria = ctk.CTkFrame(janela)
            frame_categoria.pack(fill="x", pady=5, padx=10)
            ctk.CTkCheckBox(frame_categoria, text="Tipo de Manutenção", variable=usar_categoria).pack(anchor="w")
            categoria = ctk.CTkComboBox(frame_categoria, values=CATEGORIAS_MANUTENCAO)
            categoria.pack(fill="x", pady=5)

            frame_valor = ctk.CTkFrame(janela)
            frame_valor.pack(fill="x", pady=5, padx=10)
            ctk.CTkCheckBox(frame_valor, text="Filtrar por Valor", variable=usar_valor).pack(anchor="w")
            valor_min = ctk.CTkEntry(frame_valor, placeholder_text="Valor mínimo")
            valor_min.pack(fill="x", pady=2)
            valor_max = ctk.CTkEntry(frame_valor, placeholder_text="Valor máximo")
            valor_max.pack(fill="x", pady=2)

            frame_data = ctk.CTkFrame(janela)
            frame_data.pack(fill="x", pady=5, padx=10)
            ctk.CTkCheckBox(frame_data, text="Filtrar por Data", variable=usar_data).pack(anchor="w")
            data_inicio = ctk.CTkEntry(frame_data, placeholder_text="Data início (dd/mm/aaaa)")
            data_inicio.pack(fill="x", pady=2)
            data_fim = ctk.CTkEntry(frame_data, placeholder_text="Data fim (dd/mm/aaaa)")
            data_fim.pack(fill="x", pady=2)

            def aplicar_filtro():
                df_filtrado = df.copy()

                if usar_patrimonio.get() and "PATRIMONIO" in df_filtrado.columns:
                    texto = patrimonio.get().strip()
                    if texto:
                        df_filtrado = df_filtrado[
                            df_filtrado["PATRIMONIO"].astype(str).str.contains(texto, case=False, na=False)
                        ]

                if usar_categoria.get() and "CATEGORIA" in df_filtrado.columns:
                    valor_cat = categoria.get().strip().upper()
                    if valor_cat:
                        df_filtrado = df_filtrado[df_filtrado["CATEGORIA"].astype(str).str.upper() == valor_cat]

                if usar_valor.get() and "VALOR" in df_filtrado.columns:
                    serie_valor = pd.to_numeric(df_filtrado["VALOR"], errors="coerce")
                    if valor_min.get():
                        df_filtrado = df_filtrado[serie_valor >= float(valor_min.get())]
                        serie_valor = pd.to_numeric(df_filtrado["VALOR"], errors="coerce")
                    if valor_max.get():
                        df_filtrado = df_filtrado[serie_valor <= float(valor_max.get())]

                if usar_data.get() and "DATA" in df_filtrado.columns:
                    serie_data = pd.to_datetime(df_filtrado["DATA"], errors="coerce", dayfirst=True)
                    if data_inicio.get():
                        inicio = pd.to_datetime(data_inicio.get(), dayfirst=True, errors="coerce")
                        df_filtrado = df_filtrado[serie_data >= inicio]
                        serie_data = pd.to_datetime(df_filtrado["DATA"], errors="coerce", dayfirst=True)
                    if data_fim.get():
                        fim = pd.to_datetime(data_fim.get(), dayfirst=True, errors="coerce")
                        df_filtrado = df_filtrado[serie_data <= fim]

                self.df_filtrado = df_filtrado
                atualizar(df_filtrado)
                janela.destroy()

            ctk.CTkButton(janela, text="Aplicar Filtros", command=aplicar_filtro).pack(pady=15)
            ctk.CTkButton(
                janela,
                text="Limpar Filtro",
                command=lambda: (setattr(self, "df_filtrado", df.copy()), atualizar(df), janela.destroy()),
            ).pack(pady=5)

        def editar():
            nonlocal item_selecionado
            if item_selecionado is None:
                messagebox.showwarning("Aviso", "Selecione um item")
                return

            try:
                df_edit = self.banco.carregar_dataframe(aba).fillna("").astype(str)
                alterou = False

                for coluna, campo in campos.items():
                    novo_valor = campo.get()
                    if str(df_edit.loc[item_selecionado, coluna]) != str(novo_valor):
                        df_edit.loc[item_selecionado, coluna] = str(novo_valor)
                        alterou = True

                if not alterou:
                    messagebox.showinfo("Aviso", "Nenhuma alteração feita")
                    return

                self.banco.escrever_aba(aba, df_edit)
                messagebox.showinfo("Sucesso", "Atualizado!")
                item_selecionado = None
                self.visualizar(aba)
            except Exception as erro:
                messagebox.showerror("Erro", str(erro))

        def excluir():
            nonlocal item_selecionado
            if item_selecionado is None:
                messagebox.showwarning("Aviso", "Selecione um item")
                return
            if not messagebox.askyesno("Confirmação", "Deseja excluir?"):
                return

            try:
                df_edit = self.banco.carregar_dataframe(aba).drop(index=int(item_selecionado))
                self.banco.escrever_aba(aba, df_edit.reset_index(drop=True))
                messagebox.showinfo("Sucesso", "Excluído!")
                self.visualizar(aba)
            except Exception as erro:
                messagebox.showerror("Erro", str(erro))

        botoes = ctk.CTkFrame(container)
        botoes.pack(pady=10)

        ctk.CTkButton(botoes, text="Filtrar", command=abrir_filtro).grid(row=0, column=0, padx=5)
        ctk.CTkButton(botoes, text="PDF", command=gerar_pdf).grid(row=0, column=1, padx=5)
        ctk.CTkButton(botoes, text="Editar", command=editar).grid(row=0, column=2, padx=5)
        ctk.CTkButton(botoes, text="Excluir", command=excluir).grid(row=0, column=3, padx=5)
        ctk.CTkButton(botoes, text="Voltar", command=self.voltar_menu).grid(row=0, column=4, padx=5)

    def tela_veiculos(self):
        self.limpar()
        card = self.criar_card("VEÍCULOS")

        patrimonio = ctk.CTkEntry(card, placeholder_text="Patrimônio")
        patrimonio.pack(pady=5)
        marca = ctk.CTkEntry(card, placeholder_text="Marca")
        marca.pack(pady=5)
        ano = ctk.CTkEntry(card, placeholder_text="Ano")
        ano.pack(pady=5)
        obs = ctk.CTkEntry(card, placeholder_text="Observações")
        obs.pack(pady=5)

        def salvar(_event=None):
            if not patrimonio.get().strip():
                messagebox.showwarning("Aviso", "Patrimônio obrigatório")
                patrimonio.focus()
                return

            if ano.get() and not ano.get().isdigit():
                messagebox.showerror("Erro", "Ano deve ser numérico")
                return

            dados = {
                "PATRIMONIO": patrimonio.get().upper().strip(),
                "HORIMETRO": "",
                "DATA_ATUALIZACAO": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "MARCA": marca.get().upper().strip(),
                "ANO": ano.get().strip(),
                "OBS": obs.get().strip(),
            }

            self.banco.salvar("VEICULOS", dados)
            messagebox.showinfo("Sucesso", "Salvo!")
            limpar_campos([patrimonio, marca, ano, obs])
            patrimonio.focus()

        card.bind("<Return>", salvar)

        ctk.CTkButton(card, text="Salvar", command=salvar).pack(pady=10)
        ctk.CTkButton(card, text="Visualizar", command=lambda: self.visualizar("VEICULOS")).pack(pady=5)
        ctk.CTkButton(card, text="Ver Manutenções", command=lambda: self.visualizar("MANUTENCOES")).pack(pady=5)
        ctk.CTkButton(card, text="Voltar", command=self.voltar_menu).pack(pady=10)

    def tela_manutencao(self):
        self.limpar()
        card = self.criar_card("MANUTENÇÕES")

        patrimonio = ctk.CTkEntry(card, placeholder_text="Patrimônio")
        patrimonio.pack(pady=5)
        descricao = ctk.CTkEntry(card, placeholder_text="Descrição")
        descricao.pack(pady=5)
        data_inicio = ctk.CTkEntry(card, placeholder_text="Data início")
        data_inicio.pack(pady=5)
        data_fim = ctk.CTkEntry(card, placeholder_text="Data fim")
        data_fim.pack(pady=5)
        detalhe = ctk.CTkEntry(card, placeholder_text="Detalhe da manutenção")
        detalhe.pack(pady=5)
        horimetro = ctk.CTkEntry(card, placeholder_text="Horímetro atual")
        horimetro.pack(pady=5)
        horimetro_troca = ctk.CTkEntry(card, placeholder_text="Horímetro para troca")
        horimetro_troca.pack(pady=5)
        categoria = ctk.CTkComboBox(card, values=CATEGORIAS_MANUTENCAO)
        categoria.pack(pady=5)
        categoria.set("GERAL")
        valor = ctk.CTkEntry(card, placeholder_text="Valor")
        valor.pack(pady=5)

        def atualizar_campos(opcao):
            if opcao == "GERAL":
                horimetro.pack(pady=5)
                horimetro_troca.pack(pady=5)
                valor.pack(pady=5)
                data_inicio.pack_forget()
                data_fim.pack_forget()
                return

            horimetro.pack_forget()
            horimetro_troca.pack_forget()
            valor.pack_forget()
            data_inicio.pack(pady=5)
            data_fim.pack(pady=5)

        categoria.configure(command=atualizar_campos)
        atualizar_campos(categoria.get())

        def salvar(_event=None):
            if not patrimonio.get().strip():
                messagebox.showwarning("Aviso", "Patrimônio obrigatório")
                return

            try:
                valor_convertido = float(valor.get()) if valor.get() else 0
                horimetro_atual = float(horimetro.get()) if horimetro.get() else 0
                horimetro_troca_valor = float(horimetro_troca.get()) if horimetro_troca.get() else 0
            except ValueError:
                messagebox.showerror("Erro", "Valores numéricos inválidos")
                return

            situacao_horimetro = "ATRASADO" if horimetro_atual - horimetro_troca_valor >= 0 else "EM DIA"
            data_atual = datetime.now()

            dados = {
                "DATA": data_atual.strftime("%d/%m/%Y %H:%M"),
                "PATRIMONIO": patrimonio.get().strip().upper(),
                "HORIMETRO": horimetro_atual,
                "DESCRICAO": descricao.get().strip(),
                "CATEGORIA": categoria.get().strip().upper(),
                "VALOR": valor_convertido,
                "HORIMETRO_ATUAL": horimetro_atual,
                "HORIMETRO_TROCA": horimetro_troca_valor,
                "SITUACAO_HORIMETRO": situacao_horimetro,
                "SITUACAO_DATA": "OK",
                "DETALHE": detalhe.get().strip(),
                "ANO": data_atual.strftime("%Y"),
                "DATA_INICIO": data_inicio.get().strip(),
                "DATA_FIM": data_fim.get().strip(),
                "MES": data_atual.strftime("%m"),
            }

            self.banco.salvar("MANUTENCOES", dados)
            messagebox.showinfo("Sucesso", "Salvo!")
            limpar_campos([patrimonio, descricao, categoria, valor, horimetro, horimetro_troca, detalhe])
            categoria.set("GERAL")
            atualizar_campos("GERAL")

        card.bind("<Return>", salvar)

        ctk.CTkButton(card, text="Salvar", command=salvar).pack(pady=10)
        ctk.CTkButton(card, text="Visualizar", command=lambda: self.visualizar("MANUTENCOES")).pack(pady=5)
        ctk.CTkButton(card, text="Voltar", command=self.voltar_menu).pack(pady=10)

    def tela_relatorio_veiculo(self):
        self.limpar()
        card = self.criar_card("RELATÓRIO POR VEÍCULO")
        veiculos = self.banco.carregar_veiculos()
        combo = ctk.CTkComboBox(card, values=veiculos)
        combo.pack(pady=10)

        def gerar():
            if not combo.get():
                messagebox.showwarning("Aviso", "Selecione um veículo")
                return
            self.relatorios.relatorio_por_veiculo(combo.get())

        ctk.CTkButton(card, text="Gerar PDF", command=gerar).pack(pady=10)
        ctk.CTkButton(card, text="Voltar", command=self.voltar_menu).pack(pady=10)

    def tela_relatorios_manutencao(self):
        self.limpar()
        card = self.criar_card("RELATÓRIOS DE MANUTENÇÃO")

        for categoria in ["ELETRICA", "MECANICA", "OLEO", "PNEUS"]:
            ctk.CTkButton(
                card,
                text=f"Manutenção {categoria.title()}",
                command=lambda cat=categoria: self.tela_filtro_relatorio(cat),
            ).pack(pady=5)

        ctk.CTkButton(card, text="Relatório Geral", command=self.relatorios.relatorio_manutencoes).pack(pady=10)
        ctk.CTkButton(card, text="Voltar", command=self.voltar_menu).pack(pady=10)

    def tela_filtro_relatorio(self, categoria):
        self.limpar()
        card = self.criar_card(f"RELATÓRIO DE {categoria.upper()}")

        ctk.CTkLabel(card, text="Data Inicial").pack()
        data_inicio = ctk.CTkEntry(card)
        data_inicio.pack(pady=5)

        ctk.CTkLabel(card, text="Data Final").pack()
        data_fim = ctk.CTkEntry(card)
        data_fim.pack(pady=5)

        ctk.CTkLabel(card, text="Veículo").pack()
        combo = ctk.CTkComboBox(card, values=self.banco.carregar_veiculos())
        combo.pack(pady=5)

        def gerar():
            try:
                df = carregar_df_normalizado(self.banco.arquivo, "MANUTENCOES")
                df = df[df["CATEGORIA"].astype(str).str.upper() == categoria.upper()]

                if "DATA" in df.columns:
                    serie_data = pd.to_datetime(df["DATA"], errors="coerce", dayfirst=True)
                    if data_inicio.get():
                        inicio = pd.to_datetime(data_inicio.get(), errors="coerce", dayfirst=True)
                        df = df[serie_data >= inicio]
                        serie_data = pd.to_datetime(df["DATA"], errors="coerce", dayfirst=True)
                    if data_fim.get():
                        fim = pd.to_datetime(data_fim.get(), errors="coerce", dayfirst=True)
                        df = df[serie_data <= fim]

                if combo.get():
                    df = df[df["PATRIMONIO"].astype(str) == combo.get()]

                self.relatorios.relatorio_filtrado(df, categoria)
            except Exception as erro:
                messagebox.showerror("Erro", str(erro))

        ctk.CTkButton(card, text="GERAR RELATÓRIO", command=gerar).pack(pady=10)
        ctk.CTkButton(card, text="CANCELAR", command=self.voltar_menu).pack(pady=5)
