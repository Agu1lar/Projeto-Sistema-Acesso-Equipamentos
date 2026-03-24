from __future__ import annotations

from datetime import datetime
from tkinter import messagebox, ttk

import customtkinter as ctk
import pandas as pd

from Relatorios import Relatorios
from utils.ui import (
    ACCENT,
    BG_APP,
    BG_CARD,
    BG_PANEL,
    BORDER,
    FONT_BODY,
    FONT_HEADING,
    FONT_LABEL,
    FONT_SMALL,
    TEXT,
    TEXT_MUTED,
    criar_cabecalho,
    criar_cartao_info,
    criar_label_form,
    criar_pagina,
    criar_secao,
    adicionar_tooltip,
    estilizar_botao,
    estilizar_combo,
    estilizar_entry,
)

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
    "HORIMETRO": "Horimetro",
    "DATA_ATUALIZACAO": "Atualizado em",
    "CATEGORIA": "Tipo",
    "DATA_INICIO": "Data inicio",
    "DATA_FIM": "Data fim",
    "VALOR": "Valor (R$)",
    "HORIMETRO_ATUAL": "Horimetro atual",
    "HORIMETRO_TROCA": "Horimetro troca",
}


def limpar_campos(campos):
    for campo in campos:
        if isinstance(campo, ctk.CTkEntry):
            campo.delete(0, "end")
        elif isinstance(campo, ctk.CTkComboBox):
            campo.set("")


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

    def _nova_pagina(self, titulo, descricao):
        self.limpar()
        pagina = criar_pagina(self.frame)
        criar_cabecalho(pagina, titulo, descricao)
        return pagina

    def _secao_formulario(self, pagina, titulo, descricao=""):
        _, corpo = criar_secao(pagina, titulo, descricao)
        return corpo

    def _campo(self, parent, row, column, label, widget):
        criar_label_form(parent, label).grid(row=row * 2, column=column, sticky="w", padx=8, pady=(0, 6))
        widget.grid(row=row * 2 + 1, column=column, sticky="ew", padx=8, pady=(0, 12))
        return widget

    def _acoes_horizontal(self, parent, definicoes, largura_botao=150, colunas=0):
        barra = ctk.CTkFrame(parent, fg_color="transparent")
        barra.pack(fill="x", pady=(2, 0))
        for indice, (texto, comando, primario) in enumerate(definicoes):
            linha = indice // colunas if colunas else 0
            coluna = indice % colunas if colunas else indice
            botao = ctk.CTkButton(barra, text=texto, command=comando, width=largura_botao)
            estilizar_botao(botao, primario=primario)
            botao.grid(row=linha, column=coluna, padx=(0, 10), pady=4, sticky="w")
        return barra

    def _colunas_visiveis(self, aba, df):
        desejadas = COLUNAS_EXIBICAO.get(aba, list(df.columns))
        existentes = [coluna for coluna in desejadas if coluna in df.columns]
        return existentes or list(df.columns)

    def visualizar(self, aba):
        item_selecionado = None
        pagina = self._nova_pagina(f"Dados de {aba.title()}", "Consulte, filtre, edite e exporte os registros da base.")
        df = self.banco.carregar_dataframe(aba).fillna("").copy()
        if "MES" in df.columns:
            df = df.drop(columns=["MES"])

        self.df_filtrado = df.copy()
        colunas_existentes = self._colunas_visiveis(aba, df)

        metricas = ctk.CTkFrame(pagina, fg_color="transparent")
        metricas.pack(fill="x", pady=(0, 14))
        metricas.grid_columnconfigure((0, 1, 2), weight=1)
        cards = [
            ("Registros", str(len(df)), TEXT),
            ("Colunas visiveis", str(len(colunas_existentes)), ACCENT),
            ("Filtro ativo", "Nao", TEXT),
        ]
        for indice, (titulo, valor, cor) in enumerate(cards):
            criar_cartao_info(metricas, titulo, valor, cor).grid(row=0, column=indice, sticky="nsew", padx=(0 if indice == 0 else 10, 0))

        conteudo = ctk.CTkFrame(pagina, fg_color="transparent")
        conteudo.pack(fill="both", expand=True)

        lateral = ctk.CTkFrame(conteudo, fg_color="transparent", width=360)
        lateral.pack(side="left", fill="y", padx=(0, 16))
        lateral.pack_propagate(False)

        direita = ctk.CTkFrame(conteudo, fg_color="transparent")
        direita.pack(side="right", fill="both", expand=True)

        _, edicao_corpo = criar_secao(lateral, "Edicao rapida", "Selecione um item na tabela para carregar os campos.")
        edicao_frame = ctk.CTkScrollableFrame(
            edicao_corpo,
            fg_color="transparent",
            corner_radius=0,
            width=0,
            height=430,
        )
        edicao_frame.pack(fill="both", expand=True)
        edicao_frame.grid_columnconfigure((0, 1), weight=1)

        linhas_campos = [
            ["PATRIMONIO", "CATEGORIA"],
            ["ANO", "VALOR"],
            ["HORIMETRO_ATUAL", "HORIMETRO_TROCA"],
            ["SITUACAO_HORIMETRO", "SITUACAO_DATA"],
            ["DATA", "DATA_INICIO"],
            ["DATA_FIM", "DETALHE"],
        ]
        campos = {}

        for row_index, linha in enumerate(linhas_campos):
            for col_index, coluna in enumerate(linha):
                if coluna not in df.columns:
                    continue
                entry = estilizar_entry(ctk.CTkEntry(edicao_frame, placeholder_text=coluna))
                self._campo(edicao_frame, row_index, col_index, NOMES_COLUNAS.get(coluna, coluna.title()), entry)
                campos[coluna] = entry

        _, tabela_corpo = criar_secao(
            direita,
            "Tabela",
            "Duplo clique abre os detalhes completos do registro.",
            expand=True,
        )
        topo_tabela = ctk.CTkFrame(tabela_corpo, fg_color="transparent")
        topo_tabela.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(
            topo_tabela,
            text="Role na horizontal e vertical para ver todos os dados.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
        ).pack(side="left")
        acoes_tabela = ctk.CTkFrame(topo_tabela, fg_color="transparent")
        acoes_tabela.pack(side="right")
        tabela_frame = ctk.CTkFrame(
            tabela_corpo,
            fg_color=BG_CARD,
            corner_radius=16,
            border_width=1,
            border_color=BORDER,
            height=320,
        )
        tabela_frame.pack(fill="both", expand=True)
        tabela_frame.pack_propagate(False)

        grade_tabela = ctk.CTkFrame(tabela_frame, fg_color="transparent")
        grade_tabela.pack(fill="both", expand=True, padx=8, pady=8)
        grade_tabela.grid_rowconfigure(0, weight=1)
        grade_tabela.grid_columnconfigure(0, weight=1)

        tree = ttk.Treeview(grade_tabela, columns=colunas_existentes, show="headings", height=12)
        scroll_y = ctk.CTkScrollbar(grade_tabela, orientation="vertical", command=tree.yview, width=14)
        scroll_x = ctk.CTkScrollbar(grade_tabela, orientation="horizontal", command=tree.xview, height=14)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        scroll_x.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        for coluna in colunas_existentes:
            tree.heading(coluna, text=NOMES_COLUNAS.get(coluna, coluna))
            tree.column(coluna, anchor="center", width=140)

        estado_vazio = ctk.CTkLabel(
            tabela_corpo,
            text="Nenhum registro para exibir.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
        )

        def atualizar(df_filtrado):
            nonlocal item_selecionado
            tree.delete(*tree.get_children())
            for indice, row in df_filtrado.iterrows():
                valores = [str(row.get(coluna, "")) for coluna in colunas_existentes]
                tree.insert("", "end", iid=str(indice), values=valores)

            self.df_filtrado = df_filtrado.copy()
            item_selecionado = None
            cards[2] = ("Filtro ativo", "Sim" if len(df_filtrado) != len(df) else "Nao", TEXT)
            for widget in metricas.winfo_children():
                widget.destroy()
            for indice, (titulo, valor, cor) in enumerate(cards):
                criar_cartao_info(metricas, titulo, valor, cor).grid(
                    row=0, column=indice, sticky="nsew", padx=(0 if indice == 0 else 10, 0)
                )
            if df_filtrado.empty:
                estado_vazio.pack(anchor="w", pady=(8, 0))
            else:
                estado_vazio.pack_forget()

        atualizar(df)

        def gerar_pdf():
            self.relatorios.relatorio_filtrado(self.df_filtrado, aba)

        def limpar_campos_edicao():
            for campo in campos.values():
                campo.delete(0, "end")

        def selecionar_item(_event):
            nonlocal item_selecionado
            selecionado = tree.focus()
            if not selecionado:
                return

            item_selecionado = int(selecionado)
            linha = self.df_filtrado.loc[item_selecionado]

            for coluna, campo in campos.items():
                campo.delete(0, "end")
                campo.insert(0, str(linha.get(coluna, "")))

        def mostrar_detalhes(_event=None, origem_tree=None):
            tree_origem = origem_tree or tree
            selecionado = tree_origem.focus()
            if not selecionado:
                return

            try:
                linha = self.df_filtrado.loc[int(selecionado)]
                janela = ctk.CTkToplevel(self.app, fg_color=BG_APP)
                janela.title("Detalhes da manutencao")
                janela.geometry("520x380")
                janela.transient(self.app)
                janela.grab_set()

                card = ctk.CTkFrame(janela, fg_color=BG_PANEL, corner_radius=18, border_width=1, border_color=BORDER)
                card.pack(fill="both", expand=True, padx=20, pady=20)
                ctk.CTkLabel(card, text="Detalhes do registro", font=FONT_HEADING, text_color=TEXT).pack(
                    anchor="w", padx=18, pady=(18, 8)
                )

                texto = ctk.CTkTextbox(card, fg_color=BG_CARD, corner_radius=12, border_width=1, border_color=BORDER, text_color=TEXT)
                texto.pack(fill="both", expand=True, padx=18, pady=(0, 14))
                for coluna, valor in linha.items():
                    texto.insert("end", f"{NOMES_COLUNAS.get(coluna, coluna)}: {valor}\n")
                texto.configure(state="disabled")

                botao = ctk.CTkButton(card, text="Fechar", command=janela.destroy, width=120)
                estilizar_botao(botao)
                botao.pack(anchor="e", padx=18, pady=(0, 18))
            except Exception as erro:
                messagebox.showerror("Erro", str(erro))

        def abrir_tabela_ampliada():
            janela = ctk.CTkToplevel(self.app, fg_color=BG_APP)
            janela.title(f"Tabela ampliada - {aba.title()}")
            janela.geometry("1200x700")
            janela.transient(self.app)

            pagina_ampliada = criar_pagina(janela)
            criar_cabecalho(pagina_ampliada, "Tabela ampliada", "Visualizacao completa dos registros.")

            quadro = ctk.CTkFrame(
                pagina_ampliada,
                fg_color=BG_PANEL,
                corner_radius=18,
                border_width=1,
                border_color=BORDER,
            )
            quadro.pack(fill="both", expand=True)

            area = ctk.CTkFrame(quadro, fg_color=BG_CARD, corner_radius=16)
            area.pack(fill="both", expand=True, padx=18, pady=18)
            area.grid_rowconfigure(0, weight=1)
            area.grid_columnconfigure(0, weight=1)

            tree_ampliada = ttk.Treeview(area, columns=colunas_existentes, show="headings")
            scroll_y_ampliada = ctk.CTkScrollbar(area, orientation="vertical", command=tree_ampliada.yview, width=14)
            scroll_x_ampliada = ctk.CTkScrollbar(area, orientation="horizontal", command=tree_ampliada.xview, height=14)
            tree_ampliada.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=(8, 0))
            scroll_y_ampliada.grid(row=0, column=1, sticky="ns", padx=(8, 8), pady=(8, 0))
            scroll_x_ampliada.grid(row=1, column=0, sticky="ew", padx=(8, 0), pady=(8, 8))
            tree_ampliada.configure(yscrollcommand=scroll_y_ampliada.set, xscrollcommand=scroll_x_ampliada.set)

            for coluna in colunas_existentes:
                tree_ampliada.heading(coluna, text=NOMES_COLUNAS.get(coluna, coluna))
                tree_ampliada.column(coluna, anchor="center", width=180, minwidth=140)

            for indice, row in self.df_filtrado.iterrows():
                valores = [str(row.get(coluna, "")) for coluna in colunas_existentes]
                tree_ampliada.insert("", "end", iid=str(indice), values=valores)

            tree_ampliada.bind("<Double-1>", lambda _event: mostrar_detalhes(_event, tree_ampliada))

        def abrir_filtro():
            janela = ctk.CTkToplevel(self.app, fg_color=BG_APP)
            janela.title("Filtro avancado")
            janela.geometry("520x560")
            janela.transient(self.app)
            janela.grab_set()

            pagina_filtro = criar_pagina(janela)
            criar_cabecalho(pagina_filtro, "Filtro avancado", "Ative apenas os criterios que deseja aplicar.")

            usar_patrimonio = ctk.BooleanVar()
            usar_categoria = ctk.BooleanVar()
            usar_valor = ctk.BooleanVar()
            usar_data = ctk.BooleanVar()

            _, corpo = criar_secao(pagina_filtro, "Critetios")
            corpo.grid_columnconfigure((0, 1), weight=1)

            frame_patrimonio = ctk.CTkFrame(corpo, fg_color="transparent")
            frame_patrimonio.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
            ctk.CTkCheckBox(frame_patrimonio, text="Patrimonio", variable=usar_patrimonio).pack(anchor="w")
            patrimonio = estilizar_entry(ctk.CTkEntry(frame_patrimonio, placeholder_text="Digite a placa"))
            patrimonio.pack(fill="x", pady=(8, 0))

            frame_categoria = ctk.CTkFrame(corpo, fg_color="transparent")
            frame_categoria.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
            ctk.CTkCheckBox(frame_categoria, text="Categoria", variable=usar_categoria).pack(anchor="w")
            categoria = estilizar_combo(ctk.CTkComboBox(frame_categoria, values=CATEGORIAS_MANUTENCAO))
            categoria.pack(fill="x", pady=(8, 0))

            frame_valor = ctk.CTkFrame(corpo, fg_color="transparent")
            frame_valor.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
            ctk.CTkCheckBox(frame_valor, text="Faixa de valor", variable=usar_valor).pack(anchor="w")
            valor_min = estilizar_entry(ctk.CTkEntry(frame_valor, placeholder_text="Valor minimo"))
            valor_min.pack(fill="x", pady=(8, 8))
            valor_max = estilizar_entry(ctk.CTkEntry(frame_valor, placeholder_text="Valor maximo"))
            valor_max.pack(fill="x")

            frame_data = ctk.CTkFrame(corpo, fg_color="transparent")
            frame_data.grid(row=1, column=1, sticky="nsew", padx=8, pady=8)
            ctk.CTkCheckBox(frame_data, text="Periodo", variable=usar_data).pack(anchor="w")
            data_inicio = estilizar_entry(ctk.CTkEntry(frame_data, placeholder_text="Inicio (dd/mm/aaaa)"))
            data_inicio.pack(fill="x", pady=(8, 8))
            data_fim = estilizar_entry(ctk.CTkEntry(frame_data, placeholder_text="Fim (dd/mm/aaaa)"))
            data_fim.pack(fill="x")

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

                atualizar(df_filtrado)
                limpar_campos_edicao()
                janela.destroy()

            self._acoes_horizontal(
                pagina_filtro,
                [
                    ("Aplicar filtros", aplicar_filtro, True),
                    ("Limpar", lambda: (atualizar(df), limpar_campos_edicao(), janela.destroy()), False),
                ],
            )

        botao_filtrar = ctk.CTkButton(acoes_tabela, text="⚲", command=abrir_filtro, width=36, height=32)
        estilizar_botao(botao_filtrar)
        botao_filtrar.pack(side="left", padx=(0, 8))
        adicionar_tooltip(botao_filtrar, "Filtrar")

        botao_expandir = ctk.CTkButton(acoes_tabela, text="⛶", command=abrir_tabela_ampliada, width=36, height=32)
        estilizar_botao(botao_expandir)
        botao_expandir.pack(side="left")
        adicionar_tooltip(botao_expandir, "Ampliar")

        tree.bind("<<TreeviewSelect>>", selecionar_item)
        tree.bind("<Double-1>", mostrar_detalhes)

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
                    messagebox.showinfo("Aviso", "Nenhuma alteracao feita")
                    return

                self.banco.escrever_aba(aba, df_edit)
                messagebox.showinfo("Sucesso", "Registro atualizado com sucesso.")
                item_selecionado = None
                self.visualizar(aba)
            except Exception as erro:
                messagebox.showerror("Erro", str(erro))

        def excluir():
            nonlocal item_selecionado
            if item_selecionado is None:
                messagebox.showwarning("Aviso", "Selecione um item")
                return
            if not messagebox.askyesno("Confirmacao", "Deseja excluir o registro selecionado?"):
                return

            try:
                df_edit = self.banco.carregar_dataframe(aba).drop(index=int(item_selecionado))
                self.banco.escrever_aba(aba, df_edit.reset_index(drop=True))
                messagebox.showinfo("Sucesso", "Registro excluido.")
                self.visualizar(aba)
            except Exception as erro:
                messagebox.showerror("Erro", str(erro))

        self._acoes_horizontal(
            lateral,
            [
                ("Gerar PDF", gerar_pdf, True),
                ("Salvar", editar, False),
                ("Excluir", excluir, False),
                ("Voltar", self.voltar_menu, False),
            ],
            largura_botao=150 if aba == "MANUTENCOES" else 130,
            colunas=2,
        )


    def tela_veiculos(self):
        pagina = self._nova_pagina("Cadastro de veiculos", "Registre novos ativos e mantenha os dados principais atualizados.")
        corpo = self._secao_formulario(pagina, "Dados do veiculo")
        corpo.grid_columnconfigure((0, 1), weight=1)

        patrimonio = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="Ex.: ABC1234"))
        marca = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="Marca do veiculo"))
        ano = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="Ano"))
        obs = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="Observacoes relevantes"))

        self._campo(corpo, 0, 0, "Patrimonio", patrimonio)
        self._campo(corpo, 0, 1, "Marca", marca)
        self._campo(corpo, 1, 0, "Ano", ano)
        self._campo(corpo, 1, 1, "Observacoes", obs)

        def salvar(_event=None):
            if not patrimonio.get().strip():
                messagebox.showwarning("Aviso", "Patrimonio obrigatorio")
                patrimonio.focus()
                return

            if ano.get() and not ano.get().isdigit():
                messagebox.showerror("Erro", "Ano deve ser numerico")
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
            messagebox.showinfo("Sucesso", "Veiculo salvo com sucesso.")
            limpar_campos([patrimonio, marca, ano, obs])
            patrimonio.focus()

        pagina.bind("<Return>", salvar)

        self._acoes_horizontal(
            pagina,
            [
                ("Salvar", salvar, True),
                ("Visualizar base", lambda: self.visualizar("VEICULOS"), False),
                ("Ver manutencoes", lambda: self.visualizar("MANUTENCOES"), False),
                ("Voltar", self.voltar_menu, False),
            ],
        )

    def tela_manutencao(self):
        pagina = self._nova_pagina("Cadastro de manutencao", "Lance manutencoes preventivas e corretivas sem depender de edicao manual na planilha.")
        corpo = self._secao_formulario(pagina, "Dados da manutencao")
        corpo.grid_columnconfigure((0, 1, 2), weight=1)

        patrimonio = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="Patrimonio"))
        descricao = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="Descricao"))
        categoria = estilizar_combo(ctk.CTkComboBox(corpo, values=CATEGORIAS_MANUTENCAO))
        categoria.set("GERAL")
        valor = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="0,00"))
        horimetro = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="Horimetro atual"))
        horimetro_troca = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="Horimetro para troca"))
        data_inicio = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="dd/mm/aaaa"))
        data_fim = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="dd/mm/aaaa"))
        detalhe = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="Detalhamento da atividade"))

        self._campo(corpo, 0, 0, "Patrimonio", patrimonio)
        self._campo(corpo, 0, 1, "Descricao", descricao)
        self._campo(corpo, 0, 2, "Categoria", categoria)
        self._campo(corpo, 1, 0, "Valor", valor)
        self._campo(corpo, 1, 1, "Horimetro atual", horimetro)
        self._campo(corpo, 1, 2, "Horimetro de troca", horimetro_troca)
        self._campo(corpo, 2, 0, "Data inicio", data_inicio)
        self._campo(corpo, 2, 1, "Data fim", data_fim)
        self._campo(corpo, 2, 2, "Detalhe", detalhe)

        def atualizar_campos(opcao):
            manutencao_geral = opcao == "GERAL"
            for widget in [horimetro, horimetro_troca, valor]:
                if manutencao_geral:
                    widget.configure(state="normal")
                else:
                    widget.delete(0, "end")
                    widget.configure(state="disabled")

            for widget in [data_inicio, data_fim]:
                if manutencao_geral:
                    widget.delete(0, "end")
                    widget.configure(state="disabled")
                else:
                    widget.configure(state="normal")

        categoria.configure(command=atualizar_campos)
        atualizar_campos(categoria.get())

        def salvar(_event=None):
            if not patrimonio.get().strip():
                messagebox.showwarning("Aviso", "Patrimonio obrigatorio")
                return

            try:
                valor_convertido = float(valor.get().replace(",", ".")) if valor.get() else 0
                horimetro_atual = float(horimetro.get().replace(",", ".")) if horimetro.get() else 0
                horimetro_troca_valor = float(horimetro_troca.get().replace(",", ".")) if horimetro_troca.get() else 0
            except ValueError:
                messagebox.showerror("Erro", "Valores numericos invalidos")
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
            messagebox.showinfo("Sucesso", "Manutencao salva com sucesso.")
            limpar_campos([patrimonio, descricao, categoria, valor, horimetro, horimetro_troca, detalhe, data_inicio, data_fim])
            categoria.set("GERAL")
            atualizar_campos("GERAL")

        pagina.bind("<Return>", salvar)

        self._acoes_horizontal(
            pagina,
            [
                ("Salvar", salvar, True),
                ("Visualizar base", lambda: self.visualizar("MANUTENCOES"), False),
                ("Voltar", self.voltar_menu, False),
            ],
        )

    def tela_relatorios_manutencao(self):
        pagina = self._nova_pagina("Relatorios de manutencao", "Gere relatorios por categoria ou um consolidado geral.")

        linha = ctk.CTkFrame(pagina, fg_color="transparent")
        linha.pack(fill="both", expand=True)
        linha.grid_columnconfigure((0, 1), weight=1)

        categorias = [("Eletrica", "ELETRICA"), ("Mecanica", "MECANICA"), ("Oleo", "OLEO"), ("Pneus", "PNEUS")]
        for indice, (titulo, categoria) in enumerate(categorias):
            card = ctk.CTkFrame(linha, fg_color=BG_PANEL, corner_radius=18, border_width=1, border_color=BORDER)
            card.grid(row=indice // 2, column=indice % 2, sticky="nsew", padx=8, pady=8)
            ctk.CTkLabel(card, text=titulo, font=FONT_HEADING, text_color=TEXT).pack(anchor="w", padx=16, pady=(16, 6))
            ctk.CTkLabel(
                card,
                text=f"Filtre o periodo e o veiculo para gerar o relatorio de {titulo.lower()}.",
                font=FONT_SMALL,
                text_color=TEXT_MUTED,
                wraplength=320,
                justify="left",
            ).pack(anchor="w", padx=16)
            botao = ctk.CTkButton(card, text="Abrir filtro", command=lambda cat=categoria: self.tela_filtro_relatorio(cat), width=150)
            estilizar_botao(botao)
            botao.pack(anchor="w", padx=16, pady=16)

        self._acoes_horizontal(
            pagina,
            [("Relatorio geral", self.relatorios.relatorio_manutencoes, True), ("Voltar", self.voltar_menu, False)],
        )

    def tela_filtro_relatorio(self, categoria):
        pagina = self._nova_pagina(f"Relatorio de {categoria.lower()}", "Defina os filtros antes de exportar o PDF.")
        corpo = self._secao_formulario(pagina, "Filtros")
        corpo.grid_columnconfigure((0, 1, 2), weight=1)

        data_inicio = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="dd/mm/aaaa"))
        data_fim = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text="dd/mm/aaaa"))
        combo = estilizar_combo(ctk.CTkComboBox(corpo, values=self.banco.carregar_veiculos()))

        self._campo(corpo, 0, 0, "Data inicial", data_inicio)
        self._campo(corpo, 0, 1, "Data final", data_fim)
        self._campo(corpo, 0, 2, "Veiculo", combo)

        def gerar():
            try:
                df = self.banco.carregar_dataframe("MANUTENCOES").fillna("").copy()
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

        self._acoes_horizontal(
            pagina,
            [("Gerar relatorio", gerar, True), ("Voltar", self.voltar_menu, False)],
        )
