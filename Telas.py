from __future__ import annotations

import os
from pathlib import Path
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import pandas as pd

from Relatorios import Relatorios
from services.service_automacoes import (
    cadastrar_veiculo_automatico,
    listar_inconsistencias,
    resumo_veiculo,
    atualizar_veiculo_por_manutencao,
)
from services.service_carteirinhas import MODELO_CARTEIRINHA_VERSAO, gerar_carteirinha_treinamento
from services.service_certificados import (
    STATUS_DATA_INVALIDA,
    STATUS_PENDENTE,
    STATUS_VENCIDO,
    atualizar_status_certificados_df,
    gerar_certificado_word,
    gerar_relatorio_certificados,
    listar_certificados_pendentes,
    preparar_certificado,
)
from services.service_mobile_sync import (
    HOST_PADRAO,
    PORTA_PADRAO,
    concluir_pendencia_mobile,
    excluir_pendencia_mobile,
    iniciar_servidor_mobile,
    listar_pendencias_mobile,
    obter_pendencia_mobile,
    obter_url_mobile,
    parar_servidor_mobile,
    servidor_mobile_ativo,
)
from utils.caminhos import caminho_mobile_importados, caminho_mobile_pendencias
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
    SUCCESS,
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
COR_ALERTA = "#ef4444"

COLUNAS_CARTEIRINHA = [
    "CODIGO",
    "NOME",
    "EMPRESA",
    "FUNCAO",
    "TREINAMENTO",
    "DATA_EMISSAO",
    "VALIDADE",
]

COLUNAS_CERTIFICADOS = [
    "TIPO_CERTIFICADO",
    "NOME",
    "CPF",
    "CARGA_HORARIA",
    "DATA_EMISSAO",
    "DATA_VENCIMENTO",
    "STATUS",
    "DIAS_RESTANTES",
]

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

        painel_auto = ctk.CTkFrame(corpo, fg_color=BG_CARD, corner_radius=16, border_width=1, border_color=BORDER)
        painel_auto.grid(row=6, column=0, columnspan=3, sticky="ew", padx=8, pady=(4, 12))
        ctk.CTkLabel(painel_auto, text="Resumo automatico por patrimonio", font=FONT_LABEL, text_color=TEXT).pack(
            anchor="w", padx=14, pady=(12, 4)
        )
        resumo_label = ctk.CTkLabel(
            painel_auto,
            text="Informe o patrimonio para carregar o resumo do veiculo e o historico.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            wraplength=860,
            justify="left",
        )
        resumo_label.pack(anchor="w", padx=14, pady=(0, 12))

        def atualizar_resumo_patrimonio(_event=None):
            resumo = resumo_veiculo(self.banco, patrimonio.get())
            if not resumo:
                resumo_label.configure(text="Informe o patrimonio para carregar o resumo do veiculo e o historico.")
                return

            if resumo.get("horimetro_base") and not horimetro.get().strip():
                horimetro.insert(0, str(resumo["horimetro_base"]))

            texto = (
                f"Cadastro: {'sim' if resumo['tem_cadastro'] else 'nao'} | "
                f"Marca: {resumo['marca'] or '-'} | Ano: {resumo['ano'] or '-'} | "
                f"Manutencoes: {resumo['manutencoes']} | Gasto total: R$ {resumo['gasto_total']:.2f} | "
                f"Ultima categoria: {resumo['ultima_categoria'] or '-'} | "
                f"Ultimo horimetro: {resumo['ultimo_horimetro'] or resumo['horimetro_base'] or '-'}"
            )
            resumo_label.configure(text=texto)

        patrimonio.bind("<FocusOut>", atualizar_resumo_patrimonio)

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

            resumo = resumo_veiculo(self.banco, patrimonio.get())
            if resumo and resumo["tem_cadastro"] and resumo.get("horimetro_base"):
                try:
                    horimetro_base = float(str(resumo["horimetro_base"]).replace(",", ".") or 0)
                    if horimetro_atual < horimetro_base:
                        if not messagebox.askyesno(
                            "Inconsistencia",
                            "O horimetro informado esta abaixo do valor registrado no veiculo. Deseja salvar mesmo assim?",
                        ):
                            return
                except Exception:
                    pass

            if not resumo or not resumo["tem_cadastro"]:
                if messagebox.askyesno(
                    "Cadastro automatico",
                    "Esse patrimonio nao existe em VEICULOS. Deseja cadastrar o veiculo automaticamente?",
                ):
                    cadastrar_veiculo_automatico(
                        self.banco,
                        patrimonio.get(),
                        horimetro_atual,
                        f"Criado automaticamente por manutencao em {data_atual:%d/%m/%Y %H:%M}",
                    )

            self.banco.salvar("MANUTENCOES", dados)
            atualizar_veiculo_por_manutencao(self.banco, patrimonio.get(), horimetro_atual)
            messagebox.showinfo("Sucesso", "Manutencao salva com sucesso.")
            limpar_campos([patrimonio, descricao, categoria, valor, horimetro, horimetro_troca, detalhe, data_inicio, data_fim])
            categoria.set("GERAL")
            atualizar_campos("GERAL")
            resumo_label.configure(text="Informe o patrimonio para carregar o resumo do veiculo e o historico.")

        pagina.bind("<Return>", salvar)

        self._acoes_horizontal(
            pagina,
            [
                ("Salvar", salvar, True),
                ("Visualizar base", lambda: self.visualizar("MANUTENCOES"), False),
                ("Coleta mobile", self.tela_coleta_mobile, False),
                ("Voltar", self.voltar_menu, False),
            ],
        )

    def tela_coleta_mobile(self):
        pagina = self._nova_pagina(
            "Coleta mobile",
            "Receba manutencoes e fotos enviadas do celular na mesma rede e importe para a base quando desejar.",
        )

        topo = ctk.CTkFrame(pagina, fg_color="transparent")
        topo.pack(fill="x", pady=(0, 16))
        topo.grid_columnconfigure((0, 1, 2), weight=1)

        status_url = ctk.StringVar(value="Servidor mobile parado.")
        pendencias = {"dados": []}

        for indice, (titulo, valor, cor) in enumerate(
            [
                ("Servidor", "Parado", TEXT),
                ("Porta", str(PORTA_PADRAO), ACCENT),
                ("Pendencias", "0", TEXT),
            ]
        ):
            criar_cartao_info(topo, titulo, valor, cor).grid(row=0, column=indice, sticky="nsew", padx=(0 if indice == 0 else 10, 0))

        conteudo = ctk.CTkFrame(pagina, fg_color="transparent")
        conteudo.pack(fill="both", expand=True)

        lateral = ctk.CTkFrame(conteudo, fg_color="transparent", width=390)
        lateral.pack(side="left", fill="y", padx=(0, 16))
        lateral.pack_propagate(False)

        direita = ctk.CTkFrame(conteudo, fg_color="transparent")
        direita.pack(side="right", fill="both", expand=True)

        _, corpo = criar_secao(
            lateral,
            "Servidor e descarga",
            "No celular, abra o link da rede local, preencha os dados e envie. Depois use a importacao no desktop.",
            expand=True,
        )
        ctk.CTkLabel(corpo, textvariable=status_url, font=FONT_SMALL, text_color=TEXT_MUTED, wraplength=320, justify="left").pack(
            anchor="w", pady=(0, 12)
        )
        ctk.CTkLabel(
            corpo,
            text="Requisito: celular e computador precisam estar no mesmo Wi-Fi. O servidor usa o endereco local do PC.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            wraplength=320,
            justify="left",
        ).pack(anchor="w", pady=(0, 16))

        _, tabela_corpo = criar_secao(
            direita,
            "Pendencias recebidas",
            "Revise as coletas do celular e importe individualmente ou em lote para a planilha.",
            expand=True,
        )
        tabela_frame = ctk.CTkFrame(
            tabela_corpo,
            fg_color=BG_CARD,
            corner_radius=16,
            border_width=1,
            border_color=BORDER,
            height=420,
        )
        tabela_frame.pack(fill="both", expand=True)
        tabela_frame.pack_propagate(False)

        grade = ctk.CTkFrame(tabela_frame, fg_color="transparent")
        grade.pack(fill="both", expand=True, padx=8, pady=8)
        grade.grid_rowconfigure(0, weight=1)
        grade.grid_columnconfigure(0, weight=1)

        colunas = ("id", "recebido", "patrimonio", "categoria", "fotos", "origem")
        tree = ttk.Treeview(grade, columns=colunas, show="headings", height=12)
        scroll_y = ctk.CTkScrollbar(grade, orientation="vertical", command=tree.yview, width=14)
        scroll_x = ctk.CTkScrollbar(grade, orientation="horizontal", command=tree.xview, height=14)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        scroll_x.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        definicoes = [
            ("id", "ID", 180),
            ("recebido", "Recebido em", 150),
            ("patrimonio", "Patrimonios", 180),
            ("categoria", "Itens", 70),
            ("fotos", "Fotos", 70),
            ("origem", "Origem", 110),
        ]
        for coluna, titulo, largura in definicoes:
            tree.heading(coluna, text=titulo)
            tree.column(coluna, width=largura, anchor="center")

        def _atualizar_cards(servidor_txt, total_pendencias):
            for widget in topo.winfo_children():
                widget.destroy()
            cards = [
                ("Servidor", servidor_txt, TEXT if servidor_txt == "Parado" else SUCCESS),
                ("Porta", str(PORTA_PADRAO), ACCENT),
                ("Pendencias", str(total_pendencias), TEXT),
            ]
            for indice, (titulo, valor, cor) in enumerate(cards):
                criar_cartao_info(topo, titulo, valor, cor).grid(
                    row=0, column=indice, sticky="nsew", padx=(0 if indice == 0 else 10, 0)
                )

        def atualizar_lista():
            pendencias["dados"] = listar_pendencias_mobile()
            tree.delete(*tree.get_children())
            for item in pendencias["dados"]:
                itens = item.get("itens", [])
                primeiro = itens[0].get("dados", {}) if itens else {}
                patrimonios = ", ".join(
                    filtro for filtro in [it.get("dados", {}).get("patrimonio", "").strip() for it in itens[:3]] if filtro
                )
                if len(itens) > 3:
                    patrimonios += "..."
                tree.insert(
                    "",
                    "end",
                    iid=item["id"],
                    values=(
                        item["id"],
                        item.get("recebido_em", ""),
                        patrimonios or primeiro.get("patrimonio", ""),
                        str(len(itens)),
                        str(sum(len(it.get("fotos", [])) for it in itens)),
                        item.get("origem", ""),
                    ),
                )

            if servidor_mobile_ativo():
                status_url.set(f"Servidor ativo em {obter_url_mobile(PORTA_PADRAO)}")
                _atualizar_cards("Ativo", len(pendencias["dados"]))
            else:
                status_url.set("Servidor mobile parado.")
                _atualizar_cards("Parado", len(pendencias["dados"]))

        def iniciar():
            try:
                url = iniciar_servidor_mobile(PORTA_PADRAO)
                status_url.set(f"Servidor ativo em {url}")
                atualizar_lista()
                messagebox.showinfo("Servidor iniciado", f"Acesse no celular:\n{url}")
            except Exception as erro:
                messagebox.showerror("Erro", f"Nao foi possivel iniciar o servidor mobile:\n{erro}")

        def parar():
            parar_servidor_mobile()
            atualizar_lista()

        def copiar_link():
            url = obter_url_mobile(PORTA_PADRAO)
            self.app.clipboard_clear()
            self.app.clipboard_append(url)
            messagebox.showinfo("Copiado", f"Link copiado para a area de transferencia:\n{url}")

        def _dados_manutencao_mobile(dados, submission_id, destino_item):
            detalhe_base = dados.get("detalhe", "").strip()
            detalhe_mobile = f"Coleta mobile: {submission_id}"
            if destino_item.exists():
                detalhe_mobile += f" | Fotos: {destino_item}"

            detalhes = " | ".join(parte for parte in [detalhe_base, detalhe_mobile] if parte)

            def numero(valor):
                texto = str(valor or "").strip().replace(",", ".")
                try:
                    return float(texto) if texto else 0.0
                except ValueError:
                    return 0.0

            data_atual = datetime.now()
            horimetro_atual = numero(dados.get("horimetro_atual"))
            horimetro_troca = numero(dados.get("horimetro_troca"))

            return {
                "DATA": data_atual.strftime("%d/%m/%Y %H:%M"),
                "PATRIMONIO": dados.get("patrimonio", "").strip().upper(),
                "HORIMETRO": horimetro_atual,
                "DESCRICAO": dados.get("descricao", "").strip(),
                "CATEGORIA": dados.get("categoria", "GERAL").strip().upper() or "GERAL",
                "VALOR": numero(dados.get("valor")),
                "HORIMETRO_ATUAL": horimetro_atual,
                "HORIMETRO_TROCA": horimetro_troca,
                "SITUACAO_HORIMETRO": "ATRASADO" if horimetro_atual - horimetro_troca >= 0 else "EM DIA",
                "SITUACAO_DATA": "OK",
                "DETALHE": detalhes,
                "ANO": data_atual.strftime("%Y"),
                "DATA_INICIO": dados.get("data_inicio", "").strip(),
                "DATA_FIM": dados.get("data_fim", "").strip(),
                "MES": data_atual.strftime("%m"),
            }

        def _dados_veiculo_mobile(dados, submission_id, destino_item):
            detalhes = []
            if dados.get("descricao", "").strip():
                detalhes.append(dados.get("descricao", "").strip())
            if dados.get("detalhe", "").strip():
                detalhes.append(dados.get("detalhe", "").strip())
            if destino_item.exists():
                detalhes.append(f"Coleta mobile: {submission_id} | Fotos: {destino_item}")

            return {
                "PATRIMONIO": dados.get("patrimonio", "").strip().upper(),
                "HORIMETRO": str(dados.get("horimetro_atual", "")).strip(),
                "DATA_ATUALIZACAO": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "MARCA": dados.get("marca", "").strip().upper(),
                "ANO": str(dados.get("ano", "")).strip(),
                "OBS": " | ".join(parte for parte in detalhes if parte),
            }

        def mostrar_detalhes_mobile(_event=None):
            selecionado = tree.focus()
            if not selecionado:
                return

            payload = obter_pendencia_mobile(selecionado)
            if not payload:
                return

            janela = ctk.CTkToplevel(self.app, fg_color=BG_APP)
            janela.title("Detalhes da coleta")
            janela.geometry("760x520")
            janela.transient(self.app)
            janela.grab_set()

            pagina_detalhe = criar_pagina(janela)
            criar_cabecalho(pagina_detalhe, "Detalhes da coleta", f"ID {payload.get('id', '')}")

            texto = ctk.CTkTextbox(
                pagina_detalhe,
                fg_color=BG_CARD,
                corner_radius=16,
                border_width=1,
                border_color=BORDER,
                text_color=TEXT,
            )
            texto.pack(fill="both", expand=True)

            texto.insert("end", f"Recebido em: {payload.get('recebido_em', '')}\n")
            texto.insert("end", f"Origem: {payload.get('origem', '')}\n")
            texto.insert("end", f"Itens no lote: {len(payload.get('itens', []))}\n\n")
            for indice_item, item in enumerate(payload.get("itens", []), start=1):
                dados = item.get("dados", {})
                texto.insert("end", f"[Item {indice_item}]\n")
                for chave, valor in dados.items():
                    texto.insert("end", f"{chave}: {valor}\n")
                texto.insert("end", f"Fotos: {len(item.get('fotos', []))}\n\n")
            texto.configure(state="disabled")

        def importar_ids(ids: list[str]):
            if not ids:
                messagebox.showwarning("Aviso", "Nenhuma coleta selecionada para importar.")
                return

            importadas = 0
            falhas = []
            for submission_id in ids:
                payload = obter_pendencia_mobile(submission_id)
                if not payload:
                    continue
                try:
                    itens = payload.get("itens", [])
                    if not itens:
                        raise ValueError("Nenhum item encontrado na coleta.")
                    destino_base = Path(caminho_mobile_importados()) / submission_id
                    for indice_item, item in enumerate(itens, start=1):
                        dados_item = item.get("dados", {})
                        destino = str(dados_item.get("destino", "MANUTENCAO")).strip().upper() or "MANUTENCAO"
                        pasta_fotos = destino_base / f"item_{indice_item:02d}" / "fotos"
                        if destino == "VEICULO":
                            registro = _dados_veiculo_mobile(dados_item, submission_id, pasta_fotos)
                            if not registro["PATRIMONIO"]:
                                raise ValueError(f"Item {indice_item} sem patrimonio.")
                            self.banco.salvar("VEICULOS", registro)
                        else:
                            registro = _dados_manutencao_mobile(dados_item, submission_id, pasta_fotos)
                            if not registro["PATRIMONIO"] or not registro["DESCRICAO"]:
                                raise ValueError(f"Item {indice_item} sem patrimonio ou descricao.")
                            resumo = resumo_veiculo(self.banco, registro["PATRIMONIO"])
                            if not resumo or not resumo["tem_cadastro"]:
                                cadastrar_veiculo_automatico(
                                    self.banco,
                                    registro["PATRIMONIO"],
                                    registro["HORIMETRO_ATUAL"],
                                    f"Criado automaticamente por coleta mobile {submission_id}",
                                )
                            self.banco.salvar("MANUTENCOES", registro)
                            atualizar_veiculo_por_manutencao(self.banco, registro["PATRIMONIO"], registro["HORIMETRO_ATUAL"])
                    concluir_pendencia_mobile(submission_id)
                    importadas += len(itens)
                except Exception as erro:
                    falhas.append(f"{submission_id}: {erro}")

            atualizar_lista()
            if falhas:
                messagebox.showwarning(
                    "Importacao parcial",
                    f"{importadas} manutencao(oes) importada(s).\n\nFalhas:\n" + "\n".join(falhas[:5]),
                )
            else:
                messagebox.showinfo("Sucesso", f"{importadas} manutencao(oes) importada(s) com sucesso.")

        def importar_selecionada():
            selecionado = tree.focus()
            if not selecionado:
                messagebox.showwarning("Aviso", "Selecione uma coleta.")
                return
            importar_ids([selecionado])

        def importar_todas():
            importar_ids([item["id"] for item in pendencias["dados"]])

        def excluir_selecionada():
            selecionado = tree.focus()
            if not selecionado:
                messagebox.showwarning("Aviso", "Selecione uma coleta para excluir.")
                return
            if not messagebox.askyesno("Confirmacao", "Deseja excluir a coleta pendente selecionada?"):
                return
            excluir_pendencia_mobile(selecionado)
            atualizar_lista()

        def excluir_todas():
            if not pendencias["dados"]:
                messagebox.showwarning("Aviso", "Nao ha coletas pendentes.")
                return
            if not messagebox.askyesno("Confirmacao", "Deseja excluir todas as coletas pendentes?"):
                return
            for item in list(pendencias["dados"]):
                excluir_pendencia_mobile(item["id"])
            atualizar_lista()

        def abrir_pasta_pendencias():
            os.startfile(caminho_mobile_pendencias())

        def abrir_pasta_importados():
            os.startfile(caminho_mobile_importados())

        self._acoes_horizontal(
            corpo,
            [
                ("Iniciar servidor", iniciar, True),
                ("Parar servidor", parar, False),
                ("Copiar link", copiar_link, False),
                ("Pasta pendente", abrir_pasta_pendencias, False),
                ("Pasta importada", abrir_pasta_importados, False),
                ("Voltar", self.voltar_menu, False),
            ],
            largura_botao=150,
            colunas=2,
        )

        self._acoes_horizontal(
            tabela_corpo,
            [
                ("Atualizar fila", atualizar_lista, False),
                ("Importar selecionada", importar_selecionada, True),
                ("Importar todas", importar_todas, False),
                ("Excluir selecionada", excluir_selecionada, False),
                ("Excluir todas", excluir_todas, False),
            ],
            largura_botao=170,
            colunas=2,
        )

        tree.bind("<Double-1>", mostrar_detalhes_mobile)
        atualizar_lista()

    def tela_fila_unica(self):
        pagina = self._nova_pagina(
            "Fila unica",
            "Acompanhe pendencias mobile, inconsistencias da base e atalhos para resolver tudo no mesmo lugar.",
        )

        pendencias_mobile = listar_pendencias_mobile()
        inconsistencias = listar_inconsistencias(self.banco)
        certificados_pendentes = listar_certificados_pendentes(self.banco)

        metricas = ctk.CTkFrame(pagina, fg_color="transparent")
        metricas.pack(fill="x", pady=(0, 16))
        metricas.grid_columnconfigure((0, 1, 2), weight=1)
        cards = [
            ("Pendencias mobile", str(len(pendencias_mobile)), ACCENT),
            ("Inconsistencias", str(len(inconsistencias)), TEXT),
            ("Certificados pendentes", str(len(certificados_pendentes)), COR_ALERTA if len(certificados_pendentes) else SUCCESS),
        ]
        for indice, (titulo, valor, cor) in enumerate(cards):
            criar_cartao_info(metricas, titulo, valor, cor).grid(
                row=0, column=indice, sticky="nsew", padx=(0 if indice == 0 else 10, 0)
            )

        conteudo = ctk.CTkFrame(pagina, fg_color="transparent")
        conteudo.pack(fill="both", expand=True)

        esquerda = ctk.CTkFrame(conteudo, fg_color="transparent")
        esquerda.pack(side="left", fill="both", expand=True, padx=(0, 10))
        direita = ctk.CTkFrame(conteudo, fg_color="transparent")
        direita.pack(side="right", fill="both", expand=True, padx=(10, 0))

        _, corpo_mobile = criar_secao(
            esquerda,
            "Pendencias mobile",
            "Entradas recebidas do celular aguardando importacao.",
            expand=True,
        )
        texto_mobile = ctk.CTkTextbox(corpo_mobile, fg_color=BG_CARD, text_color=TEXT, border_width=1, border_color=BORDER)
        texto_mobile.pack(fill="both", expand=True)
        for item in pendencias_mobile[:20]:
            texto_mobile.insert(
                "end",
                f"{item.get('recebido_em', '')} | {item.get('id', '')} | itens: {len(item.get('itens', []))}\n",
            )
        texto_mobile.configure(state="disabled")

        _, corpo_incons = criar_secao(
            direita,
            "Inconsistencias",
            "Cadastros incompletos, duplicidades e manutencoes sem veiculo.",
            expand=True,
        )
        texto_incons = ctk.CTkTextbox(corpo_incons, fg_color=BG_CARD, text_color=TEXT, border_width=1, border_color=BORDER)
        texto_incons.pack(fill="both", expand=True)
        for item in inconsistencias[:40]:
            texto_incons.insert("end", f"{item['origem']} | {item['tipo']} | {item['referencia']}\n")
        texto_incons.configure(state="disabled")

        _, corpo_certificados = criar_secao(
            direita,
            "Certificados a renovar",
            "Certificados vencidos, com vencimento proximo ou com data invalida.",
            expand=True,
        )
        texto_certificados = ctk.CTkTextbox(corpo_certificados, fg_color=BG_CARD, text_color=TEXT, border_width=1, border_color=BORDER)
        texto_certificados.pack(fill="both", expand=True)
        for _, linha in certificados_pendentes.head(40).iterrows():
            texto_certificados.insert(
                "end",
                f"{linha.get('NOME', '')} | {linha.get('TIPO_CERTIFICADO', '')} | {linha.get('DATA_VENCIMENTO', '')} | {linha.get('STATUS', '')}\n",
            )
        texto_certificados.configure(state="disabled")

        self._acoes_horizontal(
            pagina,
            [
                ("Abrir coleta mobile", self.tela_coleta_mobile, True),
                ("Abrir certificados", self.tela_certificados_funcionarios, False),
                ("Ver veiculos", lambda: self.visualizar("VEICULOS"), False),
                ("Ver manutencoes", lambda: self.visualizar("MANUTENCOES"), False),
                ("Voltar", self.voltar_menu, False),
            ],
        )

    def tela_certificados_funcionarios(self):
        pagina = self._nova_pagina(
            "Certificados de funcionarios",
            "Cadastre certificados internos, acompanhe vencimentos e destaque automaticamente os que precisam de renovacao.",
        )

        conteudo = ctk.CTkFrame(pagina, fg_color="transparent")
        conteudo.pack(fill="both", expand=True)

        lateral = ctk.CTkFrame(conteudo, fg_color="transparent", width=430)
        lateral.pack(side="left", fill="y", padx=(0, 16))
        lateral.pack_propagate(False)

        direita = ctk.CTkFrame(conteudo, fg_color="transparent")
        direita.pack(side="right", fill="both", expand=True)

        _, corpo_root = criar_secao(
            lateral,
            "Dados do certificado",
            "Os certificados a vencer em ate 30 dias ficam como pendentes para renovacao.",
            expand=True,
        )
        corpo = ctk.CTkScrollableFrame(corpo_root, fg_color="transparent", corner_radius=0)
        corpo.pack(fill="both", expand=True)
        corpo.grid_columnconfigure((0, 1), weight=1)

        campos = {}
        configuracao_campos = [
            ("Tipo de certificado", "TIPO_CERTIFICADO", "Ex.: NR 35"),
            ("Nome do funcionario", "NOME", "Nome completo"),
            ("CPF", "CPF", "000.000.000-00"),
            ("Carga horaria", "CARGA_HORARIA", "Ex.: 08 horas"),
            ("Periodo de emissao", "DATA_EMISSAO", "Ex.: 03/2026 ou 12/03/2026"),
            ("Data de vencimento", "DATA_VENCIMENTO", "dd/mm/aaaa"),
            ("Observacoes", "OBS", "Informacoes adicionais"),
        ]

        for indice, (label, chave, placeholder) in enumerate(configuracao_campos):
            row = indice // 2
            col = indice % 2
            widget = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text=placeholder))
            self._campo(corpo, row, col, label, widget)
            campos[chave] = widget

        ctk.CTkLabel(
            corpo,
            text="O status e calculado automaticamente a partir da data de vencimento.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
        ).grid(row=8, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 8))

        metricas = ctk.CTkFrame(direita, fg_color="transparent")
        metricas.pack(fill="x", pady=(0, 12))
        metricas.grid_columnconfigure((0, 1, 2), weight=1)

        _, tabela_corpo = criar_secao(
            direita,
            "Registros",
            "Linhas em vermelho indicam certificado vencido, pendente para renovacao ou com data invalida.",
            expand=True,
        )

        topo_tabela = ctk.CTkFrame(tabela_corpo, fg_color="transparent")
        topo_tabela.pack(fill="x", pady=(0, 8))
        resumo_tabela = ctk.CTkLabel(
            topo_tabela,
            text="Nenhum filtro aplicado.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
        )
        resumo_tabela.pack(side="left")
        acoes_tabela = ctk.CTkFrame(topo_tabela, fg_color="transparent")
        acoes_tabela.pack(side="right")

        tabela_frame = ctk.CTkFrame(
            tabela_corpo,
            fg_color=BG_CARD,
            corner_radius=16,
            border_width=1,
            border_color=BORDER,
            height=430,
        )
        tabela_frame.pack(fill="both", expand=True)
        tabela_frame.pack_propagate(False)

        grade = ctk.CTkFrame(tabela_frame, fg_color="transparent")
        grade.pack(fill="both", expand=True, padx=8, pady=8)
        grade.grid_rowconfigure(0, weight=1)
        grade.grid_columnconfigure(0, weight=1)

        tree = ttk.Treeview(grade, columns=COLUNAS_CERTIFICADOS, show="headings", height=12)
        scroll_y = ctk.CTkScrollbar(grade, orientation="vertical", command=tree.yview, width=14)
        scroll_x = ctk.CTkScrollbar(grade, orientation="horizontal", command=tree.xview, height=14)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        scroll_x.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        tree.tag_configure("pendente", foreground="#f87171")
        tree.tag_configure("vencido", foreground=COR_ALERTA)
        tree.tag_configure("invalido", foreground="#fb7185")

        for coluna in COLUNAS_CERTIFICADOS:
            tree.heading(coluna, text=coluna.replace("_", " ").title())
            largura = 150 if coluna not in {"STATUS", "DIAS_RESTANTES"} else 130
            tree.column(coluna, anchor="center", width=largura)

        estado = {"selecionado": None, "somente_pendentes": False, "df_filtrado": pd.DataFrame()}

        def limpar_widget(widget):
            widget.delete(0, "end")

        def preencher_widget(widget, valor):
            widget.delete(0, "end")
            widget.insert(0, str(valor or ""))

        def carregar_df():
            df = self.banco.carregar_dataframe("CERTIFICADOS").fillna("").copy()
            atualizado = atualizar_status_certificados_df(df)
            if not atualizado.fillna("").astype(str).equals(df.fillna("").astype(str)):
                self.banco.escrever_aba("CERTIFICADOS", atualizado)
            return atualizado.fillna("").copy()

        def valor_campo(widget):
            return widget.get().strip()

        def coletar_dados():
            dados = {chave: valor_campo(campo) for chave, campo in campos.items()}
            if estado["selecionado"] is not None:
                df_atual = carregar_df()
                if estado["selecionado"] in df_atual.index:
                    dados["DATA_CADASTRO"] = str(df_atual.loc[estado["selecionado"], "DATA_CADASTRO"]).strip()
            return preparar_certificado(dados)

        def validar(dados):
            obrigatorios = {
                "TIPO_CERTIFICADO": "Informe o tipo de certificado.",
                "NOME": "Informe o nome do funcionario.",
                "CPF": "Informe o CPF.",
                "CARGA_HORARIA": "Informe a carga horaria.",
                "DATA_EMISSAO": "Informe o periodo de emissao.",
                "DATA_VENCIMENTO": "Informe a data de vencimento.",
            }
            for chave, mensagem in obrigatorios.items():
                if not str(dados.get(chave, "")).strip():
                    messagebox.showwarning("Aviso", mensagem)
                    campos[chave].focus()
                    return False
            if dados["STATUS"] == STATUS_DATA_INVALIDA:
                messagebox.showwarning("Aviso", "Informe a data de vencimento no formato dd/mm/aaaa.")
                campos["DATA_VENCIMENTO"].focus()
                return False
            return True

        def limpar_formulario():
            for campo in campos.values():
                limpar_widget(campo)
            estado["selecionado"] = None
            tree.selection_remove(*tree.selection())
            tree.focus("")

        def atualizar_metricas(df):
            for widget in metricas.winfo_children():
                widget.destroy()
            pendentes = df[df["STATUS"] == STATUS_PENDENTE]
            vencidos = df[df["STATUS"] == STATUS_VENCIDO]
            cards = [
                ("Total", str(len(df)), TEXT),
                ("Pendentes", str(len(pendentes)), "#f87171" if len(pendentes) else SUCCESS),
                ("Vencidos", str(len(vencidos)), COR_ALERTA if len(vencidos) else SUCCESS),
            ]
            for indice, (titulo, valor, cor) in enumerate(cards):
                criar_cartao_info(metricas, titulo, valor, cor).grid(
                    row=0, column=indice, sticky="nsew", padx=(0 if indice == 0 else 10, 0)
                )

        def atualizar_tabela(df=None):
            base = carregar_df() if df is None else atualizar_status_certificados_df(df.fillna("").copy())
            estado["df_filtrado"] = base.copy()
            tree.delete(*tree.get_children())
            for indice, row in base.iterrows():
                tags = ()
                if row.get("STATUS") == STATUS_PENDENTE:
                    tags = ("pendente",)
                elif row.get("STATUS") == STATUS_VENCIDO:
                    tags = ("vencido",)
                elif row.get("STATUS") == STATUS_DATA_INVALIDA:
                    tags = ("invalido",)
                tree.insert(
                    "",
                    "end",
                    iid=str(indice),
                    values=[str(row.get(coluna, "")) for coluna in COLUNAS_CERTIFICADOS],
                    tags=tags,
                )
            atualizar_metricas(base)
            if estado["somente_pendentes"]:
                resumo_tabela.configure(text=f"{len(base)} certificado(s) pendente(s) ou vencido(s).")
            else:
                resumo_tabela.configure(text=f"{len(base)} certificado(s) exibido(s).")

        def preencher_campos(_event=None):
            selecionado = tree.focus()
            if not selecionado:
                return

            df = estado["df_filtrado"].copy()
            if int(selecionado) not in df.index:
                return

            estado["selecionado"] = int(selecionado)
            linha = df.loc[int(selecionado)]
            for chave, campo in campos.items():
                preencher_widget(campo, linha.get(chave, ""))

        def salvar():
            dados = coletar_dados()
            if not validar(dados):
                return

            df = carregar_df()
            if estado["selecionado"] is None:
                df = pd.concat([df, pd.DataFrame([dados])], ignore_index=True)
            else:
                for chave, valor in dados.items():
                    df.loc[estado["selecionado"], chave] = valor

            df = atualizar_status_certificados_df(df)
            self.banco.escrever_aba("CERTIFICADOS", df)
            atualizar_tabela(df)
            limpar_formulario()
            messagebox.showinfo("Sucesso", "Certificado salvo com sucesso.")

        def excluir():
            if estado["selecionado"] is None:
                messagebox.showwarning("Aviso", "Selecione um certificado para excluir.")
                return
            if not messagebox.askyesno("Confirmacao", "Deseja excluir o certificado selecionado?"):
                return

            df = carregar_df().drop(index=estado["selecionado"]).reset_index(drop=True)
            self.banco.escrever_aba("CERTIFICADOS", df)
            atualizar_tabela(df)
            limpar_formulario()
            messagebox.showinfo("Sucesso", "Certificado excluido.")

        def mostrar_pendentes():
            estado["somente_pendentes"] = True
            df = carregar_df()
            df = df[df["STATUS"].isin([STATUS_PENDENTE, STATUS_VENCIDO, STATUS_DATA_INVALIDA])].copy()
            atualizar_tabela(df)
            limpar_formulario()

        def mostrar_todos():
            estado["somente_pendentes"] = False
            atualizar_tabela()
            limpar_formulario()

        def notificar_pendencias():
            pendentes = listar_certificados_pendentes(self.banco)
            if pendentes.empty:
                return
            destaques = []
            for _, linha in pendentes.head(5).iterrows():
                destaques.append(
                    f"{linha.get('NOME', '')} | {linha.get('TIPO_CERTIFICADO', '')} | {linha.get('DATA_VENCIMENTO', '')} | {linha.get('STATUS', '')}"
                )
            complemento = "\n..." if len(pendentes) > 5 else ""
            messagebox.showwarning(
                "Pendentes para renovacao",
                f"Existem {len(pendentes)} certificado(s) com vencimento proximo ou vencido.\n\n"
                f"{chr(10).join(destaques)}{complemento}",
            )

        def exportar_pdf():
            df_exportacao = estado["df_filtrado"].copy()
            if df_exportacao.empty:
                messagebox.showwarning("Aviso", "Nao ha certificados para exportar no filtro atual.")
                return

            nome_padrao = f"relatorio_certificados_{datetime.now():%d-%m-%Y}.pdf"
            caminho_pdf = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                initialfile=nome_padrao,
                filetypes=[("Arquivos PDF", "*.pdf")],
                title="Salvar relatorio de certificados como",
            )
            if not caminho_pdf:
                return

            try:
                titulo = "Relatorio de Certificados"
                if estado["somente_pendentes"]:
                    titulo = "Relatorio de Certificados Pendentes"
                gerar_relatorio_certificados(df_exportacao, caminho_pdf, titulo=titulo)
                os.startfile(caminho_pdf)
                messagebox.showinfo("Sucesso", f"Relatorio PDF gerado com sucesso.\n\nArquivo: {caminho_pdf}")
            except Exception as erro:
                messagebox.showerror("Erro", f"Erro ao gerar o relatorio PDF:\n{erro}")

        atualizar_tabela()
        tree.bind("<<TreeviewSelect>>", preencher_campos)

        botao_pendentes = ctk.CTkButton(acoes_tabela, text="Pendentes", command=mostrar_pendentes, width=92, height=32)
        estilizar_botao(botao_pendentes)
        botao_pendentes.pack(side="left", padx=(0, 8))

        botao_todos = ctk.CTkButton(acoes_tabela, text="Todos", command=mostrar_todos, width=70, height=32)
        estilizar_botao(botao_todos)
        botao_todos.pack(side="left", padx=(0, 8))

        botao_pdf = ctk.CTkButton(acoes_tabela, text="PDF", command=exportar_pdf, width=68, height=32)
        estilizar_botao(botao_pdf)
        botao_pdf.pack(side="left")

        self._acoes_horizontal(
            lateral,
            [
                ("Salvar certificado", salvar, True),
                ("Ver pendentes", mostrar_pendentes, False),
                ("Atualizar status", mostrar_todos, False),
                ("Relatorio PDF", exportar_pdf, False),
                ("Excluir", excluir, False),
                ("Limpar", limpar_formulario, False),
                ("Voltar", self.voltar_menu, False),
            ],
            largura_botao=170,
            colunas=2,
        )

        self.app.after(150, notificar_pendencias)

    def tela_carteirinhas_treinamento(self):
        pagina = self._nova_pagina(
            "Carteirinhas de treinamento",
            "Cadastre o colaborador e gere uma carteirinha profissional com a logo da empresa e espaco opcional para foto.",
        )

        conteudo = ctk.CTkFrame(pagina, fg_color="transparent")
        conteudo.pack(fill="both", expand=True)

        lateral = ctk.CTkFrame(conteudo, fg_color="transparent", width=460)
        lateral.pack(side="left", fill="y", padx=(0, 16))
        lateral.pack_propagate(False)

        direita = ctk.CTkFrame(conteudo, fg_color="transparent")
        direita.pack(side="right", fill="both", expand=True)

        secao_form, corpo_root = criar_secao(
            lateral,
            "Dados da carteirinha",
            "",
            expand=True,
        )
        ctk.CTkLabel(
            corpo_root,
            text="Preencha os dados da capacitacao. A foto nao e incluida no PDF e pode ser inserida pelo cliente depois.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            wraplength=380,
            justify="left",
        ).pack(anchor="w", padx=8, pady=(0, 12))
        ctk.CTkLabel(
            corpo_root,
            text="Os mesmos dados da carteirinha tambem alimentam o certificado. Para o certificado, voce so precisa complementar endereco e escolher se o CPF sera impresso.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            wraplength=380,
            justify="left",
        ).pack(anchor="w", padx=8, pady=(0, 12))

        corpo = ctk.CTkScrollableFrame(corpo_root, fg_color="transparent", corner_radius=0)
        corpo.pack(fill="both", expand=True)
        corpo.grid_columnconfigure((0, 1), weight=1)

        campos = {}
        variaveis = {}
        campos_historico = {
            "CODIGO",
            "EMPRESA",
            "FUNCAO",
            "TREINAMENTO",
            "CARGA_HORARIA",
            "INSTRUTOR",
            "DATA_EMISSAO",
            "VALIDADE",
            "RESPONSAVEL",
        }
        campos_historico_completo = {"RESPONSAVEL", "INSTRUTOR"}
        configuracao_campos = [
            ("Codigo", "CODIGO", "Ex.: TRN-20260324-001"),
            ("Nome", "NOME", "Nome completo"),
            ("CPF", "CPF", "000.000.000-00"),
            ("Empresa", "EMPRESA", "Empresa do colaborador"),
            ("Funcao", "FUNCAO", "Funcao / cargo"),
            ("Treinamento", "TREINAMENTO", "Nome do treinamento"),
            ("Carga horaria", "CARGA_HORARIA", "Ex.: 08 horas"),
            ("Instrutor", "INSTRUTOR", "Responsavel pelo curso"),
            ("Data emissao", "DATA_EMISSAO", "dd/mm/aaaa"),
            ("Validade", "VALIDADE", "dd/mm/aaaa"),
            ("Responsavel", "RESPONSAVEL", "Responsavel pela emissao"),
            ("Observacoes", "OBS", "Informacoes adicionais"),
        ]

        for indice, (label, chave, placeholder) in enumerate(configuracao_campos):
            row = indice // 2
            col = indice % 2
            if chave in campos_historico:
                widget = estilizar_combo(ctk.CTkComboBox(corpo, values=[]))
                widget.set("")
            else:
                widget = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text=placeholder))
            self._campo(corpo, row, col, label, widget)
            campos[chave] = widget

        ctk.CTkLabel(
            corpo,
            text="Campos recomendados: nome, empresa, funcao, treinamento, emissao e validade.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
        ).grid(row=12, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 8))

        secao_certificado = ctk.CTkFrame(
            corpo,
            fg_color=BG_PANEL,
            corner_radius=16,
            border_width=1,
            border_color=BORDER,
        )
        secao_certificado.grid(row=13, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 8))
        ctk.CTkLabel(
            secao_certificado,
            text="Dados adicionais do certificado",
            font=FONT_LABEL,
            text_color=TEXT,
        ).pack(anchor="w", padx=16, pady=(14, 4))
        ctk.CTkLabel(
            secao_certificado,
            text="Nome, CPF, treinamento, carga horaria, instrutor, emissao, validade e responsavel sao reaproveitados automaticamente dos campos acima.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            wraplength=360,
            justify="left",
        ).pack(anchor="w", padx=16, pady=(0, 10))

        corpo_certificado = ctk.CTkFrame(secao_certificado, fg_color="transparent")
        corpo_certificado.pack(fill="x", padx=10, pady=(0, 12))
        corpo_certificado.grid_columnconfigure((0, 1), weight=1)

        configuracao_certificado = [
            ("Rua", "CERTIFICADO_RUA", "Ex.: Av. Assis Chateaubriand"),
            ("Numero", "CERTIFICADO_NUMERO", "Ex.: 889"),
            ("Bairro", "CERTIFICADO_BAIRRO", "Ex.: Floresta"),
            ("CEP", "CERTIFICADO_CEP", "00000-000"),
        ]

        for indice, (label, chave, placeholder) in enumerate(configuracao_certificado):
            row = indice // 2
            col = indice % 2
            widget = estilizar_entry(ctk.CTkEntry(corpo_certificado, placeholder_text=placeholder))
            self._campo(corpo_certificado, row, col, label, widget)
            campos[chave] = widget

        linha_opcoes_certificado = ctk.CTkFrame(corpo_certificado, fg_color="transparent")
        linha_opcoes_certificado.grid(row=2, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 6))

        variaveis["CERTIFICADO_IMPRIMIR_CPF"] = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            linha_opcoes_certificado,
            text="Imprimir CPF no certificado",
            variable=variaveis["CERTIFICADO_IMPRIMIR_CPF"],
        ).pack(anchor="w", pady=(0, 6))

        variaveis["CERTIFICADO_ARTICULADA"] = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            linha_opcoes_certificado,
            text="Usar modelo Articulada",
            variable=variaveis["CERTIFICADO_ARTICULADA"],
        ).pack(anchor="w")

        estado = {"selecionado": None, "df_filtrado": pd.DataFrame()}

        def carregar_df():
            return self.banco.carregar_dataframe("TREINAMENTOS").fillna("").copy()

        def gerar_codigo():
            return f"TRN-{datetime.now():%Y%m%d-%H%M%S}"

        def valor_campo(widget):
            return widget.get().strip()

        def limpar_widget(widget):
            if isinstance(widget, ctk.CTkEntry):
                widget.delete(0, "end")
            else:
                widget.set("")

        def preencher_widget(widget, valor):
            texto = str(valor or "")
            if isinstance(widget, ctk.CTkEntry):
                widget.delete(0, "end")
                widget.insert(0, texto)
            else:
                widget.set(texto)

        def valor_variavel(chave):
            if chave == "CERTIFICADO_ARTICULADA":
                return "ARTICULADA" if variaveis[chave].get() else "TESOURA"
            return "SIM" if variaveis[chave].get() else "NAO"

        def preencher_variavel(chave, valor):
            if chave == "CERTIFICADO_ARTICULADA":
                variaveis[chave].set(str(valor or "").strip().upper() == "ARTICULADA")
                return
            variaveis[chave].set(str(valor or "").strip().upper() not in {"", "0", "NAO", "N", "FALSE"})

        def atualizar_historico_campos(df=None):
            base = carregar_df() if df is None else df.fillna("").copy()
            for chave in campos_historico:
                if chave not in base.columns:
                    continue
                valores = []
                vistos = set()
                for valor in reversed(base[chave].astype(str).tolist()):
                    texto = valor.strip()
                    if not texto or texto in vistos:
                        continue
                    valores.append(texto)
                    vistos.add(texto)
                try:
                    if chave in campos_historico_completo:
                        campos[chave].configure(values=valores[:30])
                    else:
                        campos[chave].configure(values=valores[:1])
                except Exception:
                    pass

        def coletar_dados():
            dados = {chave: valor_campo(campo) for chave, campo in campos.items()}
            for chave in variaveis:
                valor = valor_variavel(chave)
                if chave == "CERTIFICADO_ARTICULADA":
                    dados["CERTIFICADO_MODELO"] = valor
                else:
                    dados[chave] = valor
            if not dados["CODIGO"]:
                dados["CODIGO"] = gerar_codigo()
            dados["DATA_CADASTRO"] = datetime.now().strftime("%d/%m/%Y %H:%M")
            dados["DATA_ATUALIZACAO"] = datetime.now().strftime("%d/%m/%Y %H:%M")
            return dados

        def validar(dados):
            obrigatorios = {
                "NOME": "Informe o nome do colaborador.",
                "EMPRESA": "Informe a empresa.",
                "FUNCAO": "Informe a funcao.",
                "TREINAMENTO": "Informe o treinamento.",
                "DATA_EMISSAO": "Informe a data de emissao.",
                "VALIDADE": "Informe a validade.",
                "RESPONSAVEL": "Informe o responsavel.",
            }
            for chave, mensagem in obrigatorios.items():
                if not dados.get(chave, "").strip():
                    messagebox.showwarning("Aviso", mensagem)
                    campos[chave].focus()
                    return False
            return True

        def validar_certificado(dados):
            obrigatorios = {
                "NOME": "Informe o nome do colaborador.",
                "TREINAMENTO": "Informe o treinamento.",
                "CARGA_HORARIA": "Informe a carga horaria.",
                "DATA_EMISSAO": "Informe a data de emissao.",
                "INSTRUTOR": "Informe o instrutor.",
                "RESPONSAVEL": "Informe o responsavel.",
                "CERTIFICADO_RUA": "Informe a rua para o certificado.",
                "CERTIFICADO_NUMERO": "Informe o numero para o certificado.",
                "CERTIFICADO_BAIRRO": "Informe o bairro para o certificado.",
                "CERTIFICADO_CEP": "Informe o CEP para o certificado.",
            }
            for chave, mensagem in obrigatorios.items():
                if not str(dados.get(chave, "")).strip():
                    messagebox.showwarning("Aviso", mensagem)
                    if chave in campos:
                        campos[chave].focus()
                    return False
            return True

        _, tabela_corpo = criar_secao(
            direita,
            "Registros",
            "Selecione um colaborador para editar, excluir ou gerar a carteirinha em PDF.",
            expand=True,
        )
        topo_tabela = ctk.CTkFrame(tabela_corpo, fg_color="transparent")
        topo_tabela.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(
            topo_tabela,
            text="Use os atalhos para filtrar registros e exportar uma ou varias carteirinhas.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
        ).pack(side="left")
        acoes_tabela = ctk.CTkFrame(topo_tabela, fg_color="transparent")
        acoes_tabela.pack(side="right")
        resumo_tabela = ctk.CTkLabel(
            tabela_corpo,
            text="Nenhum filtro aplicado.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
        )
        resumo_tabela.pack(anchor="w", pady=(0, 8))

        tabela_frame = ctk.CTkFrame(
            tabela_corpo,
            fg_color=BG_CARD,
            corner_radius=16,
            border_width=1,
            border_color=BORDER,
            height=420,
        )
        tabela_frame.pack(fill="both", expand=True)
        tabela_frame.pack_propagate(False)

        grade = ctk.CTkFrame(tabela_frame, fg_color="transparent")
        grade.pack(fill="both", expand=True, padx=8, pady=8)
        grade.grid_rowconfigure(0, weight=1)
        grade.grid_columnconfigure(0, weight=1)

        tree = ttk.Treeview(grade, columns=COLUNAS_CARTEIRINHA, show="headings", height=12)
        scroll_y = ctk.CTkScrollbar(grade, orientation="vertical", command=tree.yview, width=14)
        scroll_x = ctk.CTkScrollbar(grade, orientation="horizontal", command=tree.xview, height=14)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        scroll_x.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        for coluna in COLUNAS_CARTEIRINHA:
            tree.heading(coluna, text=coluna.replace("_", " ").title())
            tree.column(coluna, anchor="center", width=145)

        def limpar_formulario():
            for campo in campos.values():
                limpar_widget(campo)
            variaveis["CERTIFICADO_IMPRIMIR_CPF"].set(True)
            variaveis["CERTIFICADO_ARTICULADA"].set(False)
            estado["selecionado"] = None
            tree.selection_remove(*tree.selection())
            tree.focus("")

        def atualizar_tabela(df=None):
            df = carregar_df() if df is None else df.fillna("").copy()
            estado["df_filtrado"] = df.copy()
            atualizar_historico_campos(df)
            tree.delete(*tree.get_children())
            for indice, row in df.iterrows():
                tree.insert("", "end", iid=str(indice), values=[str(row.get(coluna, "")) for coluna in COLUNAS_CARTEIRINHA])
            total_base = len(carregar_df())
            if len(df) == total_base:
                resumo_tabela.configure(text=f"{len(df)} registro(s) exibido(s). Nenhum filtro aplicado.")
            else:
                resumo_tabela.configure(text=f"{len(df)} registro(s) exibido(s) apos filtro.")
            return df

        def preencher_campos(_event=None):
            selecionado = tree.focus()
            if not selecionado:
                return

            df = estado["df_filtrado"].copy()
            if int(selecionado) not in df.index:
                return

            estado["selecionado"] = int(selecionado)
            linha = df.loc[int(selecionado)]
            for chave, campo in campos.items():
                preencher_widget(campo, linha.get(chave, ""))
            preencher_variavel("CERTIFICADO_IMPRIMIR_CPF", linha.get("CERTIFICADO_IMPRIMIR_CPF", "SIM"))
            preencher_variavel("CERTIFICADO_ARTICULADA", linha.get("CERTIFICADO_MODELO", "TESOURA"))

        def salvar():
            dados = coletar_dados()
            if not validar(dados):
                return

            df = carregar_df()
            if estado["selecionado"] is None:
                df = pd.concat([df, pd.DataFrame([dados])], ignore_index=True)
            else:
                indice = estado["selecionado"]
                for chave, valor in dados.items():
                    df.loc[indice, chave] = valor

            self.banco.escrever_aba("TREINAMENTOS", df)
            atualizar_tabela()
            limpar_formulario()
            messagebox.showinfo("Sucesso", "Registro da carteirinha salvo com sucesso.")

        def _selecionar_caminho_pdf(dados):
            nome_padrao = f"carteirinha_{dados['NOME'].replace(' ', '_')}_{datetime.now():%d-%m-%Y}.pdf"
            return filedialog.asksaveasfilename(
                defaultextension=".pdf",
                initialfile=nome_padrao,
                filetypes=[("Arquivos PDF", "*.pdf")],
                title="Salvar carteirinha como",
            )

        def _selecionar_caminho_certificado_word(dados):
            nome_padrao = f"certificado_{dados['NOME'].replace(' ', '_')}_{datetime.now():%d-%m-%Y}.docx"
            return filedialog.asksaveasfilename(
                defaultextension=".docx",
                initialfile=nome_padrao,
                filetypes=[("Documento do Word", "*.docx")],
                title="Salvar certificado Word como",
            )

        def _salvar_df(df):
            self.banco.escrever_aba("TREINAMENTOS", df)
            atualizar_tabela(df)

        def _gerar_pdf_com_tratamento(dados, caminho_pdf):
            try:
                gerar_carteirinha_treinamento(dados, caminho_pdf)
                return caminho_pdf
            except PermissionError:
                messagebox.showwarning(
                    "Arquivo em uso",
                    "Nao foi possivel sobrescrever o PDF porque ele esta aberto ou sem permissao de escrita.\n\n"
                    "Feche o arquivo atual e escolha outro local ou nome para continuar.",
                )
                novo_caminho = _selecionar_caminho_pdf(dados)
                if not novo_caminho:
                    return None
                gerar_carteirinha_treinamento(dados, novo_caminho)
                return novo_caminho

        def _gerar_certificado_com_tratamento(dados, caminho_arquivo):
            try:
                gerar_certificado_word(dados, caminho_arquivo)
                return caminho_arquivo
            except PermissionError:
                messagebox.showwarning(
                    "Arquivo em uso",
                    "Nao foi possivel sobrescrever o arquivo Word do certificado porque ele esta aberto ou sem permissao de escrita.\n\n"
                    "Feche o arquivo atual e escolha outro local ou nome para continuar.",
                )
                novo_caminho = _selecionar_caminho_certificado_word(dados)
                if not novo_caminho:
                    return None
                gerar_certificado_word(dados, novo_caminho)
                return novo_caminho

        def _gerar_e_registrar_pdf(dados, caminho_pdf, indice=None, abrir_arquivo=True):
            caminho_final = _gerar_pdf_com_tratamento(dados, caminho_pdf)
            if not caminho_final:
                return None
            df = carregar_df()

            if indice is None:
                filtro = df["CODIGO"].astype(str) == str(dados.get("CODIGO", ""))
                indices = df.index[filtro].tolist()
                indice = indices[-1] if indices else None

            if indice is not None and indice in df.index:
                df.loc[indice, "PDF_CAMINHO"] = caminho_final
                df.loc[indice, "MODELO_VERSAO"] = MODELO_CARTEIRINHA_VERSAO
                df.loc[indice, "DATA_ATUALIZACAO"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                _salvar_df(df)

            if abrir_arquivo:
                os.startfile(caminho_final)
            return caminho_final

        def _gerar_e_registrar_certificado_word(dados, caminho_arquivo, indice=None, abrir_arquivo=True):
            caminho_final = _gerar_certificado_com_tratamento(dados, caminho_arquivo)
            if not caminho_final:
                return None
            df = carregar_df()

            if indice is None:
                filtro = df["CODIGO"].astype(str) == str(dados.get("CODIGO", ""))
                indices = df.index[filtro].tolist()
                indice = indices[-1] if indices else None

            if indice is not None and indice in df.index:
                df.loc[indice, "CERTIFICADO_WORD_CAMINHO"] = caminho_final
                df.loc[indice, "DATA_ATUALIZACAO"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                _salvar_df(df)

            if abrir_arquivo:
                os.startfile(caminho_final)
            return caminho_final

        def atualizar_carteirinha():
            if estado["selecionado"] is None:
                messagebox.showwarning("Aviso", "Selecione uma carteirinha para atualizar.")
                return

            dados = coletar_dados()
            if not validar(dados):
                return

            indice = estado["selecionado"]
            df = carregar_df()
            if indice not in df.index:
                messagebox.showwarning("Aviso", "O registro selecionado nao esta mais disponivel.")
                return

            for chave, valor in dados.items():
                df.loc[indice, chave] = valor

            caminho_pdf = str(df.loc[indice, "PDF_CAMINHO"]).strip() if "PDF_CAMINHO" in df.columns else ""
            if not caminho_pdf:
                caminho_pdf = _selecionar_caminho_pdf(dados)
                if not caminho_pdf:
                    return

            try:
                df.loc[indice, "PDF_CAMINHO"] = caminho_pdf
                df.loc[indice, "MODELO_VERSAO"] = MODELO_CARTEIRINHA_VERSAO
                df.loc[indice, "DATA_ATUALIZACAO"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                _salvar_df(df)
                caminho_final = _gerar_pdf_com_tratamento(dados, caminho_pdf)
                if not caminho_final:
                    return
                if caminho_final != caminho_pdf:
                    df = carregar_df()
                    if indice in df.index:
                        df.loc[indice, "PDF_CAMINHO"] = caminho_final
                        df.loc[indice, "MODELO_VERSAO"] = MODELO_CARTEIRINHA_VERSAO
                        df.loc[indice, "DATA_ATUALIZACAO"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                        _salvar_df(df)
                os.startfile(caminho_final)
                messagebox.showinfo(
                    "Sucesso",
                    "Carteirinha atualizada com sucesso.\n\nO PDF foi refeito com o modelo atual.",
                )
            except Exception as erro:
                messagebox.showerror("Erro", f"Erro ao atualizar a carteirinha:\n{erro}")

        def excluir():
            if estado["selecionado"] is None:
                messagebox.showwarning("Aviso", "Selecione um registro para excluir.")
                return
            if not messagebox.askyesno("Confirmacao", "Deseja excluir a carteirinha selecionada?"):
                return

            df = carregar_df().drop(index=estado["selecionado"]).reset_index(drop=True)
            self.banco.escrever_aba("TREINAMENTOS", df)
            atualizar_tabela()
            limpar_formulario()
            messagebox.showinfo("Sucesso", "Carteirinha excluida.")

        def adicionar_nova():
            limpar_formulario()
            campos["CODIGO"].insert(0, gerar_codigo())
            campos["NOME"].focus()

        def gerar_pdf():
            dados = coletar_dados()
            indice = estado["selecionado"]

            if not validar(dados):
                return

            caminho = _selecionar_caminho_pdf(dados)
            if not caminho:
                return

            try:
                caminho_final = _gerar_e_registrar_pdf(dados, caminho, indice=indice, abrir_arquivo=True)
                if not caminho_final:
                    return
                messagebox.showinfo("Sucesso", f"Carteirinha gerada com sucesso.\n\nArquivo: {caminho_final}")
            except Exception as erro:
                messagebox.showerror("Erro", f"Erro ao gerar a carteirinha:\n{erro}")

        def gerar_word_certificado():
            dados = coletar_dados()
            indice = estado["selecionado"]

            if not validar_certificado(dados):
                return

            caminho = _selecionar_caminho_certificado_word(dados)
            if not caminho:
                return

            try:
                caminho_final = _gerar_e_registrar_certificado_word(dados, caminho, indice=indice, abrir_arquivo=True)
                if not caminho_final:
                    return
                messagebox.showinfo("Sucesso", f"Certificado Word gerado com sucesso.\n\nArquivo: {caminho_final}")
            except Exception as erro:
                messagebox.showerror("Erro", f"Erro ao gerar o certificado:\n{erro}")

        def imprimir_certificado():
            dados = coletar_dados()
            indice = estado["selecionado"]

            if not validar_certificado(dados):
                return

            caminho_word = ""
            if indice is not None:
                df = carregar_df()
                if indice in df.index and "CERTIFICADO_WORD_CAMINHO" in df.columns:
                    caminho_word = str(df.loc[indice, "CERTIFICADO_WORD_CAMINHO"]).strip()

            if not caminho_word:
                caminho_word = _selecionar_caminho_certificado_word(dados)
                if not caminho_word:
                    return

            try:
                caminho_final = _gerar_e_registrar_certificado_word(dados, caminho_word, indice=indice, abrir_arquivo=False)
                if not caminho_final:
                    return
                os.startfile(caminho_final, "print")
                messagebox.showinfo("Sucesso", f"Certificado Word enviado para impressao.\n\nArquivo: {caminho_final}")
            except Exception as erro:
                messagebox.showerror("Erro", f"Erro ao imprimir o certificado:\n{erro}")

        def _nome_pdf_carteirinha(dados):
            nome = str(dados.get("NOME", "")).strip().replace(" ", "_")
            codigo = str(dados.get("CODIGO", "")).strip().replace(" ", "_")
            base = f"carteirinha_{codigo}_{nome}".strip("_")
            return f"{base or 'carteirinha'}.pdf"

        def exportar_filtro():
            df_filtrado = estado["df_filtrado"].copy()
            if df_filtrado.empty:
                messagebox.showwarning("Aviso", "Nao ha carteirinhas no filtro atual para exportar.")
                return

            pasta_destino = filedialog.askdirectory(title="Escolha a pasta para salvar as carteirinhas filtradas")
            if not pasta_destino:
                return

            erros = []
            quantidade = 0
            for _, linha in df_filtrado.iterrows():
                dados = linha.to_dict()
                if not validar(dados):
                    return
                caminho_pdf = os.path.join(pasta_destino, _nome_pdf_carteirinha(dados))
                try:
                    if _gerar_e_registrar_pdf(dados, caminho_pdf, indice=linha.name, abrir_arquivo=False) is None:
                        return
                    quantidade += 1
                except Exception as erro:
                    erros.append(f"{dados.get('NOME', 'Sem nome')}: {erro}")

            if erros:
                detalhes = "\n".join(erros[:5])
                extra = "\n..." if len(erros) > 5 else ""
                messagebox.showwarning(
                    "Exportacao parcial",
                    f"{quantidade} carteirinha(s) gerada(s) em:\n{pasta_destino}\n\nFalhas:\n{detalhes}{extra}",
                )
                return

            os.startfile(pasta_destino)
            messagebox.showinfo(
                "Sucesso",
                f"{quantidade} carteirinha(s) gerada(s) com sucesso na pasta:\n{pasta_destino}",
            )

        def abrir_filtro():
            janela = ctk.CTkToplevel(self.app, fg_color=BG_APP)
            janela.title("Filtro de carteirinhas")
            janela.geometry("560x420")
            janela.transient(self.app)
            janela.grab_set()

            pagina_filtro = criar_pagina(janela)
            criar_cabecalho(pagina_filtro, "Filtro de carteirinhas", "Filtre por treinamento, empresa, codigo ou combine criterios.")

            usar_treinamento = ctk.BooleanVar()
            usar_empresa = ctk.BooleanVar()
            usar_codigo = ctk.BooleanVar()

            area_scroll = ctk.CTkScrollableFrame(pagina_filtro, fg_color="transparent", corner_radius=0)
            area_scroll.pack(fill="both", expand=True)

            _, corpo_filtro = criar_secao(area_scroll, "Criterios")
            corpo_filtro.grid_columnconfigure((0, 1), weight=1)

            frame_treinamento = ctk.CTkFrame(corpo_filtro, fg_color="transparent")
            frame_treinamento.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
            ctk.CTkCheckBox(frame_treinamento, text="Tipo de treinamento", variable=usar_treinamento).pack(anchor="w")
            treinamento = estilizar_entry(ctk.CTkEntry(frame_treinamento, placeholder_text="Digite o treinamento"))
            treinamento.pack(fill="x", pady=(8, 0))

            frame_empresa = ctk.CTkFrame(corpo_filtro, fg_color="transparent")
            frame_empresa.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
            ctk.CTkCheckBox(frame_empresa, text="Empresa", variable=usar_empresa).pack(anchor="w")
            empresa = estilizar_entry(ctk.CTkEntry(frame_empresa, placeholder_text="Digite a empresa"))
            empresa.pack(fill="x", pady=(8, 0))

            frame_codigo = ctk.CTkFrame(corpo_filtro, fg_color="transparent")
            frame_codigo.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
            ctk.CTkCheckBox(frame_codigo, text="Codigo", variable=usar_codigo).pack(anchor="w")
            codigo = estilizar_entry(ctk.CTkEntry(frame_codigo, placeholder_text="Digite o codigo"))
            codigo.pack(fill="x", pady=(8, 0))

            def aplicar_filtro():
                df_filtrado = carregar_df()

                if usar_treinamento.get():
                    valor = treinamento.get().strip()
                    if valor:
                        df_filtrado = df_filtrado[
                            df_filtrado["TREINAMENTO"].astype(str).str.contains(valor, case=False, na=False)
                        ]

                if usar_empresa.get():
                    valor = empresa.get().strip()
                    if valor:
                        df_filtrado = df_filtrado[
                            df_filtrado["EMPRESA"].astype(str).str.contains(valor, case=False, na=False)
                        ]

                if usar_codigo.get():
                    valor = codigo.get().strip()
                    if valor:
                        df_filtrado = df_filtrado[
                            df_filtrado["CODIGO"].astype(str).str.contains(valor, case=False, na=False)
                        ]

                atualizar_tabela(df_filtrado)
                limpar_formulario()
                janela.destroy()

            self._acoes_horizontal(
                pagina_filtro,
                [
                    ("Aplicar filtros", aplicar_filtro, True),
                    ("Limpar", lambda: (atualizar_tabela(), limpar_formulario(), janela.destroy()), False),
                ],
            )

        atualizar_tabela()
        tree.bind("<<TreeviewSelect>>", preencher_campos)
        tree.bind("<Double-1>", lambda _event: gerar_pdf())

        botao_filtrar = ctk.CTkButton(acoes_tabela, text="⚲", command=abrir_filtro, width=36, height=32)
        estilizar_botao(botao_filtrar)
        botao_filtrar.pack(side="left", padx=(0, 8))
        adicionar_tooltip(botao_filtrar, "Filtrar")

        botao_novo = ctk.CTkButton(acoes_tabela, text="+", command=adicionar_nova, width=36, height=32)
        estilizar_botao(botao_novo)
        botao_novo.pack(side="left", padx=(0, 8))
        adicionar_tooltip(botao_novo, "Adicionar carteira")

        botao_pdf = ctk.CTkButton(acoes_tabela, text="PDF", command=gerar_pdf, width=54, height=32)
        estilizar_botao(botao_pdf)
        botao_pdf.pack(side="left", padx=(0, 8))
        adicionar_tooltip(botao_pdf, "Salvar PDF")

        botao_certificado = ctk.CTkButton(acoes_tabela, text="Word", command=gerar_word_certificado, width=58, height=32)
        estilizar_botao(botao_certificado)
        botao_certificado.pack(side="left", padx=(0, 8))
        adicionar_tooltip(botao_certificado, "Gerar certificado Word")

        botao_lote = ctk.CTkButton(acoes_tabela, text="Lote", command=exportar_filtro, width=60, height=32)
        estilizar_botao(botao_lote)
        botao_lote.pack(side="left")
        adicionar_tooltip(botao_lote, "Exportar filtro")

        self._acoes_horizontal(
            lateral,
            [
                ("Adicionar carteira", adicionar_nova, False),
                ("Salvar registro", salvar, True),
                ("Atualizar carteira", atualizar_carteirinha, False),
                ("Salvar PDF", gerar_pdf, False),
                ("Gerar certificado Word", gerar_word_certificado, False),
                ("Imprimir certificado", imprimir_certificado, False),
                ("Exportar filtro", exportar_filtro, False),
                ("Excluir", excluir, False),
                ("Limpar", limpar_formulario, False),
                ("Voltar", self.voltar_menu, False),
            ],
            largura_botao=160,
            colunas=2,
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
