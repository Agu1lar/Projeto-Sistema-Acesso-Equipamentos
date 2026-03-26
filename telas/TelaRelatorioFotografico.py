from __future__ import annotations

from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from services.service_images import salvar_imagens
from services.service_pdf import gerar_relatorio
from services.service_relatorios_fotograficos import (
    abrir_relatorio,
    caminho_pdf_modelo,
    exportar_relatorios_para_pasta,
    listar_relatorios,
    obter_pasta_base,
    salvar_pasta_base,
)
from utils.ui import (
    ACCENT,
    BG_CARD,
    BORDER,
    FONT_HEADING,
    FONT_SMALL,
    TEXT,
    TEXT_MUTED,
    criar_cabecalho,
    criar_cartao_info,
    criar_label_form,
    criar_pagina,
    criar_secao,
    estilizar_botao,
    estilizar_entry,
)


class TelaRelatorioFotografico:
    def __init__(self, app, frame, banco, voltar_menu):
        self.app = app
        self.frame = frame
        self.banco = banco
        self.voltar_menu = voltar_menu
        self.imagens_selecionadas = []
        self.campos = {}
        self.salvar_automatico = ctk.BooleanVar(value=True)
        self.relatorios_encontrados = []

    def abrir(self):
        self.limpar()
        pagina = criar_pagina(self.frame)
        criar_cabecalho(
            pagina,
            "Relatorio fotografico",
            "Monte o documento visual, salve por modelo e consulte o historico sem sair da aplicacao.",
        )
        conteudo = ctk.CTkScrollableFrame(pagina, fg_color="transparent", corner_radius=0)
        conteudo.pack(fill="both", expand=True)

        metricas = ctk.CTkFrame(conteudo, fg_color="transparent")
        metricas.pack(fill="x", pady=(0, 16))
        metricas.grid_columnconfigure((0, 1, 2), weight=1)

        infos = [
            ("Imagens selecionadas", str(len(self.imagens_selecionadas)), TEXT),
            ("Pasta base", "Configurada", ACCENT),
            ("Relatorios salvos", str(len(listar_relatorios())), TEXT),
        ]
        for indice, (titulo, valor, cor) in enumerate(infos):
            criar_cartao_info(metricas, titulo, valor, cor).grid(
                row=0, column=indice, sticky="nsew", padx=(0 if indice == 0 else 10, 0)
            )

        _, corpo = criar_secao(conteudo, "Dados do relatorio", "Preencha os campos principais e selecione as imagens.")
        corpo.grid_columnconfigure((0, 1), weight=1)

        for indice, (placeholder, chave) in enumerate(
            [
                ("Cliente", "cliente"),
                ("Obra", "obra"),
                ("Contrato", "contrato"),
                ("Equipamento", "equipamento"),
                ("Modelo", "modelo"),
                ("Patrimonio", "patrimonio"),
                ("Numero de serie", "serie"),
            ]
        ):
            row = indice // 2
            col = indice % 2
            criar_label_form(corpo, placeholder).grid(row=row * 2, column=col, sticky="w", padx=8, pady=(0, 6))
            campo = estilizar_entry(ctk.CTkEntry(corpo, placeholder_text=placeholder))
            campo.grid(row=row * 2 + 1, column=col, sticky="ew", padx=8, pady=(0, 12))
            self.campos[chave] = campo

        utilitarios = ctk.CTkFrame(corpo, fg_color=BG_CARD, corner_radius=16, border_width=1, border_color=BORDER)
        utilitarios.grid(row=8, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 12))
        ctk.CTkLabel(utilitarios, text=f"Pasta base atual: {obter_pasta_base()}", text_color=TEXT_MUTED, wraplength=820).pack(
            anchor="w", padx=14, pady=(12, 6)
        )
        ctk.CTkCheckBox(
            utilitarios,
            text="Salvar automaticamente na pasta do modelo",
            variable=self.salvar_automatico,
        ).pack(anchor="w", padx=14, pady=(0, 12))

        botoes_util = ctk.CTkFrame(utilitarios, fg_color="transparent")
        botoes_util.pack(anchor="w", padx=14, pady=(0, 12))
        for indice, (texto, comando, primario) in enumerate(
            [
                ("Definir pasta base", self.definir_pasta_base, False),
                ("Pesquisar relatorios", self.abrir_pesquisa, False),
            ]
        ):
            botao = ctk.CTkButton(botoes_util, text=texto, command=comando, width=180)
            estilizar_botao(botao, primario=primario)
            botao.grid(row=0, column=indice, padx=(0, 10))

        _, imagens_corpo = criar_secao(conteudo, "Imagens", "As imagens selecionadas serao copiadas para a pasta do contrato.")
        self.label_imagens = ctk.CTkLabel(
            imagens_corpo,
            text="Nenhuma imagem selecionada",
            text_color=TEXT_MUTED,
            font=FONT_SMALL,
        )
        self.label_imagens.pack(anchor="w", pady=(0, 12))

        barra = ctk.CTkFrame(imagens_corpo, fg_color="transparent")
        barra.pack(fill="x")
        for indice, (texto, comando, primario) in enumerate(
            [
                ("Selecionar imagens", self.selecionar_imagens, False),
                ("Gerar relatorio", self.gerar, True),
                ("Voltar", self.voltar_menu, False),
            ]
        ):
            botao = ctk.CTkButton(barra, text=texto, command=comando, width=170)
            estilizar_botao(botao, primario=primario)
            botao.grid(row=0, column=indice, padx=(0, 10), pady=4)

    def abrir_pesquisa(self):
        self.limpar()
        pagina = criar_pagina(self.frame)
        criar_cabecalho(
            pagina,
            "Pesquisa de relatorios fotograficos",
            "Busque por modelo, por data ou apenas por data para listar tudo daquele dia.",
        )

        _, filtros = criar_secao(pagina, "Filtros")
        filtros.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        criar_label_form(filtros, "Modelo").grid(row=0, column=0, sticky="w", padx=8, pady=(0, 6))
        self.filtro_modelo = estilizar_entry(ctk.CTkEntry(filtros, placeholder_text="Modelo"))
        self.filtro_modelo.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 12))

        criar_label_form(filtros, "Data").grid(row=0, column=1, sticky="w", padx=8, pady=(0, 6))
        self.filtro_data = estilizar_entry(ctk.CTkEntry(filtros, placeholder_text="dd/mm/aaaa"))
        self.filtro_data.grid(row=1, column=1, sticky="ew", padx=8, pady=(0, 12))

        for indice, (texto, comando, primario) in enumerate(
            [
                ("Buscar", self.buscar_relatorios, True),
                ("Limpar", self.limpar_filtros_relatorios, False),
                ("Nova geracao", self.abrir, False),
                ("Voltar", self.voltar_menu, False),
            ],
            start=2,
        ):
            botao = ctk.CTkButton(filtros, text=texto, command=comando, width=140)
            estilizar_botao(botao, primario=primario)
            botao.grid(row=1, column=indice, padx=8, pady=(0, 12), sticky="ew")

        self.label_resultado = ctk.CTkLabel(
            pagina,
            text="Nenhuma busca executada.",
            text_color=TEXT_MUTED,
            font=FONT_SMALL,
        )
        self.label_resultado.pack(anchor="w", pady=(0, 10))

        _, tabela_corpo = criar_secao(pagina, "Resultados", expand=True)
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

        self.tree_relatorios = ttk.Treeview(tabela_frame, columns=("modelo", "data", "arquivo"), show="headings", height=12)
        for coluna, titulo, largura, anchor in [
            ("modelo", "Modelo", 180, "center"),
            ("data", "Data", 160, "center"),
            ("arquivo", "Arquivo", 420, "w"),
        ]:
            self.tree_relatorios.heading(coluna, text=titulo)
            self.tree_relatorios.column(coluna, width=largura, anchor=anchor)

        self.tree_relatorios.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        self.tree_relatorios.bind("<Double-1>", self.abrir_relatorio_selecionado)

        scroll_y = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tree_relatorios.yview)
        scroll_y.pack(side="right", fill="y", pady=8)
        self.tree_relatorios.configure(yscrollcommand=scroll_y.set)

        rodape = ctk.CTkFrame(pagina, fg_color="transparent")
        rodape.pack(fill="x", pady=(10, 0))
        for indice, (texto, comando, primario) in enumerate(
            [
                ("Abrir selecionado", self.abrir_relatorio_selecionado, True),
                ("Exportar lote", self.exportar_lote_relatorios, False),
            ]
        ):
            botao = ctk.CTkButton(rodape, text=texto, command=comando, width=180)
            estilizar_botao(botao, primario=primario)
            botao.pack(side="left", padx=(0, 10))

        self.buscar_relatorios()

    def limpar(self):
        for widget in self.frame.winfo_children():
            widget.destroy()

    def selecionar_imagens(self):
        caminhos = filedialog.askopenfilenames(
            title="Selecionar imagens",
            filetypes=[("Imagens", "*.png *.jpg *.jpeg")],
        )
        if not caminhos:
            return

        self.imagens_selecionadas = list(caminhos)
        datas = sorted(datetime.fromtimestamp(Path(caminho).stat().st_mtime) for caminho in caminhos if Path(caminho).exists())
        complemento = ""
        if datas:
            complemento = f" | periodo: {datas[0]:%d/%m/%Y %H:%M} ate {datas[-1]:%d/%m/%Y %H:%M}"
        self.label_imagens.configure(text=f"{len(caminhos)} imagem(ns) selecionada(s){complemento}")

    def definir_pasta_base(self):
        pasta = filedialog.askdirectory(title="Escolha a pasta base dos relatorios por modelo")
        if not pasta:
            return

        salvar_pasta_base(pasta)
        messagebox.showinfo("Sucesso", "Pasta base salva com sucesso.")
        self.abrir()

    def gerar(self):
        contrato = self.campos["contrato"].get().strip()
        modelo = self.campos["modelo"].get().strip()

        if not contrato:
            messagebox.showwarning("Aviso", "Informe o contrato")
            return
        if not modelo:
            messagebox.showwarning("Aviso", "Informe o modelo da maquina")
            return
        if not self.imagens_selecionadas:
            messagebox.showwarning("Aviso", "Selecione imagens")
            return

        dados = {
            "Cliente": self.campos["cliente"].get().strip(),
            "Obra": self.campos["obra"].get().strip(),
            "Contrato": contrato,
            "Equipamento": self.campos["equipamento"].get().strip(),
            "Modelo": modelo,
            "Serie": self.campos["serie"].get().strip(),
            "Patrimonio": self.campos["patrimonio"].get().strip(),
            "Data": datetime.now().strftime("%d/%m/%Y"),
        }

        try:
            imagens_salvas = salvar_imagens(self.imagens_selecionadas, contrato)
            if self.salvar_automatico.get():
                caminho_pdf = caminho_pdf_modelo(modelo, contrato, datetime.now())
            else:
                nome_arquivo = f"Relatorio_{modelo}_{contrato}_{datetime.now():%d-%m-%Y}.pdf"
                caminho_pdf = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    initialfile=nome_arquivo,
                    filetypes=[("Arquivo PDF", "*.pdf")],
                    title="Salvar relatorio como",
                )
                if not caminho_pdf:
                    return

            gerar_relatorio(dados, imagens_salvas, caminho_pdf)
            messagebox.showinfo("Sucesso", f"Relatorio gerado:\n{caminho_pdf}")
        except Exception as erro:
            messagebox.showerror("Erro", str(erro))

    def buscar_relatorios(self):
        try:
            modelo = self.filtro_modelo.get().strip() if hasattr(self, "filtro_modelo") else ""
            data = self.filtro_data.get().strip() if hasattr(self, "filtro_data") else ""
            self.relatorios_encontrados = listar_relatorios(modelo=modelo, data=data)
        except ValueError as erro:
            messagebox.showwarning("Filtro invalido", str(erro))
            return
        except Exception as erro:
            messagebox.showerror("Erro", str(erro))
            return

        if hasattr(self, "tree_relatorios"):
            self.tree_relatorios.delete(*self.tree_relatorios.get_children())
            for indice, relatorio in enumerate(self.relatorios_encontrados):
                self.tree_relatorios.insert(
                    "",
                    "end",
                    iid=str(indice),
                    values=(relatorio["modelo"], relatorio["data"], relatorio["arquivo"]),
                )

        if hasattr(self, "label_resultado"):
            self.label_resultado.configure(text=f"{len(self.relatorios_encontrados)} relatorio(s) encontrado(s).")

    def limpar_filtros_relatorios(self):
        self.filtro_modelo.delete(0, "end")
        self.filtro_data.delete(0, "end")
        self.buscar_relatorios()

    def abrir_relatorio_selecionado(self, _event=None):
        selecionado = self.tree_relatorios.focus()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um relatorio.")
            return

        relatorio = self.relatorios_encontrados[int(selecionado)]
        try:
            abrir_relatorio(relatorio["caminho"])
        except Exception as erro:
            messagebox.showerror("Erro", str(erro))

    def exportar_lote_relatorios(self):
        if not self.relatorios_encontrados:
            messagebox.showwarning("Aviso", "Nenhum relatorio encontrado para exportar.")
            return

        pasta = filedialog.askdirectory(title="Escolha a pasta para exportar os relatorios filtrados")
        if not pasta:
            return

        try:
            copiados = exportar_relatorios_para_pasta(self.relatorios_encontrados, pasta)
            messagebox.showinfo("Sucesso", f"{copiados} relatorio(s) exportado(s) para:\n{pasta}")
        except Exception as erro:
            messagebox.showerror("Erro", str(erro))
