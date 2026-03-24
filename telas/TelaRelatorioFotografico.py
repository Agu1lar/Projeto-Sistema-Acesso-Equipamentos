from __future__ import annotations

from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from services.service_images import salvar_imagens
from services.service_pdf import gerar_relatorio
from services.service_relatorios_fotograficos import (
    abrir_relatorio,
    caminho_pdf_modelo,
    listar_relatorios,
    obter_pasta_base,
    salvar_pasta_base,
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

        container = ctk.CTkFrame(self.frame)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            container,
            text="RELATÓRIO FOTOGRÁFICO",
            font=("Arial", 22, "bold"),
        ).pack(pady=10)

        for placeholder, chave in [
            ("Cliente", "cliente"),
            ("Obra", "obra"),
            ("Contrato", "contrato"),
            ("Equipamento", "equipamento"),
            ("Modelo", "modelo"),
            ("Patrimônio", "patrimonio"),
            ("Número de Série", "serie"),
        ]:
            campo = ctk.CTkEntry(container, placeholder_text=placeholder, width=320)
            campo.pack(pady=5)
            self.campos[chave] = campo

        info_pasta = f"Pasta base: {obter_pasta_base()}"
        self.label_pasta = ctk.CTkLabel(container, text=info_pasta, wraplength=700)
        self.label_pasta.pack(pady=(10, 4))

        ctk.CTkCheckBox(
            container,
            text="Salvar automaticamente na pasta do modelo",
            variable=self.salvar_automatico,
        ).pack(pady=4)

        acoes = ctk.CTkFrame(container, fg_color="transparent")
        acoes.pack(pady=6)

        ctk.CTkButton(acoes, text="Definir Pasta Base", command=self.definir_pasta_base).grid(row=0, column=0, padx=5)
        ctk.CTkButton(acoes, text="Pesquisar Relatórios", command=self.abrir_pesquisa).grid(row=0, column=1, padx=5)

        self.label_imagens = ctk.CTkLabel(container, text="Nenhuma imagem selecionada")
        self.label_imagens.pack(pady=10)

        ctk.CTkButton(
            container,
            text="Selecionar Imagens",
            command=self.selecionar_imagens,
        ).pack(pady=5)

        ctk.CTkButton(
            container,
            text="Gerar Relatório",
            command=self.gerar,
        ).pack(pady=10)

        ctk.CTkButton(
            container,
            text="Voltar",
            command=self.voltar_menu,
        ).pack(pady=10)

    def abrir_pesquisa(self):
        self.limpar()

        container = ctk.CTkFrame(self.frame)
        container.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            container,
            text="PESQUISA DE RELATÓRIOS",
            font=("Arial", 22, "bold"),
        ).pack(pady=10)

        filtros = ctk.CTkFrame(container)
        filtros.pack(fill="x", padx=10, pady=10)

        self.filtro_modelo = ctk.CTkEntry(filtros, placeholder_text="Modelo", width=250)
        self.filtro_modelo.grid(row=0, column=0, padx=5, pady=5)

        self.filtro_data = ctk.CTkEntry(filtros, placeholder_text="Data (dd/mm/aaaa)", width=180)
        self.filtro_data.grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkButton(filtros, text="Buscar", command=self.buscar_relatorios).grid(row=0, column=2, padx=5, pady=5)
        ctk.CTkButton(filtros, text="Limpar", command=self.limpar_filtros_relatorios).grid(row=0, column=3, padx=5, pady=5)
        ctk.CTkButton(filtros, text="Abrir Tela de Geração", command=self.abrir).grid(row=0, column=4, padx=5, pady=5)

        self.label_resultado = ctk.CTkLabel(
            container,
            text="Filtre por modelo, por data ou apenas por data para ver todos os relatórios do dia.",
        )
        self.label_resultado.pack(pady=5)

        tabela_frame = ctk.CTkFrame(container)
        tabela_frame.pack(fill="both", expand=True, padx=10, pady=10)

        colunas = ("modelo", "data", "arquivo")
        self.tree_relatorios = ttk.Treeview(tabela_frame, columns=colunas, show="headings")
        self.tree_relatorios.heading("modelo", text="Modelo")
        self.tree_relatorios.heading("data", text="Data")
        self.tree_relatorios.heading("arquivo", text="Arquivo")
        self.tree_relatorios.column("modelo", width=180, anchor="center")
        self.tree_relatorios.column("data", width=140, anchor="center")
        self.tree_relatorios.column("arquivo", width=360, anchor="w")
        self.tree_relatorios.pack(side="left", fill="both", expand=True)
        self.tree_relatorios.bind("<Double-1>", self.abrir_relatorio_selecionado)

        scroll_y = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tree_relatorios.yview)
        scroll_y.pack(side="right", fill="y")
        self.tree_relatorios.configure(yscrollcommand=scroll_y.set)

        botoes = ctk.CTkFrame(container, fg_color="transparent")
        botoes.pack(pady=5)

        ctk.CTkButton(botoes, text="Abrir Selecionado", command=self.abrir_relatorio_selecionado).grid(
            row=0, column=0, padx=5
        )
        ctk.CTkButton(botoes, text="Voltar", command=self.voltar_menu).grid(row=0, column=1, padx=5)

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
        self.label_imagens.configure(text=f"{len(caminhos)} imagem(ns) selecionada(s)")

    def definir_pasta_base(self):
        pasta = filedialog.askdirectory(title="Escolha a pasta base dos relatórios por modelo")
        if not pasta:
            return

        salvar_pasta_base(pasta)
        self.label_pasta.configure(text=f"Pasta base: {obter_pasta_base()}")
        messagebox.showinfo("Sucesso", "Pasta base salva com sucesso.")

    def gerar(self):
        contrato = self.campos["contrato"].get().strip()
        modelo = self.campos["modelo"].get().strip()

        if not contrato:
            messagebox.showwarning("Aviso", "Informe o contrato")
            return

        if not modelo:
            messagebox.showwarning("Aviso", "Informe o modelo da máquina")
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
            "Série": self.campos["serie"].get().strip(),
            "Patrimônio": self.campos["patrimonio"].get().strip(),
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
                    title="Salvar relatório como",
                )

                if not caminho_pdf:
                    return

            gerar_relatorio(dados, imagens_salvas, caminho_pdf)
            messagebox.showinfo("Sucesso", f"Relatório gerado:\n{caminho_pdf}")
        except Exception as erro:
            messagebox.showerror("Erro", str(erro))

    def buscar_relatorios(self):
        try:
            modelo = self.filtro_modelo.get().strip() if hasattr(self, "filtro_modelo") else ""
            data = self.filtro_data.get().strip() if hasattr(self, "filtro_data") else ""
            self.relatorios_encontrados = listar_relatorios(modelo=modelo, data=data)
        except ValueError as erro:
            messagebox.showwarning("Filtro inválido", str(erro))
            return
        except Exception as erro:
            messagebox.showerror("Erro", str(erro))
            return

        self.tree_relatorios.delete(*self.tree_relatorios.get_children())

        for indice, relatorio in enumerate(self.relatorios_encontrados):
            self.tree_relatorios.insert(
                "",
                "end",
                iid=str(indice),
                values=(relatorio["modelo"], relatorio["data"], relatorio["arquivo"]),
            )

        total = len(self.relatorios_encontrados)
        self.label_resultado.configure(text=f"{total} relatório(s) encontrado(s).")

    def limpar_filtros_relatorios(self):
        self.filtro_modelo.delete(0, "end")
        self.filtro_data.delete(0, "end")
        self.buscar_relatorios()

    def abrir_relatorio_selecionado(self, _event=None):
        selecionado = self.tree_relatorios.focus()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um relatório.")
            return

        relatorio = self.relatorios_encontrados[int(selecionado)]
        try:
            abrir_relatorio(relatorio["caminho"])
        except Exception as erro:
            messagebox.showerror("Erro", str(erro))
