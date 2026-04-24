import customtkinter as ctk
from PIL import Image

from telas.TelaRelatorioFotografico import TelaRelatorioFotografico
from utils.caminhos import caminho_recurso
from utils.ui import (
    BG_APP,
    BG_CARD,
    BG_PANEL,
    BORDER,
    FONT_BODY,
    FONT_HEADING,
    FONT_LABEL,
    FONT_SMALL,
    FONT_TITLE,
    TEXT,
    TEXT_MUTED,
    configurar_aparencia,
    configurar_tabela_estilo,
    estilizar_botao,
)


class SistemaApp:
    def __init__(self):
        configurar_aparencia()
        configurar_tabela_estilo()

        self.app = ctk.CTk(fg_color=BG_APP)
        self.app.title("Relatorio Fotografico")
        self.app.geometry("1280x760")
        self.app.minsize(1100, 680)

        self.container = ctk.CTkFrame(self.app, fg_color=BG_APP)
        self.container.pack(fill="both", expand=True)

        self.menu_lateral = ctk.CTkFrame(
            self.container,
            width=260,
            fg_color=BG_PANEL,
            corner_radius=0,
            border_width=1,
            border_color=BORDER,
        )
        self.menu_lateral.pack(side="left", fill="y")
        self.menu_lateral.pack_propagate(False)

        self.frame = ctk.CTkFrame(self.container, fg_color=BG_APP, corner_radius=0)
        self.frame.pack(side="right", fill="both", expand=True)

        self.tela_foto = TelaRelatorioFotografico(self.app, self.frame, banco=None, voltar_menu=self.menu)
        self._criar_menu()
        self.menu()

    def _criar_menu(self):
        for widget in self.menu_lateral.winfo_children():
            widget.destroy()

        topo = ctk.CTkFrame(self.menu_lateral, fg_color="transparent")
        topo.pack(fill="x", padx=20, pady=(24, 18))

        self._carregar_logo(topo)
        ctk.CTkLabel(topo, text="Relatorio Fotografico", font=FONT_HEADING, text_color=TEXT).pack(anchor="w", pady=(10, 2))
        ctk.CTkLabel(
            topo,
            text="Versao enxuta para gerar, consultar e exportar relatorios fotograficos.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            wraplength=200,
            justify="left",
        ).pack(anchor="w")

        nav = ctk.CTkFrame(self.menu_lateral, fg_color="transparent")
        nav.pack(fill="x", padx=18, pady=(8, 18))

        botoes = [
            ("Inicio", self.menu, True),
            ("Gerar relatorio", self.tela_foto.abrir, False),
            ("Pesquisar relatorios", self.tela_foto.abrir_pesquisa, False),
        ]

        for texto, comando, primario in botoes:
            botao = ctk.CTkButton(nav, text=texto, command=comando, anchor="w")
            estilizar_botao(botao, primario=primario)
            botao.pack(fill="x", pady=6)

        rodape = ctk.CTkFrame(self.menu_lateral, fg_color=BG_CARD, corner_radius=16, border_width=1, border_color=BORDER)
        rodape.pack(side="bottom", fill="x", padx=18, pady=18)
        ctk.CTkLabel(rodape, text="Fluxo ativo", font=FONT_LABEL, text_color=TEXT).pack(anchor="w", padx=14, pady=(12, 2))
        ctk.CTkLabel(
            rodape,
            text="Geracao e historico de PDFs fotograficos.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            wraplength=210,
            justify="left",
        ).pack(anchor="w", padx=14, pady=(0, 12))

    def _carregar_logo(self, parent):
        for arquivo_logo in ("Logo.png", "assets/Logo.png"):
            try:
                self.logo = ctk.CTkImage(Image.open(caminho_recurso(arquivo_logo)), size=(122, 58))
                ctk.CTkLabel(parent, image=self.logo, text="").pack(anchor="w")
                return
            except Exception:
                continue

    def limpar(self):
        for widget in self.frame.winfo_children():
            widget.destroy()

    def menu(self):
        self.limpar()

        pagina = ctk.CTkFrame(self.frame, fg_color="transparent")
        pagina.pack(fill="both", expand=True, padx=26, pady=24)

        ctk.CTkLabel(pagina, text="Relatorio Fotografico", font=FONT_TITLE, text_color=TEXT).pack(anchor="w")
        ctk.CTkLabel(
            pagina,
            text="Este executavel foi reduzido para manter apenas a geracao e a consulta de relatorios fotograficos.",
            font=FONT_BODY,
            text_color=TEXT_MUTED,
            wraplength=860,
            justify="left",
        ).pack(anchor="w", pady=(4, 20))

        destaque = ctk.CTkFrame(pagina, fg_color=BG_PANEL, corner_radius=20, border_width=1, border_color=BORDER)
        destaque.pack(fill="x", pady=(0, 18))

        esquerda = ctk.CTkFrame(destaque, fg_color="transparent")
        esquerda.pack(side="left", fill="both", expand=True, padx=22, pady=22)
        ctk.CTkLabel(esquerda, text="Acoes disponiveis", font=FONT_HEADING, text_color=TEXT).pack(anchor="w")
        ctk.CTkLabel(
            esquerda,
            text="Selecione imagens, gere o PDF na pasta do modelo e abra o historico salvo sem depender dos outros modulos do sistema anterior.",
            font=FONT_BODY,
            text_color=TEXT_MUTED,
            wraplength=560,
            justify="left",
        ).pack(anchor="w", pady=(6, 18))

        acoes = ctk.CTkFrame(esquerda, fg_color="transparent")
        acoes.pack(anchor="w")
        for indice, (texto, comando, primario) in enumerate(
            [
                ("Gerar relatorio", self.tela_foto.abrir, True),
                ("Pesquisar relatorios", self.tela_foto.abrir_pesquisa, False),
            ]
        ):
            botao = ctk.CTkButton(acoes, text=texto, command=comando, width=190)
            estilizar_botao(botao, primario=primario)
            botao.grid(row=0, column=indice, padx=(0, 10), pady=4)

        direita = ctk.CTkFrame(destaque, fg_color=BG_CARD, corner_radius=18, border_width=1, border_color=BORDER, width=280)
        direita.pack(side="right", fill="y", padx=22, pady=22)
        direita.pack_propagate(False)
        ctk.CTkLabel(direita, text="Escopo", font=FONT_LABEL, text_color=TEXT).pack(anchor="w", padx=16, pady=(16, 6))
        ctk.CTkLabel(
            direita,
            text="Removidos os fluxos de manutencao, certificados, carteirinhas, laudos e planilha.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            wraplength=240,
            justify="left",
        ).pack(anchor="w", padx=16)

    def rodar(self):
        self.app.mainloop()
