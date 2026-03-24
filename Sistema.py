import customtkinter as ctk
from PIL import Image

from Telas import Telas
from banco.BancoDados import BancoDados
from telas.TelaRelatorioFotografico import TelaRelatorioFotografico
from utils.caminhos import caminho_banco, caminho_recurso


class SistemaApp:
    def __init__(self, arquivo=None):
        self.arquivo = arquivo or caminho_banco()
        self.banco = BancoDados(self.arquivo)

        self.app = ctk.CTk()
        self.app.title("Sistema de Manutenção")
        self.app.geometry("1000x600")

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.container = ctk.CTkFrame(self.app)
        self.container.pack(fill="both", expand=True)

        self.menu_lateral = ctk.CTkFrame(self.container, width=220)
        self.menu_lateral.pack(side="left", fill="y")
        self.menu_lateral.pack_propagate(False)

        self.frame = ctk.CTkFrame(self.container)
        self.frame.pack(side="right", fill="both", expand=True)

        self.telas = Telas(self.app, self.frame, self.banco, self.menu)
        self.tela_foto = TelaRelatorioFotografico(self.app, self.frame, self.banco, self.menu)

        self.criar_menu()
        self.menu()
        self._carregar_logo()

    def _carregar_logo(self):
        for arquivo_logo in ("Logo.png", "assets/Logo.png"):
            try:
                self.logo = ctk.CTkImage(
                    Image.open(caminho_recurso(arquivo_logo)),
                    size=(120, 60),
                )
                self.label_logo = ctk.CTkLabel(self.frame, image=self.logo, text="")
                self.label_logo.place(relx=0.98, rely=0.02, anchor="ne")
                return
            except Exception:
                continue

    def limpar(self):
        for widget in self.frame.winfo_children():
            widget.destroy()

    def menu(self):
        self.limpar()
        ctk.CTkLabel(
            self.frame,
            text="MENU PRINCIPAL",
            font=("Arial", 28, "bold"),
        ).pack(pady=30)

    def criar_menu(self):
        botoes = [
            ("Veículos", self.telas.tela_veiculos),
            ("Manutenção", self.telas.tela_manutencao),
            ("Relatório por Veículo", self.telas.tela_relatorio_veiculo),
            ("Relatórios", self.telas.tela_relatorios_manutencao),
            ("Relatório Fotográfico", self.tela_foto.abrir),
        ]

        for texto, comando in botoes:
            ctk.CTkButton(
                self.menu_lateral,
                text=texto,
                command=comando,
                width=180,
                height=40,
            ).pack(pady=8)

    def rodar(self):
        self.app.mainloop()
