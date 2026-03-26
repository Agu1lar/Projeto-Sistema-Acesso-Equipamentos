import customtkinter as ctk
from PIL import Image
from tkinter import filedialog, messagebox

from Telas import Telas
from banco.BancoDados import BancoDados
from services.service_certificados import listar_certificados_pendentes
from services.service_relatorios_fotograficos import listar_relatorios
from telas.TelaRelatorioFotografico import TelaRelatorioFotografico
from utils.caminhos import caminho_banco, caminho_recurso
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
    FONT_TITLE,
    SUCCESS,
    TEXT,
    TEXT_MUTED,
    configurar_aparencia,
    configurar_tabela_estilo,
    criar_cartao_info,
    estilizar_botao,
)


class SistemaApp:
    def __init__(self, arquivo=None):
        self.arquivo = arquivo
        self.banco = None
        self.telas = None
        self.tela_foto = None
        self._notificacao_certificados_exibida = False

        configurar_aparencia()
        configurar_tabela_estilo()

        self.app = ctk.CTk(fg_color=BG_APP)
        self.app.title("Sistema de Manutencao")
        self.app.geometry("1280x760")
        self.app.minsize(1160, 700)

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

        self._criar_menu()
        if self.arquivo:
            self._carregar_base(self.arquivo)
        else:
            self.tela_inicial()

    def _criar_menu(self):
        for widget in self.menu_lateral.winfo_children():
            widget.destroy()

        topo = ctk.CTkFrame(self.menu_lateral, fg_color="transparent")
        topo.pack(fill="x", padx=20, pady=(24, 18))

        self._carregar_logo(topo)
        ctk.CTkLabel(topo, text="Sistema de Manutencao", font=FONT_HEADING, text_color=TEXT).pack(anchor="w", pady=(10, 2))
        ctk.CTkLabel(
            topo,
            text="Operacao centralizada para manutencoes, certificados e relatorios.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            wraplength=200,
            justify="left",
        ).pack(anchor="w")

        nav = ctk.CTkFrame(self.menu_lateral, fg_color="transparent")
        nav.pack(fill="x", padx=18, pady=(8, 18))

        if self.banco:
            botoes = [
                ("Visao Geral", self.menu),
                ("Fila Unica", self.telas.tela_fila_unica),
                ("Manutencao", self.telas.tela_manutencao),
                ("Coleta Mobile", self.telas.tela_coleta_mobile),
                ("Certificados", self.telas.tela_certificados_funcionarios),
                ("Carteirinhas", self.telas.tela_carteirinhas_treinamento),
                ("Relatorios", self.telas.tela_relatorios_manutencao),
                ("Relatorio Fotografico", self.tela_foto.abrir),
            ]
        else:
            botoes = [
                ("Abrir base", self.abrir_base_existente),
                ("Criar base", self.criar_nova_base),
            ]

        for indice, (texto, comando) in enumerate(botoes):
            botao = ctk.CTkButton(nav, text=texto, command=comando, anchor="w")
            estilizar_botao(botao, primario=indice == 0)
            botao.pack(fill="x", pady=6)

        rodape = ctk.CTkFrame(self.menu_lateral, fg_color=BG_CARD, corner_radius=16, border_width=1, border_color=BORDER)
        rodape.pack(side="bottom", fill="x", padx=18, pady=18)
        ctk.CTkLabel(rodape, text="Base ativa", font=FONT_LABEL, text_color=TEXT).pack(anchor="w", padx=14, pady=(12, 2))
        ctk.CTkLabel(rodape, text=self.arquivo or "Nenhuma base selecionada", font=FONT_SMALL, text_color=TEXT_MUTED, wraplength=210, justify="left").pack(
            anchor="w", padx=14, pady=(0, 12)
        )

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

    def _carregar_base(self, arquivo):
        self.arquivo = arquivo or caminho_banco()
        self.banco = BancoDados(self.arquivo)
        self.telas = Telas(self.app, self.frame, self.banco, self.menu)
        self.tela_foto = TelaRelatorioFotografico(self.app, self.frame, self.banco, self.menu)
        self._criar_menu()
        self.menu()

    def _notificar_certificados_pendentes(self):
        if not self.banco or self._notificacao_certificados_exibida:
            return

        pendentes = listar_certificados_pendentes(self.banco)
        if pendentes.empty:
            return

        self._notificacao_certificados_exibida = True
        destaques = []
        for _, linha in pendentes.head(5).iterrows():
            nome = str(linha.get("NOME", "Sem nome")).strip() or "Sem nome"
            tipo = str(linha.get("TIPO_CERTIFICADO", "Sem tipo")).strip() or "Sem tipo"
            vencimento = str(linha.get("DATA_VENCIMENTO", "")).strip() or "sem data"
            destaques.append(f"{nome} | {tipo} | vence em {vencimento}")

        complemento = "\n..." if len(pendentes) > 5 else ""
        messagebox.showwarning(
            "Certificados pendentes",
            f"Existem {len(pendentes)} certificado(s) vencido(s) ou a vencer em ate 30 dias.\n\n"
            f"{chr(10).join(destaques)}{complemento}",
            parent=self.app,
        )

    def abrir_base_existente(self):
        arquivo = filedialog.askopenfilename(
            parent=self.app,
            title="Selecionar planilha",
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if arquivo:
            self._carregar_base(arquivo)

    def criar_nova_base(self):
        arquivo = filedialog.asksaveasfilename(
            parent=self.app,
            title="Criar planilha",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if not arquivo:
            return

        if not messagebox.askyesno("Criar", "Criar nova planilha?", parent=self.app):
            return

        BancoDados(arquivo).criar_estrutura()
        self._carregar_base(arquivo)

    def tela_inicial(self):
        self.limpar()

        pagina = ctk.CTkFrame(self.frame, fg_color="transparent")
        pagina.pack(fill="both", expand=True, padx=26, pady=24)

        ctk.CTkLabel(pagina, text="Selecionar Base", font=FONT_TITLE, text_color=TEXT).pack(anchor="w")
        ctk.CTkLabel(
            pagina,
            text="Abra uma planilha existente ou crie uma nova base direto na aplicacao.",
            font=FONT_BODY,
            text_color=TEXT_MUTED,
        ).pack(anchor="w", pady=(4, 20))

        painel = ctk.CTkFrame(pagina, fg_color=BG_PANEL, corner_radius=20, border_width=1, border_color=BORDER)
        painel.pack(fill="x", pady=(0, 18))
        esquerda = ctk.CTkFrame(painel, fg_color="transparent")
        esquerda.pack(side="left", fill="both", expand=True, padx=22, pady=22)
        ctk.CTkLabel(esquerda, text="Inicio rapido", font=FONT_HEADING, text_color=TEXT).pack(anchor="w")
        ctk.CTkLabel(
            esquerda,
            text="Ao selecionar a base, o menu lateral e todas as telas da aplicacao sao habilitados.",
            font=FONT_BODY,
            text_color=TEXT_MUTED,
            wraplength=560,
            justify="left",
        ).pack(anchor="w", pady=(6, 18))

        acoes = ctk.CTkFrame(esquerda, fg_color="transparent")
        acoes.pack(anchor="w")
        for indice, (texto, comando, primario) in enumerate(
            [
                ("Abrir planilha", self.abrir_base_existente, True),
                ("Criar planilha", self.criar_nova_base, False),
            ]
        ):
            botao = ctk.CTkButton(acoes, text=texto, command=comando, width=190)
            estilizar_botao(botao, primario=primario)
            botao.grid(row=0, column=indice, padx=(0, 10), pady=4)

    def menu(self):
        self.limpar()

        pagina = ctk.CTkFrame(self.frame, fg_color="transparent")
        pagina.pack(fill="both", expand=True, padx=26, pady=24)

        ctk.CTkLabel(pagina, text="Painel Operacional", font=FONT_TITLE, text_color=TEXT).pack(anchor="w")
        ctk.CTkLabel(
            pagina,
            text="Acompanhe o status da base e entre rapidamente nos fluxos mais usados.",
            font=FONT_BODY,
            text_color=TEXT_MUTED,
        ).pack(anchor="w", pady=(4, 18))

        linha_metricas = ctk.CTkFrame(pagina, fg_color="transparent")
        linha_metricas.pack(fill="x", pady=(0, 18))
        linha_metricas.grid_columnconfigure((0, 1, 2, 3), weight=1)

        manutencoes = len(self.banco.carregar_dataframe("MANUTENCOES"))
        carteirinhas = len(self.banco.carregar_dataframe("TREINAMENTOS"))
        certificados_pendentes = len(listar_certificados_pendentes(self.banco))
        relatorios_foto = len(listar_relatorios())

        metricas = [
            ("Manutencoes registradas", str(manutencoes), ACCENT),
            ("Carteirinhas emitidas", str(carteirinhas), TEXT),
            ("Certificados pendentes", str(certificados_pendentes), "#ef4444" if certificados_pendentes else TEXT),
            ("Relatorios fotograficos", str(relatorios_foto), SUCCESS),
        ]

        for indice, (titulo, valor, cor) in enumerate(metricas):
            card = criar_cartao_info(linha_metricas, titulo, valor, cor)
            card.grid(row=0, column=indice, sticky="nsew", padx=(0 if indice == 0 else 10, 0))

        destaque = ctk.CTkFrame(pagina, fg_color=BG_PANEL, corner_radius=20, border_width=1, border_color=BORDER)
        destaque.pack(fill="x", pady=(0, 18))

        esquerda = ctk.CTkFrame(destaque, fg_color="transparent")
        esquerda.pack(side="left", fill="both", expand=True, padx=22, pady=22)
        ctk.CTkLabel(esquerda, text="Fluxo recomendado", font=FONT_HEADING, text_color=TEXT).pack(anchor="w")
        ctk.CTkLabel(
            esquerda,
            text="Lance a manutencao, acompanhe certificados e gere os relatorios sem voltar para a planilha.",
            font=FONT_BODY,
            text_color=TEXT_MUTED,
            wraplength=560,
            justify="left",
        ).pack(anchor="w", pady=(6, 18))

        acoes = ctk.CTkFrame(esquerda, fg_color="transparent")
        acoes.pack(anchor="w")
        for indice, (texto, comando, primario) in enumerate(
            [
                ("Fila unica", self.telas.tela_fila_unica, True),
                ("Nova manutencao", self.telas.tela_manutencao, False),
                ("Coleta mobile", self.telas.tela_coleta_mobile, False),
                ("Novo certificado", self.telas.tela_certificados_funcionarios, False),
                ("Nova carteirinha", self.telas.tela_carteirinhas_treinamento, False),
                ("Relatorio fotografico", self.tela_foto.abrir, False),
            ]
        ):
            botao = ctk.CTkButton(acoes, text=texto, command=comando, width=190)
            estilizar_botao(botao, primario=primario)
            botao.grid(row=0, column=indice, padx=(0, 10), pady=4)

        direita = ctk.CTkFrame(destaque, fg_color=BG_CARD, corner_radius=18, border_width=1, border_color=BORDER, width=280)
        direita.pack(side="right", fill="y", padx=22, pady=22)
        direita.pack_propagate(False)
        ctk.CTkLabel(direita, text="Base protegida", font=FONT_LABEL, text_color=TEXT).pack(anchor="w", padx=16, pady=(16, 6))
        ctk.CTkLabel(
            direita,
            text="A aplicacao tenta reparar colunas, recriar abas obrigatorias e gerar backup antes de sobrescrever a planilha.",
            font=FONT_SMALL,
            text_color=TEXT_MUTED,
            wraplength=240,
            justify="left",
        ).pack(anchor="w", padx=16)

        atalhos = ctk.CTkFrame(pagina, fg_color="transparent")
        atalhos.pack(fill="both", expand=True)
        atalhos.grid_columnconfigure((0, 1), weight=1)

        for indice, (titulo, descricao, comando) in enumerate(
            [
                ("Consulta rapida", "Abra a base em modo de visualizacao, filtre, edite e exporte PDF.", lambda: self.telas.visualizar("MANUTENCOES")),
                ("Historico fotografico", "Pesquise relatorios por modelo e data sem navegar por pastas.", self.tela_foto.abrir_pesquisa),
            ]
        ):
            card = ctk.CTkFrame(atalhos, fg_color=BG_PANEL, corner_radius=18, border_width=1, border_color=BORDER)
            card.grid(row=0, column=indice, sticky="nsew", padx=(0 if indice == 0 else 10, 0))
            ctk.CTkLabel(card, text=titulo, font=FONT_HEADING, text_color=TEXT).pack(anchor="w", padx=18, pady=(18, 8))
            ctk.CTkLabel(card, text=descricao, font=FONT_BODY, text_color=TEXT_MUTED, wraplength=420, justify="left").pack(
                anchor="w", padx=18
            )
            botao = ctk.CTkButton(card, text="Abrir", command=comando, width=140)
            estilizar_botao(botao)
            botao.pack(anchor="w", padx=18, pady=18)

        rodape_metricas = ctk.CTkFrame(pagina, fg_color="transparent")
        rodape_metricas.pack(fill="x", pady=(18, 0))
        criar_cartao_info(
            rodape_metricas,
            "Relatorios fotograficos",
            str(relatorios_foto),
            TEXT,
        ).pack(fill="x")
        self.app.after(150, self._notificar_certificados_pendentes)

    def rodar(self):
        self.app.mainloop()
