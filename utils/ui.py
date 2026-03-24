from __future__ import annotations

from tkinter import Toplevel
from tkinter import ttk

import customtkinter as ctk

BG_APP = "#0f172a"
BG_PANEL = "#111827"
BG_CARD = "#1e293b"
BG_CARD_ALT = "#243041"
ACCENT = "#f97316"
ACCENT_HOVER = "#ea580c"
TEXT = "#f8fafc"
TEXT_MUTED = "#94a3b8"
BORDER = "#334155"
SUCCESS = "#10b981"

FONT_TITLE = ("Segoe UI Semibold", 28)
FONT_SUBTITLE = ("Segoe UI", 13)
FONT_HEADING = ("Segoe UI Semibold", 20)
FONT_LABEL = ("Segoe UI Semibold", 12)
FONT_BODY = ("Segoe UI", 12)
FONT_SMALL = ("Segoe UI", 11)


def configurar_aparencia():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")


def configurar_tabela_estilo():
    style = ttk.Style()
    try:
        style.theme_use("default")
    except Exception:
        pass

    style.configure(
        "Treeview",
        background=BG_PANEL,
        foreground=TEXT,
        fieldbackground=BG_PANEL,
        rowheight=34,
        borderwidth=0,
        font=FONT_BODY,
    )
    style.map("Treeview", background=[("selected", BG_CARD_ALT)], foreground=[("selected", TEXT)])
    style.configure(
        "Treeview.Heading",
        background=BG_CARD_ALT,
        foreground=TEXT,
        relief="flat",
        borderwidth=0,
        font=FONT_LABEL,
        padding=8,
    )
    style.map("Treeview.Heading", background=[("active", BG_CARD)])


def criar_pagina(parent):
    frame = ctk.CTkFrame(parent, fg_color="transparent")
    frame.pack(fill="both", expand=True, padx=24, pady=24)
    return frame


def criar_cabecalho(parent, titulo: str, descricao: str = ""):
    frame = ctk.CTkFrame(parent, fg_color="transparent")
    frame.pack(fill="x", pady=(0, 18))

    ctk.CTkLabel(frame, text=titulo, font=FONT_TITLE, text_color=TEXT).pack(anchor="w")
    if descricao:
        ctk.CTkLabel(frame, text=descricao, font=FONT_SUBTITLE, text_color=TEXT_MUTED).pack(anchor="w", pady=(4, 0))
    return frame


def criar_secao(parent, titulo: str = "", descricao: str = "", expand: bool = False):
    secao = ctk.CTkFrame(
        parent,
        fg_color=BG_PANEL,
        corner_radius=18,
        border_width=1,
        border_color=BORDER,
    )
    secao.pack(fill="both", expand=expand, pady=(0, 16))

    if titulo or descricao:
        topo = ctk.CTkFrame(secao, fg_color="transparent")
        topo.pack(fill="x", padx=18, pady=(16, 8))
        if titulo:
            ctk.CTkLabel(topo, text=titulo, font=FONT_HEADING, text_color=TEXT).pack(anchor="w")
        if descricao:
            ctk.CTkLabel(topo, text=descricao, font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", pady=(4, 0))

    corpo = ctk.CTkFrame(secao, fg_color="transparent")
    corpo.pack(fill="both", expand=True, padx=18, pady=(0, 18))
    return secao, corpo


def criar_cartao_info(parent, titulo: str, valor: str, destaque: str = TEXT_MUTED):
    card = ctk.CTkFrame(
        parent,
        fg_color=BG_CARD,
        corner_radius=16,
        border_width=1,
        border_color=BORDER,
    )
    ctk.CTkLabel(card, text=titulo, font=FONT_SMALL, text_color=TEXT_MUTED).pack(anchor="w", padx=16, pady=(14, 6))
    ctk.CTkLabel(card, text=valor, font=("Segoe UI Semibold", 24), text_color=destaque).pack(
        anchor="w", padx=16, pady=(0, 14)
    )
    return card


def estilizar_botao(botao, primario: bool = False):
    botao.configure(
        height=40,
        corner_radius=12,
        font=FONT_LABEL,
        fg_color=ACCENT if primario else BG_CARD_ALT,
        hover_color=ACCENT_HOVER if primario else BG_CARD,
        text_color=TEXT,
        border_width=0,
    )
    return botao


def adicionar_tooltip(widget, texto: str):
    tooltip = {"janela": None}

    def mostrar(_event=None):
        if tooltip["janela"] is not None:
            return
        x = widget.winfo_rootx() + 18
        y = widget.winfo_rooty() + widget.winfo_height() + 8
        janela = Toplevel(widget)
        janela.wm_overrideredirect(True)
        janela.wm_geometry(f"+{x}+{y}")
        janela.configure(bg=BG_PANEL)
        label = ctk.CTkLabel(
            janela,
            text=texto,
            font=FONT_SMALL,
            text_color=TEXT,
            fg_color=BG_PANEL,
            corner_radius=8,
            padx=10,
            pady=6,
        )
        label.pack()
        tooltip["janela"] = janela

    def esconder(_event=None):
        if tooltip["janela"] is not None:
            tooltip["janela"].destroy()
            tooltip["janela"] = None

    widget.bind("<Enter>", mostrar, add="+")
    widget.bind("<Leave>", esconder, add="+")
    widget.bind("<ButtonPress>", esconder, add="+")
    return widget


def criar_label_form(parent, texto: str):
    return ctk.CTkLabel(parent, text=texto, font=FONT_LABEL, text_color=TEXT)


def estilizar_entry(entry):
    entry.configure(
        height=40,
        corner_radius=12,
        fg_color=BG_CARD,
        border_color=BORDER,
        text_color=TEXT,
        placeholder_text_color=TEXT_MUTED,
        font=FONT_BODY,
    )
    return entry


def estilizar_combo(combo):
    combo.configure(
        height=40,
        corner_radius=12,
        fg_color=BG_CARD,
        border_color=BORDER,
        button_color=BG_CARD_ALT,
        button_hover_color=BG_PANEL,
        text_color=TEXT,
        dropdown_fg_color=BG_PANEL,
        dropdown_text_color=TEXT,
        dropdown_hover_color=BG_CARD_ALT,
        font=FONT_BODY,
    )
    return combo
