import os
from tkinter import filedialog, messagebox

from banco.BancoDados import BancoDados
from Sistema import SistemaApp


def selecionar_arquivo():
    return filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
    )


def preparar_banco(arquivo):
    if os.path.exists(arquivo):
        return True

    if messagebox.askyesno("Criar", "Criar nova planilha?"):
        BancoDados(arquivo).criar_estrutura()
        return True

    return False


def main():
    arquivo = selecionar_arquivo()
    if not arquivo:
        return

    if not preparar_banco(arquivo):
        return

    app = SistemaApp(arquivo)
    app.rodar()


if __name__ == "__main__":
    main()
