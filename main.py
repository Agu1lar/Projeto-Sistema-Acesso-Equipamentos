import os

from PIL import Image

from Sistema import SistemaApp
from services.service_pdf import _localizar_logo


def executar_smoke_test() -> int:
    # Valida imports e recursos essenciais antes de abrir a interface.
    caminho_logo = _localizar_logo()
    if not caminho_logo or not os.path.exists(caminho_logo):
        raise FileNotFoundError("Logo.png nao foi encontrada no executavel.")

    with Image.open(caminho_logo) as imagem_logo:
        imagem_logo.verify()

    return 0


def main():
    if os.environ.get("RELATORIO_SMOKE_TEST") == "1":
        raise SystemExit(executar_smoke_test())

    app = SistemaApp()
    app.rodar()


if __name__ == "__main__":
    raise SystemExit(main())
