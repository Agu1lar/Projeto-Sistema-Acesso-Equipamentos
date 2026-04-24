from __future__ import annotations

import re
import shutil
from datetime import datetime
from pathlib import Path

from utils.caminhos import caminho_imagens


def _slug(texto: str) -> str:
    valor = re.sub(r"[^\w\-\.]+", "_", str(texto or "").strip())
    return valor.strip("._") or "arquivo"


def _titulo_imagem(caminho_origem: Path) -> str:
    return caminho_origem.stem.replace("_", " ").replace("-", " ").strip() or "Imagem"


def salvar_imagens(lista_caminhos, contrato):
    data_base = datetime.now().strftime("%Y-%m-%d")
    pasta_destino = Path(caminho_imagens(Path(_slug(contrato)) / data_base))
    imagens_salvas = []

    for caminho_origem in lista_caminhos:
        origem = Path(caminho_origem)
        if not origem.exists():
            continue

        timestamp = datetime.fromtimestamp(origem.stat().st_mtime).strftime("%Y%m%d_%H%M%S")
        nome_base = f"{_slug(contrato)}_{_slug(origem.stem)}_{timestamp}"
        destino = pasta_destino / f"{nome_base}{origem.suffix.lower()}"
        contador = 1

        while destino.exists():
            destino = pasta_destino / f"{nome_base}_{contador}{origem.suffix.lower()}"
            contador += 1

        shutil.copy2(origem, destino)
        imagens_salvas.append(
            {
                "caminho": str(destino),
                "titulo": _titulo_imagem(origem),
            }
        )

    return imagens_salvas
