from __future__ import annotations

import shutil
import re
from datetime import datetime
from pathlib import Path

from utils.caminhos import caminho_imagens


def _slug(texto: str) -> str:
    valor = re.sub(r"[^\w\-\.]+", "_", str(texto or "").strip())
    return valor.strip("._") or "arquivo"


def salvar_imagens(lista_caminhos, contrato):
    data_base = datetime.now().strftime("%Y-%m-%d")
    pasta_destino = Path(caminho_imagens(Path(_slug(contrato)) / data_base))
    caminhos_salvos = []

    for caminho_origem in lista_caminhos:
        origem = Path(caminho_origem)
        if not origem.exists():
            continue

        timestamp = datetime.fromtimestamp(origem.stat().st_mtime).strftime("%Y%m%d_%H%M%S")
        destino = pasta_destino / f"{_slug(contrato)}_{timestamp}{origem.suffix.lower()}"
        contador = 1

        while destino.exists():
            destino = pasta_destino / f"{_slug(contrato)}_{timestamp}_{contador}{origem.suffix.lower()}"
            contador += 1

        shutil.copy2(origem, destino)
        caminhos_salvos.append(str(destino))

    return caminhos_salvos
