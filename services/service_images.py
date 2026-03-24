from __future__ import annotations

import shutil
from pathlib import Path

from utils.caminhos import caminho_imagens


def salvar_imagens(lista_caminhos, contrato):
    pasta_destino = Path(caminho_imagens(contrato))
    caminhos_salvos = []

    for caminho_origem in lista_caminhos:
        origem = Path(caminho_origem)
        if not origem.exists():
            continue

        destino = pasta_destino / origem.name
        contador = 1

        while destino.exists():
            destino = pasta_destino / f"{origem.stem}_{contador}{origem.suffix}"
            contador += 1

        shutil.copy2(origem, destino)
        caminhos_salvos.append(str(destino))

    return caminhos_salvos
