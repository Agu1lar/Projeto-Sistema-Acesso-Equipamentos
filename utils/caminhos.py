from pathlib import Path
import sys


def caminho_base() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def caminho_dados() -> Path:
    pasta = Path.home() / "Documents" / "SistemaManutencao"
    pasta.mkdir(parents=True, exist_ok=True)
    return pasta


def caminho_banco(nome_arquivo: str = "banco_dados.xlsx") -> str:
    return str(caminho_dados() / nome_arquivo)


def caminho_imagens(contrato=None) -> str:
    base = caminho_dados() / "imagens"
    pasta = base / str(contrato) if contrato else base
    pasta.mkdir(parents=True, exist_ok=True)
    return str(pasta)


def caminho_relatorios_fotograficos() -> str:
    pasta = caminho_dados() / "relatorios_fotograficos"
    pasta.mkdir(parents=True, exist_ok=True)
    return str(pasta)


def caminho_config_relatorios() -> str:
    return str(caminho_dados() / "config_relatorios_fotograficos.json")


def caminho_config_notificacoes() -> str:
    return str(caminho_dados() / "config_notificacoes.json")


def caminho_mobile_sync() -> str:
    pasta = caminho_dados() / "mobile_sync"
    pasta.mkdir(parents=True, exist_ok=True)
    return str(pasta)


def caminho_mobile_pendencias() -> str:
    pasta = Path(caminho_mobile_sync()) / "pendentes"
    pasta.mkdir(parents=True, exist_ok=True)
    return str(pasta)


def caminho_mobile_importados() -> str:
    pasta = Path(caminho_mobile_sync()) / "importados"
    pasta.mkdir(parents=True, exist_ok=True)
    return str(pasta)


def caminho_recurso(relativo: str) -> str:
    try:
        base = Path(sys._MEIPASS)
    except AttributeError:
        base = caminho_base()

    return str(base / relativo)
