from __future__ import annotations

from pathlib import Path


def preparar_envio_whatsapp(image_path: Path) -> Path:
    """Fase futura: devolve a imagem pronta sem automatizar login ou envio."""
    if not image_path.is_file():
        raise FileNotFoundError(image_path)
    return image_path
