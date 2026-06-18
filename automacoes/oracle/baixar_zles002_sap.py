from __future__ import annotations

import os
import shutil
from pathlib import Path


def baixar_zles002(destino: Path, dts: list[str]) -> Path:
    if not dts:
        raise ValueError("A ZLES002 nao pode ser solicitada sem DTs validas.")

    mode = os.getenv("SAP_MODE", "manual").strip().lower()
    if mode != "manual":
        raise RuntimeError(
            "Automacao SAP/Citrix ainda nao foi habilitada. "
            "Use SAP_MODE=manual ou implemente a integracao somente com credenciais autorizadas."
        )

    origem = Path(os.environ["SAP_ZLES002_FILE"]).expanduser().resolve()
    if not origem.is_file():
        raise FileNotFoundError(f"Arquivo local da ZLES002 nao encontrado: {origem}")
    if origem.stat().st_size == 0:
        raise ValueError(f"Arquivo local da ZLES002 esta vazio: {origem}")

    destino_real = destino.with_suffix(origem.suffix.lower())
    destino_real.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(origem, destino_real)
    return destino_real
