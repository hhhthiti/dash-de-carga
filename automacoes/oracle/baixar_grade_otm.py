from __future__ import annotations

import os
import shutil
from pathlib import Path


def baixar_grade(destino: Path) -> Path:
    mode = os.getenv("OTM_MODE", "manual").strip().lower()
    if mode != "manual":
        raise RuntimeError(
            "Download OTM automatico ainda nao foi habilitado. "
            "Use OTM_MODE=manual ou implemente a integracao somente com credenciais autorizadas."
        )

    origem = Path(os.environ["OTM_GRADE_FILE"]).expanduser().resolve()
    if not origem.is_file():
        raise FileNotFoundError(f"Arquivo local da grade nao encontrado: {origem}")
    if origem.stat().st_size == 0:
        raise ValueError(f"Arquivo local da grade esta vazio: {origem}")

    destino_real = destino.with_suffix(origem.suffix.lower())
    destino_real.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(origem, destino_real)
    return destino_real
