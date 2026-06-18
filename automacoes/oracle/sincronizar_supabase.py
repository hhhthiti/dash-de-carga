from __future__ import annotations

import os
from typing import Any

import requests

from common import chunks


class SupabaseClient:
    """Fallback direto. O fluxo recomendado usa WorkerClient."""

    def __init__(self) -> None:
        base = os.environ["SUPABASE_URL"].rstrip("/")
        key = os.environ["SUPABASE_SERVICE_ROLE_KEY"].strip()
        if not key:
            raise ValueError("SUPABASE_SERVICE_ROLE_KEY vazio.")
        self.rest_url = f"{base}/rest/v1"
        self.session = requests.Session()
        self.session.headers.update({
            "apikey": key,
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
        })

    def upsert_cargas(self, rows: list[dict[str, Any]]) -> None:
        if not rows:
            raise ValueError("Supabase fallback recusou cargas vazias.")
        for batch in chunks(rows, 100):
            response = self.session.post(
                f"{self.rest_url}/reporte_carga?on_conflict=dt,data_ref",
                json=batch,
                headers={"Prefer": "resolution=merge-duplicates,return=minimal"},
                timeout=(10, 90),
            )
            response.raise_for_status()
