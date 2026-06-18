from __future__ import annotations

import os
import time
from typing import Any

import requests


class WorkerClient:
    def __init__(self) -> None:
        self.base_url = os.environ["WORKER_BASE_URL"].rstrip("/")
        self.secret = os.environ["ORACLE_JOB_SECRET"].strip()
        if not self.secret:
            raise ValueError("ORACLE_JOB_SECRET vazio.")
        self.session = requests.Session()
        self.session.headers.update({
            "Content-Type": "application/json",
            "X-Oracle-Job-Secret": self.secret,
        })

    def post(self, path: str, payload: dict[str, Any], attempts: int = 3) -> dict[str, Any]:
        last_error: Exception | None = None
        for attempt in range(attempts):
            try:
                response = self.session.post(f"{self.base_url}{path}", json=payload, timeout=(10, 90))
                if response.status_code >= 500 or response.status_code == 429:
                    raise RuntimeError(f"Worker HTTP {response.status_code}: {response.text[:500]}")
                response.raise_for_status()
                return response.json()
            except (requests.RequestException, RuntimeError) as exc:
                last_error = exc
                if attempt + 1 < attempts:
                    time.sleep(2 ** attempt)
        raise RuntimeError(f"Falha ao chamar {path}: {last_error}")

    def sync_grade(self, run_id: str, cargas: list[dict[str, Any]], minimum_expected: int) -> dict[str, Any]:
        return self.post("/api/oracle/sync-grade", {
            "run_id": run_id,
            "minimum_expected": minimum_expected,
            "mark_missing": True,
            "cargas": cargas,
        })

    def sync_zles002(
        self,
        run_id: str,
        cargas: list[dict[str, Any]],
        materiais: list[dict[str, Any]],
    ) -> dict[str, Any]:
        return self.post("/api/oracle/sync-zles002", {
            "run_id": run_id,
            "cargas": cargas,
            "materiais": materiais,
        })

    def job_status(self, payload: dict[str, Any]) -> dict[str, Any]:
        return self.post("/api/oracle/job-status", payload)

    def report_image(self, payload: dict[str, Any]) -> dict[str, Any]:
        return self.post("/api/oracle/report-image", payload)
