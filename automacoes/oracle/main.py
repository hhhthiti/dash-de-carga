from __future__ import annotations

import argparse
import logging
import os
import sys
import uuid
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from baixar_grade_otm import baixar_grade
from baixar_zles002_sap import baixar_zles002
from common import configure_logging, env_bool, write_json
from gerar_imagem_reporte import gerar_imagem_reporte, imagem_base64
from processar_grade import processar_grade
from processar_zles002 import processar_zles002
from sincronizar_supabase import SupabaseClient
from sincronizar_worker import WorkerClient


LOGGER = logging.getLogger("oracle.main")


def _now() -> datetime:
    return datetime.now(ZoneInfo(os.getenv("TIME_ZONE", "America/Sao_Paulo")))


def _status_payload(
    run_id: str,
    status: str,
    stage: str,
    message: str,
    started_at: datetime,
    grade_count: int = 0,
    zles_count: int = 0,
    material_count: int = 0,
    details: dict | None = None,
) -> dict:
    return {
        "run_id": run_id,
        "status": status,
        "stage": stage,
        "message": message,
        "grade_count": grade_count,
        "zles_count": zles_count,
        "material_count": material_count,
        "started_at": started_at.isoformat(),
        "finished_at": _now().isoformat() if status in {"success", "warning", "error"} else None,
        "details": details or {},
    }


def executar(grade_file: Path | None = None, zles_file: Path | None = None) -> dict:
    configure_logging()
    run_id = str(uuid.uuid4())
    started_at = _now()
    work_dir = Path(os.getenv("WORK_DIR", Path(__file__).resolve().parent / "runtime")).expanduser().resolve()
    input_dir = work_dir / "entrada"
    output_dir = work_dir / "saida"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    worker = WorkerClient()
    grade_count = zles_count = material_count = 0
    worker.job_status(_status_payload(run_id, "running", "starting", "Execucao iniciada.", started_at))

    try:
        grade_path = grade_file.resolve() if grade_file else baixar_grade(input_dir / "grade_atual.xlsx")
        grade_result = processar_grade(grade_path, started_at.date())
        grade_count = len(grade_result["cargas"])
        minimum = max(1, int(os.getenv("MIN_GRADE_DTS", "1")))
        maximum = max(minimum, int(os.getenv("MAX_GRADE_DTS", "3000")))
        if not minimum <= grade_count <= maximum:
            raise ValueError(f"Quantidade de DTs da grade fora do intervalo seguro: {grade_count} ({minimum}-{maximum}).")

        write_json(output_dir / "grade_normalizada.json", grade_result)
        worker.job_status(_status_payload(
            run_id, "running", "sync-grade", f"Grade validada com {grade_count} DTs.", started_at, grade_count
        ))
        grade_sync = worker.sync_grade(run_id, grade_result["cargas"], minimum)

        refs_by_dt = {row["dt"]: row["data_ref"] for row in grade_result["cargas"]}
        zles_path = zles_file.resolve() if zles_file else baixar_zles002(input_dir / "zles002_atual.xlsx", grade_result["dts"])
        zles_result = processar_zles002(zles_path, refs_by_dt)
        zles_count = len(zles_result["matched_dts"])
        material_count = len(zles_result["materiais"])
        if not zles_count:
            raise ValueError("ZLES002 sem correspondencia com a grade; sincronizacao cancelada.")

        write_json(output_dir / "zles002_normalizada.json", zles_result)
        worker.job_status(_status_payload(
            run_id,
            "running",
            "sync-zles002",
            f"ZLES002 validada para {zles_count} DTs e {material_count} materiais.",
            started_at,
            grade_count,
            zles_count,
            material_count,
        ))
        zles_sync = worker.sync_zles002(run_id, zles_result["cargas"], zles_result["materiais"])

        if env_bool("SYNC_DIRECT_SUPABASE"):
            SupabaseClient().upsert_cargas(grade_result["cargas"])

        message = f"Concluido: {grade_count} DTs na grade, {zles_count} DTs na ZLES002, {material_count} materiais."
        image_path = gerar_imagem_reporte(
            output_dir / f"reporte_automacao_{started_at:%Y%m%d_%H%M}.png",
            started_at.strftime("%d/%m/%Y"),
            grade_count,
            zles_count,
            material_count,
            "success",
            message,
        )
        image_warning = ""
        try:
            worker.report_image({
                "run_id": run_id,
                "data_ref": started_at.strftime("%d/%m/%Y"),
                "content_type": "image/png",
                "image_base64": imagem_base64(image_path),
            })
        except Exception as image_error:
            image_warning = f" Dados sincronizados, mas a imagem nao foi enviada: {image_error}"
            LOGGER.warning(image_warning)

        result = {
            "run_id": run_id,
            "status": "warning" if image_warning else "success",
            "grade": grade_sync,
            "zles002": zles_sync,
            "image": str(image_path),
        }
        worker.job_status(_status_payload(
            run_id,
            result["status"],
            "finished",
            message + image_warning,
            started_at,
            grade_count,
            zles_count,
            material_count,
            result,
        ))
        LOGGER.info(message)
        return result
    except Exception as exc:
        LOGGER.exception("Falha na execucao %s", run_id)
        try:
            worker.job_status(_status_payload(
                run_id,
                "error",
                "failed",
                str(exc),
                started_at,
                grade_count,
                zles_count,
                material_count,
            ))
        except Exception:
            LOGGER.exception("Tambem falhou o registro de status no Worker.")
        raise


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Automacao do Dashboard de Carga")
    parser.add_argument("--grade", type=Path, help="Arquivo local de grade para teste.")
    parser.add_argument("--zles002", type=Path, help="Arquivo local ZLES002 para teste.")
    return parser


if __name__ == "__main__":
    args = _parser().parse_args()
    try:
        executar(args.grade, args.zles002)
    except Exception as error:
        print(f"ERRO: {error}", file=sys.stderr)
        raise SystemExit(1)
