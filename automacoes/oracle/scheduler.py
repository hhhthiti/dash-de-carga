from __future__ import annotations

import logging
import os
from zoneinfo import ZoneInfo

from apscheduler.schedulers.blocking import BlockingScheduler

from common import configure_logging
from main import executar


LOGGER = logging.getLogger("oracle.scheduler")


def executar_protegido() -> None:
    try:
        executar()
    except Exception:
        LOGGER.exception("Execucao agendada falhou; o scheduler continuara ativo.")


def main() -> None:
    configure_logging()
    timezone = ZoneInfo(os.getenv("TIME_ZONE", "America/Sao_Paulo"))
    scheduler = BlockingScheduler(timezone=timezone)
    scheduler.add_job(
        executar_protegido,
        trigger="interval",
        minutes=20,
        id="dashboard-carga",
        max_instances=1,
        coalesce=True,
        misfire_grace_time=600,
    )
    LOGGER.info("Scheduler iniciado: execucao a cada 20 minutos.")
    executar_protegido()
    scheduler.start()


if __name__ == "__main__":
    main()
