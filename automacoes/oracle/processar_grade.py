from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from openpyxl.utils.datetime import from_excel

from common import format_br_datetime, format_data_ref, normalize_dt, normalize_text, read_tabular


HEADERS = {
    "dt": ["DT", "N TRANSPORTE", "NUMERO TRANSPORTE", "TRANSPORTE", "NR TRANSPORTE"],
    "agenda": ["AGENDA TRANSPORTADOR", "INICIO AGENDA TRANSPORTADOR", "DATA HORA AGENDA TRANSPORTADOR", "AGENDA"],
    "fim": ["FIM AGENDA TRANSPORTADOR", "FIM DA AGENDA TRANSPORTADOR", "DATA HORA FIM AGENDA TRANSPORTADOR", "FIM AGENDA", "AGENDA FIM"],
    "local": ["LOCAL", "LOCAL CARREGAMENTO", "LOCAL DE CARREGAMENTO", "CENTRO", "CENTRO CD", "CD", "PLANTA"],
    "transportadora": ["NOME TRANSPORTADORA", "TRANSPORTADORA", "NOME TRANSP", "TRANSP"],
    "peso": ["PESO", "PESO LIQUIDO"],
    "tipo": ["TIPO VEICULO", "TIPO DE VEICULO", "TIPO"],
    "doca": ["DOCA", "DOCA CARREGAMENTO", "DOCA DE CARREGAMENTO"],
}


def _pick(headers: list[str], names: list[str], *, reject_fim: bool = False) -> int:
    normalized = [normalize_text(value) for value in headers]
    wanted = [normalize_text(value) for value in names]
    for index, header in enumerate(normalized):
        if reject_fim and ("FIM" in header or "TERMINO" in header):
            continue
        if header in wanted:
            return index
    for index, header in enumerate(normalized):
        if reject_fim and ("FIM" in header or "TERMINO" in header):
            continue
        if any(name in header or header in name for name in wanted if header):
            return index
    return -1


def _header_map(row: list[Any]) -> dict[str, int]:
    values = [str(value or "").strip() for value in row]
    return {
        "dt": _pick(values, HEADERS["dt"]),
        "agenda": _pick(values, HEADERS["agenda"], reject_fim=True),
        "fim": _pick(values, HEADERS["fim"]),
        "local": _pick(values, HEADERS["local"]),
        "transportadora": _pick(values, HEADERS["transportadora"]),
        "peso": _pick(values, HEADERS["peso"]),
        "tipo": _pick(values, HEADERS["tipo"]),
        "doca": _pick(values, HEADERS["doca"]),
    }


def _score(mapping: dict[str, int]) -> int:
    return (4 if mapping["dt"] >= 0 else 0) + (4 if mapping["agenda"] >= 0 else 0) + sum(
        1 for key in ("fim", "local", "transportadora", "doca") if mapping[key] >= 0
    )


def _value(row: list[Any], index: int) -> Any:
    return row[index] if 0 <= index < len(row) else ""


def _parse_datetime(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, (int, float)) and value > 1:
        try:
            return from_excel(value)
        except (ValueError, OverflowError):
            return None
    raw = str(value or "").strip()
    if not raw:
        return None
    raw = raw.replace("T", " ").replace(",", " ")
    for fmt in (
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d-%m-%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d/%m/%Y",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(re.sub(r"\s+", " ", raw), fmt)
        except ValueError:
            continue
    return None


def _local_permitido(value: Any) -> bool:
    numbers = re.findall(r"\d+", str(value or ""))
    last = numbers[-1].lstrip("0") if numbers else ""
    return last.endswith("1110") or last.endswith("1111")


def _doca_fab_mog_sem_ifnt(value: Any) -> bool:
    normalized = normalize_text(value).replace("-", "_").replace(" ", "_")
    return bool(re.search(r"(?:^|_)DOCA_\d+_FAB_MOG(?:_|$)", normalized)) and "IFNT" not in normalized


def processar_grade(path: Path, operation_date: date | None = None) -> dict[str, Any]:
    rows = read_tabular(path)
    if not rows:
        raise ValueError("Grade vazia.")

    candidates = [(index, _header_map(row)) for index, row in enumerate(rows[:50])]
    header_index, mapping = max(candidates, key=lambda item: _score(item[1]))
    if mapping["dt"] < 0 or mapping["agenda"] < 0:
        raise ValueError("Colunas DT e AGENDA TRANSPORTADOR nao foram encontradas.")

    base = operation_date or date.today()
    allowed_dates = {base, base + timedelta(days=1)}
    diagnostics: list[dict[str, Any]] = []
    by_dt: dict[str, dict[str, Any]] = {}

    for line_number, row in enumerate(rows[header_index + 1:], start=header_index + 2):
        dt = normalize_dt(_value(row, mapping["dt"]))
        if not dt:
            continue
        agenda = _parse_datetime(_value(row, mapping["agenda"]))
        fim = _parse_datetime(_value(row, mapping["fim"])) or agenda
        local = str(_value(row, mapping["local"]) or "").strip()
        doca = str(_value(row, mapping["doca"]) or "").strip()
        reason = ""
        if not _local_permitido(local):
            reason = "LOCAL fora de 1110/1111"
        elif _doca_fab_mog_sem_ifnt(doca):
            reason = "DOCA FAB_MOG sem IFNT"
        elif agenda is None:
            reason = "AGENDA invalida"
        elif fim is None or fim.date() not in allowed_dates:
            reason = "FIM AGENDA fora de hoje/amanha"

        if reason:
            diagnostics.append({"linha": line_number, "dt": dt, "motivo": reason})
            continue

        assert agenda is not None
        assert fim is not None
        data_ref = format_data_ref(fim)
        carga = {
            "dt": dt,
            "data_ref": data_ref,
            "transportadora": str(_value(row, mapping["transportadora"]) or "").strip(),
            "grade_carregamento": format_br_datetime(agenda),
            "fim_carregamento": format_br_datetime(fim),
            "agenda": format_br_datetime(agenda),
            "local_cd": local,
            "tipo": str(_value(row, mapping["tipo"]) or "").strip(),
            "toneladas": str(_value(row, mapping["peso"]) or "").strip(),
            "dia_ref": "HOJE" if fim.date() == base else "AMANHA",
            "doca_null": not doca or normalize_text(doca) in {"NULL", "N/A", "NAO INFORMADO"},
        }
        previous = by_dt.get(dt)
        if previous and previous["data_ref"] == data_ref:
            for field, value in carga.items():
                if not previous.get(field) and value:
                    previous[field] = value
        else:
            by_dt[dt] = carga

    cargas = list(by_dt.values())
    if not cargas:
        raise ValueError(f"Nenhuma DT valida na grade. Linhas rejeitadas: {len(diagnostics)}.")
    return {
        "cargas": cargas,
        "dts": sorted({row["dt"] for row in cargas}),
        "refs": sorted({row["data_ref"] for row in cargas}),
        "diagnostics": diagnostics,
        "source_file": str(path),
    }
