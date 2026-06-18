from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Any

from common import normalize_dt, normalize_text, parse_number, read_tabular


HEADER_NAMES = {
    "dt": ["N TRANSPORTE", "NUMERO TRANSPORTE", "TRANSPORTE", "DT"],
    "descricao": ["DESCRICAO DE DOCUMENTO", "DESCRICAO DOCUMENTO", "DESCRI"],
    "peso": ["PESO LIQUIDO", "PESO"],
    "remessa": ["NR REMESSA RECEBIMENTO", "NR REMESSA", "REMESSA"],
    "material": ["MATERIAL"],
    "quantidade": ["QTDE REMESSA", "QUANTIDADE REMESSA", "QTDE", "QUANTIDADE"],
    "hora": ["HORA CHEGADA", "HORA"],
    "sap": ["NUMERO SAP", "NR SAP", "N SAP", "SAP", "PORTARIA"],
    "centro": ["CENTRO", "CTR"],
    "info_agenda": ["INF AGENDA ENTREGA", "AGENDA ENTREGA"],
    "cliente": ["NOME CLIENTE FORNECEDOR", "CLIENTE FORNECEDOR", "NOME CLIENTE"],
    "local": ["LOCAL"],
    "status": ["STATUS"],
    "faturamento": ["FATURAMENTO"],
}

STATUS_MAP = {
    "CARREGANDO": "CARREGANDO",
    "EXPEDIDO": "EXPEDIDO",
    "PATIO": "PATIO",
    "EM PATIO": "PATIO",
    "AGUARDANDO CHEGADA": "AG CHEGADA",
    "AG CHEGADA": "AG CHEGADA",
    "NO SHOW": "NO SHOW",
    "SEPARANDO": "SEPARANDO",
    "EM FATURAMENTO": "EM FATURAMENTO",
    "VEICULO RECUSADO": "VEICULO RECUSADO",
    "MOTORISTA FOI EMBORA": "FOI EMBORA",
    "FOI EMBORA": "FOI EMBORA",
    "DT EXCLUIDA": "DT EXCLUIDA",
    "BAIXA N OCUPACAO": "AG CHEGADA",
    "BAIXA NOCUPACAO": "AG CHEGADA",
}

FATURAMENTO_MAP = {
    "CUSTO DE FRETE": "EM FATURAMENTO",
    "PROBLEMA DE DT": "EM FATURAMENTO",
    "PROBELMA DE DT": "EM FATURAMENTO",
    "PROBLEMA JSL": "EM FATURAMENTO",
    "PROBELMA JSL": "EM FATURAMENTO",
}


def _find_header(headers: list[Any], names: list[str]) -> int:
    normalized = [normalize_text(value).replace(".", " ") for value in headers]
    wanted = [normalize_text(value) for value in names]
    for index, header in enumerate(normalized):
        if header in wanted:
            return index
    for index, header in enumerate(normalized):
        if any(name in header or header in name for name in wanted if header):
            return index
    return -1


def _header_map(row: list[Any]) -> dict[str, int]:
    return {key: _find_header(row, names) for key, names in HEADER_NAMES.items()}


def _score(mapping: dict[str, int]) -> int:
    return (5 if mapping["dt"] >= 0 else 0) + sum(
        1 for key in ("descricao", "peso", "remessa", "material", "quantidade", "centro") if mapping[key] >= 0
    )


def _value(row: list[Any], index: int) -> Any:
    return row[index] if 0 <= index < len(row) else ""


def _normalize_hour(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    if ":" in raw:
        parts = raw.replace("h", ":").split(":")
        try:
            return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
        except (ValueError, IndexError):
            return ""
    try:
        number = float(raw.replace(",", "."))
    except ValueError:
        return ""
    if 0 < number < 1:
        total = round(number * 24 * 60)
        return f"{(total // 60) % 24:02d}:{total % 60:02d}"
    digits = "".join(character for character in raw if character.isdigit()).zfill(4)
    return f"{digits[-4:-2]}:{digits[-2:]}" if len(digits) >= 4 else ""


def _paletizacao(value: Any) -> str:
    normalized = normalize_text(value)
    if "TORDE" in normalized or "TORDESILHAS" in normalized:
        return "TORDESILHAS"
    if "PLT" in normalized or "PALET" in normalized:
        return "PALETIZADA"
    return "ESTIVADA"


def _tipo_operacao(description: str, centro: str, local: str) -> str:
    text = normalize_text(description)
    local_text = normalize_text(local)
    if any(term in text for term in ("TRANSFER", "FILIAL", "ABAST", "TNF")):
        return "TRANSFERENCIA"
    if any(term in text for term in ("PREFAT", "PRE FAT")):
        return "PREFATURA"
    if centro == "1110" or "MOGI" in local_text:
        return "VENDA MOGI"
    if centro in {"", "1111"} or "ARUJA" in local_text:
        return "VENDA ARUJA"
    return "VENDA NORMAL"


def _status(status_raw: Any, billing_raw: Any) -> str:
    status = STATUS_MAP.get(normalize_text(status_raw))
    if status:
        return status
    return FATURAMENTO_MAP.get(normalize_text(billing_raw), "")


def _number_text(value: float) -> str:
    return str(int(value)) if value.is_integer() else f"{value:.2f}".rstrip("0").rstrip(".")


def processar_zles002(path: Path, refs_by_dt: dict[str, str]) -> dict[str, Any]:
    rows = read_tabular(path)
    if not rows:
        raise ValueError("ZLES002 vazia.")
    candidates = [(index, _header_map(row)) for index, row in enumerate(rows[:50])]
    header_index, mapping = max(candidates, key=lambda item: _score(item[1]))
    if mapping["dt"] < 0:
        raise ValueError("Coluna de transporte/DT nao encontrada na ZLES002.")

    cargas: dict[str, dict[str, Any]] = {}
    material_totals: dict[tuple[str, str], float] = defaultdict(float)
    material_meta: dict[tuple[str, str], dict[str, str]] = {}
    processed_lines = 0

    for row in rows[header_index + 1:]:
        dt = normalize_dt(_value(row, mapping["dt"]))
        data_ref = refs_by_dt.get(dt)
        if not dt or not data_ref:
            continue

        description = str(_value(row, mapping["descricao"]) or "").strip()
        centro = str(_value(row, mapping["centro"]) or "").strip().replace(".0", "")
        local = str(_value(row, mapping["local"]) or "").strip()
        info_agenda = str(_value(row, mapping["info_agenda"]) or "").strip()
        current = cargas.setdefault(dt, {
            "dt": dt,
            "data_ref": data_ref,
            "descricao_documento": description,
            "tipo_operacao": _tipo_operacao(description, centro, local),
            "centro": centro,
            "nome_cliente_fornecedor": str(_value(row, mapping["cliente"]) or "").strip(),
            "peso_liquido": 0.0,
            "hora_chegada": _normalize_hour(_value(row, mapping["hora"])),
            "n_portaria": normalize_dt(_value(row, mapping["sap"])),
            "paletizacao": _paletizacao(info_agenda),
            "status": _status(_value(row, mapping["status"]), _value(row, mapping["faturamento"])),
        })
        current["peso_liquido"] += parse_number(_value(row, mapping["peso"]))
        for field, value in (
            ("descricao_documento", description),
            ("centro", centro),
            ("nome_cliente_fornecedor", str(_value(row, mapping["cliente"]) or "").strip()),
            ("hora_chegada", _normalize_hour(_value(row, mapping["hora"]))),
            ("n_portaria", normalize_dt(_value(row, mapping["sap"]))),
            ("status", _status(_value(row, mapping["status"]), _value(row, mapping["faturamento"]))),
        ):
            if value and not current.get(field):
                current[field] = value
        if _paletizacao(info_agenda) == "TORDESILHAS":
            current["paletizacao"] = "TORDESILHAS"
        elif _paletizacao(info_agenda) == "PALETIZADA" and current["paletizacao"] == "ESTIVADA":
            current["paletizacao"] = "PALETIZADA"

        material = normalize_dt(_value(row, mapping["material"]))
        if material:
            key = (dt, material)
            material_totals[key] += parse_number(_value(row, mapping["quantidade"]))
            material_meta[key] = {"data_ref": data_ref, "paletizacao": current["paletizacao"]}
        processed_lines += 1

    carga_rows = []
    for row in cargas.values():
        row["peso_liquido"] = _number_text(float(row["peso_liquido"]))
        carga_rows.append({key: value for key, value in row.items() if value not in ("", None)})

    materials_by_dt: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for (dt, material), quantity in material_totals.items():
        meta = material_meta[(dt, material)]
        materials_by_dt[dt].append({
            "dt": dt,
            "data_ref": meta["data_ref"],
            "material": material,
            "quantidade": _number_text(quantity),
            "observacao": "",
            "paletizacao": meta["paletizacao"],
        })

    material_rows = []
    for dt, rows_for_dt in materials_by_dt.items():
        for order, row in enumerate(sorted(rows_for_dt, key=lambda item: item["material"])):
            row["ordem"] = order
            material_rows.append(row)

    if not carga_rows:
        raise ValueError("A ZLES002 nao possui DTs que correspondam a grade processada.")
    return {
        "cargas": carga_rows,
        "materiais": material_rows,
        "processed_lines": processed_lines,
        "matched_dts": sorted(cargas),
        "source_file": str(path),
    }
