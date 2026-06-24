from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

import openpyxl


@dataclass
class RentRollUnit:
    unit: str
    unit_type: str
    sqft: int
    resident_id: str
    name: str
    market_rent: float
    actual_rent: float


@dataclass
class RentRoll:
    property_name: str
    as_of: str
    units: list[RentRollUnit] = field(default_factory=list)


def _is_unit_row(row_values: tuple) -> bool:
    val = row_values[0]
    if val is None:
        return False
    return bool(re.match(r"^\d+", str(val).strip()))


def _safe_float(val) -> float:
    if val is None:
        return 0.0
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0.0


def _safe_str(val) -> str:
    if val is None:
        return ""
    return str(val).strip()


def parse_rent_roll(file_bytes: bytes) -> RentRoll:
    from io import BytesIO

    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active

    all_rows = list(ws.iter_rows(values_only=True))

    property_name = ""
    as_of = ""
    for row in all_rows[:10]:
        val = _safe_str(row[0]) if row[0] is not None else ""
        if val and not property_name and val.lower() not in ("rent roll", ""):
            if property_name == "" and val.lower() != "rent roll":
                property_name = val
        if val.lower().startswith("as of"):
            as_of = val.split("=")[-1].strip() if "=" in val else val

    units: list[RentRollUnit] = []
    for row in all_rows:
        if not _is_unit_row(row):
            continue

        padded = list(row) + [None] * 13
        unit_num = _safe_str(padded[0])
        unit_type = _safe_str(padded[1])
        sqft_raw = padded[2]
        resident_id = _safe_str(padded[3])
        name = _safe_str(padded[4])
        market_rent = _safe_float(padded[5])
        actual_rent = _safe_float(padded[6])

        try:
            sqft = int(float(str(sqft_raw))) if sqft_raw is not None else 0
        except (TypeError, ValueError):
            sqft = 0

        if sqft == 0:
            continue

        units.append(RentRollUnit(
            unit=unit_num,
            unit_type=unit_type,
            sqft=sqft,
            resident_id=resident_id,
            name=name,
            market_rent=market_rent,
            actual_rent=actual_rent,
        ))

    return RentRoll(property_name=property_name, as_of=as_of, units=units)
