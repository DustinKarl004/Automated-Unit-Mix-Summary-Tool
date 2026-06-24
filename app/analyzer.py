"""
Unit mix analysis — replicates the Excel formulas in the analysis template (E1:N28).

Formula logic (from the template):
  # of Units       = COUNTIFS(sqft = sf)
  # of Vacant      = COUNTIFS(sqft=sf, name="Vacant") + COUNTIFS(sqft=sf, name="Down")
  # of Occupied    = total - vacant
  %                = units / total_units
  GPR              = AVERAGEIFS(market_rent, sqft=sf)          [avg over ALL units]
  In Place Rent    = AVERAGEIFS(actual_rent, sqft=sf,
                       name<>"Vacant", name<>"Down")            [avg over occupied only]
  Loss to Lease    = In Place - GPR
  Net Rental Income= AVERAGEIFS(actual_rent, sqft=sf)          [avg over ALL units incl. 0]
  Vacancy          = Net - In Place

Monthly totals = SUMPRODUCT(per_unit_value × unit_count)
Annual totals  = Monthly × 12
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from app.parser import RentRoll, RentRollUnit


def _is_non_revenue(unit: RentRollUnit) -> bool:
    name_upper = unit.name.upper()
    return name_upper in ("VACANT", "DOWN")


@dataclass
class UnitTypeRow:
    sqft: int
    num_units: int
    num_occupied: int
    num_vacant: int
    pct: float
    gpr: float
    loss_to_lease: float
    in_place_rent: float
    vacancy: float
    net_rental_income: float


@dataclass
class UnitMixSummary:
    property_name: str
    as_of: str
    rows: list[UnitTypeRow]
    total_units: int
    total_occupied: int
    total_vacant: int
    monthly_gpr: float
    monthly_loss: float
    monthly_in_place: float
    monthly_vacancy: float
    monthly_net: float
    annual_gpr: float
    annual_loss: float
    annual_in_place: float
    annual_vacancy: float
    annual_net: float


def analyze(rent_roll: RentRoll) -> UnitMixSummary:
    units = rent_roll.units
    if not units:
        raise ValueError("No unit data found in the uploaded file.")

    total_units = len(units)

    # Group by sqft, preserving insertion order (first-seen ordering)
    sqft_order: list[int] = []
    groups: dict[int, list[RentRollUnit]] = {}
    for u in units:
        if u.sqft not in groups:
            groups[u.sqft] = []
            sqft_order.append(u.sqft)
        groups[u.sqft].append(u)

    rows: list[UnitTypeRow] = []

    for sf in sqft_order:
        group = groups[sf]
        num_units = len(group)

        occupied = [u for u in group if not _is_non_revenue(u)]
        non_rev = [u for u in group if _is_non_revenue(u)]

        num_vacant = len(non_rev)
        num_occupied = num_units - num_vacant
        pct = num_units / total_units if total_units else 0

        # GPR = average market rent over ALL units in type
        gpr = sum(u.market_rent for u in group) / num_units if num_units else 0

        # In Place Rent = average actual rent for NON-vacant/non-down units only
        in_place_rent = (
            sum(u.actual_rent for u in occupied) / len(occupied)
            if occupied else 0
        )

        # Loss to Lease = In Place - GPR
        loss_to_lease = in_place_rent - gpr

        # Net Rental Income = average actual rent over ALL units (vacant = 0)
        net_rental_income = sum(u.actual_rent for u in group) / num_units if num_units else 0

        # Vacancy = Net - In Place
        vacancy = net_rental_income - in_place_rent

        rows.append(UnitTypeRow(
            sqft=sf,
            num_units=num_units,
            num_occupied=num_occupied,
            num_vacant=num_vacant,
            pct=pct,
            gpr=gpr,
            loss_to_lease=loss_to_lease,
            in_place_rent=in_place_rent,
            vacancy=vacancy,
            net_rental_income=net_rental_income,
        ))

    total_occupied = sum(r.num_occupied for r in rows)
    total_vacant = sum(r.num_vacant for r in rows)

    # Monthly totals = SUMPRODUCT(per_unit_value × unit_count)
    monthly_gpr = sum(r.gpr * r.num_units for r in rows)
    monthly_in_place = sum(r.in_place_rent * r.num_units for r in rows)
    monthly_loss = sum(r.loss_to_lease * r.num_units for r in rows)
    monthly_net = sum(r.net_rental_income * r.num_units for r in rows)
    monthly_vacancy = monthly_net - monthly_in_place

    return UnitMixSummary(
        property_name=rent_roll.property_name,
        as_of=rent_roll.as_of,
        rows=rows,
        total_units=total_units,
        total_occupied=total_occupied,
        total_vacant=total_vacant,
        monthly_gpr=monthly_gpr,
        monthly_loss=monthly_loss,
        monthly_in_place=monthly_in_place,
        monthly_vacancy=monthly_vacancy,
        monthly_net=monthly_net,
        annual_gpr=monthly_gpr * 12,
        annual_loss=monthly_loss * 12,
        annual_in_place=monthly_in_place * 12,
        annual_vacancy=monthly_vacancy * 12,
        annual_net=monthly_net * 12,
    )
