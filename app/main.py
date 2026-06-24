"""FastAPI application — serves the chat UI and processes rent roll uploads."""

from __future__ import annotations

import base64
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

from app.analyzer import UnitMixSummary, analyze
from app.exporter import export_unit_mix
from app.parser import parse_rent_roll

app = FastAPI(title="Unit Mix Summary Tool")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

STATIC_DIR = Path(__file__).parent.parent / "static"
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = STATIC_DIR / "index.html"
    if not html_path.exists():
        raise HTTPException(status_code=404, detail="Frontend not found")
    return HTMLResponse(content=html_path.read_text(encoding="utf-8"))


@app.post("/api/analyze")
async def analyze_rent_roll(file: UploadFile = File(...)):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Please upload an Excel file (.xlsx or .xls)")

    content = await file.read()
    try:
        rent_roll = parse_rent_roll(content)
        summary = analyze(rent_roll)
    except Exception as exc:
        raise HTTPException(status_code=422, detail=str(exc))

    # Serialize for JSON
    rows = [
        {
            "sqft": r.sqft,
            "num_units": r.num_units,
            "num_occupied": r.num_occupied,
            "num_vacant": r.num_vacant,
            "pct": r.pct,
            "gpr": round(r.gpr, 2),
            "loss_to_lease": round(r.loss_to_lease, 2),
            "in_place_rent": round(r.in_place_rent, 2),
            "vacancy": round(r.vacancy, 2),
            "net_rental_income": round(r.net_rental_income, 2),
        }
        for r in summary.rows
    ]

    # Generate Excel and embed as base64 so the frontend can offer a download
    excel_bytes = export_unit_mix(summary)
    excel_b64 = base64.b64encode(excel_bytes).decode()

    return JSONResponse({
        "property_name": summary.property_name,
        "as_of": summary.as_of,
        "rows": rows,
        "totals": {
            "total_units": summary.total_units,
            "total_occupied": summary.total_occupied,
            "total_vacant": summary.total_vacant,
        },
        "monthly": {
            "gpr": round(summary.monthly_gpr, 2),
            "loss": round(summary.monthly_loss, 2),
            "in_place": round(summary.monthly_in_place, 2),
            "vacancy": round(summary.monthly_vacancy, 2),
            "net": round(summary.monthly_net, 2),
        },
        "annual": {
            "gpr": round(summary.annual_gpr, 2),
            "loss": round(summary.annual_loss, 2),
            "in_place": round(summary.annual_in_place, 2),
            "vacancy": round(summary.annual_vacancy, 2),
            "net": round(summary.annual_net, 2),
        },
        "excel_b64": excel_b64,
        "filename": file.filename.replace(".xlsx", "").replace(".xls", "") + "_unit_mix.xlsx",
    })
