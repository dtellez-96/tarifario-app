from decimal import Decimal
from io import BytesIO
from itertools import product

import openpyxl
import pytest

from app.extensions import db
from app.models import DestinationCharge, ImportBatch, Quote, Rate, SourceRow
from app.services.imports import import_workbook
from app.services.pricing import QuoteInput, ValidationError, calculate
from conftest import WORKBOOK


def legacy_oracle(raw_rates, raw_charges, origin, destination, quantities, bl):
    """Independent literal transcription of Module1 for this workbook's valid inputs."""
    best = {}
    number = lambda value: Decimal(str(value)) if isinstance(value, (int, float)) else Decimal(0)
    text = lambda value: str(value or "").strip().upper()
    for row in raw_rates:
        if text(row[4]) != origin or text(row[5]) != destination:
            continue
        agent, carrier = text(row[1]), text(row[7])
        freight = sum(number(row[8 + i]) * quantities[i] for i in range(3))
        total = freight + number(row[11]) * bl
        for charge in raw_charges:
            if (text(charge[0]) == agent and text(charge[6]) == destination
                    and text(charge[8]) == "VIGENTE" and text(charge[7]) in ("", "TODAS", carrier)):
                basis = text(charge[4])
                multiplier = sum(quantities) if basis == "CONTENEDOR" else (
                    bl if basis == "BL" else (1 if basis in ("FIJO", "EMBARQUE", "FIJO EMBARQUE") else 0))
                total += number(charge[5]) * multiplier
        result = (agent, carrier, total)
        if agent not in best or total < best[agent][2]:
            best[agent] = result
    ordered = list(best.values())
    for i in range(len(ordered) - 1):
        for j in range(i + 1, len(ordered)):
            if ordered[j][2] < ordered[i][2]:
                ordered[i], ordered[j] = ordered[j], ordered[i]
    return ordered[:5]


def test_import_counts_and_idempotence(application):
    assert db.session.scalar(db.select(db.func.count(Rate.id))) == 360
    assert db.session.scalar(db.select(db.func.count(DestinationCharge.id))) == 100
    assert db.session.scalar(db.select(db.func.count(SourceRow.id))) == 996
    summary = db.session.get(ImportBatch, 1).summary
    assert summary["quarantined_rates"] == 126
    assert summary["legacy_history"] == 1
    with pytest.raises(ValidationError, match="ya fue importado"):
        import_workbook(WORKBOOK.read_bytes(), WORKBOOK.name)


def test_all_routes_against_independent_vba_transcription(application):
    book = openpyxl.load_workbook(WORKBOOK, data_only=True)
    raw_rates = list(book["TARIFARIOS_AGENTES"].values)[1:]
    raw_charges = list(book["Cargos_Destino"].values)[4:]
    origins = sorted({str(r[4]).strip().upper() for r in raw_rates if r[4]})
    destinations = sorted({str(r[5]).strip().upper() for r in raw_rates if r[5]})
    rates = db.session.scalars(db.select(Rate).order_by(Rate.source_row_id)).all()
    charges = db.session.scalars(db.select(DestinationCharge).order_by(DestinationCharge.source_row_id)).all()
    cases = 0
    for origin, destination, quantities, bl in product(
        origins, destinations, [(1, 0, 0), (0, 1, 0), (0, 0, 2), (1, 2, 3), (50, 50, 50)], [1, 3],
    ):
        params = QuoteInput.parse(dict(origin=origin, destination=destination,
            quantity20=quantities[0], quantity40st=quantities[1], quantity40hc=quantities[2],
            bl=bl, fob="100000", quote_date="2026-11-10"))
        actual = calculate(params, rates, charges, "legacy")
        observed = [(o["agent"], o["carrier"], Decimal(o["total"])) for o in actual["options"]]
        assert observed == legacy_oracle(raw_rates, raw_charges, origin, destination, quantities, bl)
        cases += 1
    assert cases == 900
    book.close()


def test_bad_structure_and_atomic_rollback(application):
    book = openpyxl.load_workbook(WORKBOOK)
    book["Cargos_Destino"]["A4"] = "Broken"
    output = BytesIO()
    book.save(output)
    with pytest.raises(ValidationError, match="Encabezados"):
        import_workbook(output.getvalue(), "invalid.xlsx")
    db.session.rollback()
    assert db.session.scalar(db.select(db.func.count(ImportBatch.id))) == 1
    assert db.session.scalar(db.select(db.func.count(Rate.id))) == 360


def test_rejects_non_zip(application):
    with pytest.raises(ValidationError, match="Archivo Excel"):
        import_workbook(b"not excel", "bad.xlsm")


def test_quote_end_to_end_and_snapshot(client, application):
    form = dict(origin="SHANGHAI", destination="CARTAGENA", quantity20="0",
                quantity40st="0", quantity40hc="2", bl="1", fob="100000",
                quote_date="2026-11-10", batch_id="1", mode="strict")
    response = client.post("/quotes", data=form)
    assert response.status_code == 303
    location = response.headers["Location"]
    detail = client.get(location)
    assert detail.status_code == 200 and b"COSTO TOTAL" in detail.data
    quote_id = location.rsplit("/", 1)[1]
    quote = db.session.get(Quote, quote_id)
    assert quote.snapshot["options"]
    option = quote.snapshot["options"][0]
    adjusted = client.post(location + "/adjust", data={
        "rate_id": option["rate_id"], "origin_cost": "25.50", "filling_cost": "10",
        "reason": "Ajuste documentado de prueba",
    })
    assert adjusted.status_code == 303
    db.session.refresh(quote)
    candidate = next(c for c in quote.snapshot["candidates"] if c["rate_id"] == option["rate_id"])
    assert Decimal(candidate["total"]) == Decimal(option["total"]) + Decimal("35.50")
    assert client.post(location + "/save").status_code == 303
    assert client.post(location + "/save").status_code == 303
    assert client.post(location + "/adjust", data={}).status_code == 409
    assert quote_id.encode() in client.get("/history").data


@pytest.mark.parametrize("path", ["/", "/history", "/imports", "/imports/1/issues",
                                     "/catalog/rates", "/catalog/charges", "/catalog/ports",
                                     "/static/app.css", "/static/app.js"])
def test_pages(client, path):
    assert client.get(path).status_code == 200


def test_validation_and_local_access(client):
    assert client.post("/quotes", data={}).status_code == 422
    assert client.get("/", environ_overrides={"REMOTE_ADDR": "192.0.2.1"}).status_code == 403
    assert client.post("/imports", data={}).status_code == 422
    assert "img-src 'self' data:" in client.get("/").headers["Content-Security-Policy"]


def test_csrf_required(application):
    application.config["WTF_CSRF_ENABLED"] = True
    try:
        client = application.test_client()
        assert client.post("/quotes", data={}).status_code == 400
        assert client.get("/").status_code == 200
    finally:
        application.config["WTF_CSRF_ENABLED"] = False


def test_legacy_not_saved_and_empty_route(client):
    base = dict(origin="NO EXISTE", destination="CARTAGENA", quantity20="1",
                quantity40st="0", quantity40hc="0", bl="1", fob="0",
                quote_date="2026-11-10", batch_id="1")
    for mode in ("strict", "legacy"):
        response = client.post("/quotes", data={**base, "mode": mode})
        location = response.headers["Location"]
        assert b"SIN OPCION" in client.get(location).data
        assert client.post(location + "/save").status_code == 422
