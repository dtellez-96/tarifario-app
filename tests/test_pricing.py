from datetime import date
from decimal import Decimal
from types import SimpleNamespace as NS

import pytest

from app.services.pricing import QuoteInput, ValidationError, calculate, rank_candidates


def params(**kwargs):
    values = dict(origin="SHANGHAI", destination="CARTAGENA", quantity20="1",
                  quantity40st="2", quantity40hc="3", bl="2", fob="100000",
                  quote_date="2026-11-10")
    values.update(kwargs)
    return QuoteInput.parse(values)


def rate(**kwargs):
    values = dict(id=1, source_row_id=1, agent_id=1, destination_id=1,
                  agent=NS(name="TEST"), carrier=NS(name="ONE"),
                  origin=NS(name="SHANGHAI"), destination=NS(name="CARTAGENA"),
                  valid_from=date(2026, 11, 9), valid_until=date(2026, 11, 16),
                  currency="USD", freight20=Decimal(100), freight40st=Decimal(200),
                  freight40hc=Decimal(300), documentation=Decimal(50), via="",
                  transit="30", free_days="14", source_status="VIGENTE")
    values.update(kwargs)
    return NS(**values)


def charge(**kwargs):
    values = dict(id=1, agent_id=1, destination_id=1, carrier_scope="TODAS",
                  concept="THC", basis="CONTENEDOR", amount=Decimal(10),
                  status="VIGENTE", valid_until=None)
    values.update(kwargs)
    return NS(**values)


def test_container_bl_and_fixed_charges():
    result = calculate(params(), [rate()], [
        charge(), charge(id=2, basis="BL", amount=Decimal(20)),
        charge(id=3, basis="FIJO", amount=Decimal(30)),
        charge(id=4, basis="EMBARQUE", amount=Decimal(5)),
        charge(id=5, basis="FIJO EMBARQUE", amount=Decimal(7)),
    ])
    option = result["options"][0]
    assert Decimal(option["freight"]) == 1400
    assert Decimal(option["documentation"]) == 100
    assert Decimal(option["destination_cost"]) == 142
    assert Decimal(option["total"]) == 1642
    assert Decimal(option["impact"]) == Decimal("0.01642")


@pytest.mark.parametrize("changes", [
    {"quantity20": "-1"}, {"quantity40st": "1.5"}, {"quantity40hc": "51"},
    {"bl": "0"}, {"fob": "NaN"}, {"fob": "Infinity"}, {"fob": "-1"},
    {"quote_date": "bad"}, {"origin": ""}, {"destination": ""},
    {"quantity20": "0", "quantity40st": "0", "quantity40hc": "0"},
    {"fob": "1.12345"}, {"quantity20": "abc"},
])
def test_invalid_inputs_are_explicit(changes):
    with pytest.raises(ValidationError):
        params(**changes)


@pytest.mark.parametrize("changes", [
    {"valid_until": date(2026, 11, 9)}, {"valid_from": date(2026, 11, 11)},
    {"valid_from": None}, {"valid_until": None}, {"currency": None}, {"currency": "EUR"},
    {"documentation": None}, {"freight20": None}, {"freight40hc": Decimal(0)},
    {"carrier": None},
])
def test_strict_excludes_and_legacy_warns(changes):
    strict = calculate(params(), [rate(**changes)], [])
    assert strict["excluded"] and not strict["options"]
    legacy = calculate(params(), [rate(**changes)], [], "legacy")
    assert legacy["options"][0]["warnings"]


def test_validity_boundaries_and_unrequested_equipment():
    for when in ("2026-11-09", "2026-11-16"):
        result = calculate(params(quote_date=when, quantity20="0"), [rate(freight20=None)], [])
        assert result["options"]


@pytest.mark.parametrize("scope,expected", [("TODAS", 60), ("", 60), ("ONE", 60), ("COSCO", 0)])
def test_carrier_scope(scope, expected):
    result = calculate(params(), [rate()], [charge(carrier_scope=scope)])
    assert Decimal(result["options"][0]["destination_cost"]) == expected


@pytest.mark.parametrize("changes", [
    {"status": "VENCIDO"}, {"agent_id": 2}, {"destination_id": 2},
])
def test_nonmatching_charges(changes):
    result = calculate(params(), [rate()], [charge(**changes)])
    assert Decimal(result["options"][0]["destination_cost"]) == 0


def test_percentage_is_not_invented_and_expired_charge_blocks_strict():
    for c in [charge(basis="PORCENTAJE_FLETE"), charge(basis="DESCONOCIDA"),
              charge(amount=None), charge(valid_until=date(2026, 11, 1))]:
        assert not calculate(params(), [rate()], [c])["options"]
        assert calculate(params(), [rate()], [c], "legacy")["options"]
    legacy = calculate(params(), [rate()], [charge(basis="PORCENTAJE")], "legacy")
    assert Decimal(legacy["options"][0]["destination_cost"]) == 0


def test_best_per_agent_top_five_and_ties():
    candidates = [
        {"agent": f"A{i}", "total": str(8 - i), "source_row": i, "rate_id": i}
        for i in range(7)
    ]
    candidates += [{"agent": "A0", "total": "1", "source_row": 10, "rate_id": 10}]
    best, count = rank_candidates(candidates, "strict")
    assert len(best) == 5 and count == 7
    assert best[0]["rate_id"] == 10
    tie = [
        {"agent": "A", "total": "2", "source_row": 1},
        {"agent": "B", "total": "2", "source_row": 2},
        {"agent": "C", "total": "1", "source_row": 3},
    ]
    assert [c["agent"] for c in rank_candidates(tie, "legacy")[0]] == ["C", "B", "A"]


def test_zero_fob_and_normalization():
    result = calculate(params(origin=" shanghai ", destination="cartagena", fob="0"), [rate()], [])
    assert result["options"][0]["impact"] == "0"
