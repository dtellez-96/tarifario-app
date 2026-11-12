from dataclasses import asdict, dataclass
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

ZERO = Decimal("0")
VERSION = "fcl-1.0"
BASES = {"CONTENEDOR", "BL", "FIJO", "EMBARQUE", "FIJO EMBARQUE"}


class ValidationError(ValueError):
    pass


def normalize(value):
    return str(value or "").strip().upper()


def decimal_input(value, label):
    try:
        result = Decimal(str(value).strip())
    except (InvalidOperation, ValueError):
        raise ValidationError(f"{label}: ingrese un numero decimal valido.") from None
    if not result.is_finite() or result < 0 or result > Decimal("999999999999"):
        raise ValidationError(f"{label}: valor fuera de rango.")
    if result.as_tuple().exponent < -4:
        raise ValidationError(f"{label}: use como maximo cuatro decimales.")
    return result


def money(value):
    return str(Decimal(value).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


@dataclass(frozen=True)
class QuoteInput:
    origin: str
    destination: str
    quantity20: int
    quantity40st: int
    quantity40hc: int
    bl: int
    fob: Decimal
    quote_date: date

    @property
    def containers(self):
        return self.quantity20 + self.quantity40st + self.quantity40hc

    @classmethod
    def parse(cls, form):
        counts = {}
        for field in ("quantity20", "quantity40st", "quantity40hc", "bl"):
            val = decimal_input(form.get(field, ""), field)
            if val != val.to_integral_value() or val > 50:
                raise ValidationError(f"{field}: debe ser entero entre 0 y 50.")
            counts[field] = int(val)
        if counts["bl"] < 1:
            raise ValidationError("BL: debe estar entre 1 y 50.")
        try:
            when = date.fromisoformat(form.get("quote_date", ""))
        except ValueError:
            raise ValidationError("Fecha de cotizacion invalida.") from None
        result = cls(
            normalize(form.get("origin")), normalize(form.get("destination")),
            **counts, fob=decimal_input(form.get("fob", "0"), "FOB"), quote_date=when,
        )
        if not result.origin or not result.destination or result.containers == 0:
            raise ValidationError("Seleccione la ruta e ingrese al menos un contenedor.")
        return result

    def snapshot(self):
        return {key: str(val) if isinstance(val, (date, Decimal)) else val
                for key, val in asdict(self).items()}


def recalculate(candidate, params):
    total = sum(Decimal(candidate[key]) for key in
                ("freight", "documentation", "origin_cost", "filling_cost", "destination_cost"))
    count = sum(params[key] for key in ("quantity20", "quantity40st", "quantity40hc"))
    fob = Decimal(params["fob"])
    candidate.update(total=str(total), average=str(total / count),
                     impact=str(total / fob if fob else ZERO))


def rank_candidates(candidates, mode):
    best = {}
    for candidate in candidates:
        if candidate["agent"] not in best or Decimal(candidate["total"]) < Decimal(best[candidate["agent"]]["total"]):
            best[candidate["agent"]] = candidate
    ordered = list(best.values())
    if mode == "legacy":
        # Match the VBA exchange sort, including its non-stable tie behavior.
        for i in range(len(ordered) - 1):
            for j in range(i + 1, len(ordered)):
                if Decimal(ordered[j]["total"]) < Decimal(ordered[i]["total"]):
                    ordered[i], ordered[j] = ordered[j], ordered[i]
    else:
        ordered.sort(key=lambda c: (Decimal(c["total"]), c["source_row"]))
    return ordered[:5], len(ordered)


def calculate(params, rates, charges, mode="strict"):
    if mode not in {"strict", "legacy"}:
        raise ValidationError("Modo de calculo invalido.")
    candidates, excluded = [], []
    for rate in rates:
        if rate.origin.name != params.origin or rate.destination.name != params.destination:
            continue
        warnings, reasons = [], []
        quantities = (params.quantity20, params.quantity40st, params.quantity40hc)
        values = (rate.freight20, rate.freight40st, rate.freight40hc)
        if not rate.valid_from or not rate.valid_until:
            reasons.append("Vigencia incompleta")
        elif not rate.valid_from <= params.quote_date <= rate.valid_until:
            reasons.append("Fuera de vigencia para la fecha solicitada")
        if rate.currency != "USD":
            reasons.append("Moneda ausente o distinta de USD")
        if not rate.carrier:
            reasons.append("Naviera ausente")
        if any(q and (v is None or v <= 0) for q, v in zip(quantities, values)):
            reasons.append("Flete solicitado ausente, cero o negativo")
        if rate.documentation is None or rate.documentation < 0:
            reasons.append("Documentacion BL ausente o negativa")
        if mode == "strict" and reasons:
            excluded.append({"rate_id": rate.id, "source_row": rate.source_row_id, "reasons": reasons})
            continue
        warnings.extend(reasons)
        carrier = rate.carrier.name if rate.carrier else ""
        applicable = [c for c in charges if c.agent_id == rate.agent_id
                      and c.destination_id == rate.destination_id and c.status == "VIGENTE"
                      and c.carrier_scope in {"", "TODAS", carrier}]
        destination = ZERO
        breakdown = []
        charge_errors = []
        for charge in applicable:
            if mode == "strict" and charge.valid_until and charge.valid_until < params.quote_date:
                charge_errors.append(f"Cargo {charge.id} vencido")
                continue
            if charge.basis not in BASES or charge.amount is None or charge.amount < 0:
                message = f"Cargo {charge.id}: base o importe no calculable ({charge.basis})"
                if mode == "strict":
                    charge_errors.append(message)
                    continue
                warnings.append(message)
            multiplier = params.containers if charge.basis == "CONTENEDOR" else (
                params.bl if charge.basis == "BL" else (1 if charge.basis in BASES else 0))
            amount = (charge.amount or ZERO) * multiplier
            destination += amount
            breakdown.append({"id": charge.id, "concept": charge.concept, "basis": charge.basis,
                              "unit": str(charge.amount or ZERO), "multiplier": multiplier,
                              "subtotal": str(amount)})
        if charge_errors:
            excluded.append({"rate_id": rate.id, "source_row": rate.source_row_id, "reasons": charge_errors})
            continue
        if not applicable:
            warnings.append("Sin cargos destino coincidentes; no implica que sean gratuitos.")
        else:
            warnings.append("Cargos destino sin moneda en Excel; suma nominal pendiente de confirmar en USD.")
            if any(not c.valid_until for c in applicable):
                warnings.append("Existen cargos sin fecha de vencimiento.")
        if rate.agent.name == "DSV" and any(v == 450 for v in values):
            warnings.append("Alerta de Guia: tarifa DSV de 450 requiere revision comercial.")
        warnings.append("Gastos en origen y llenado iniciales en cero; requieren revision manual.")
        lines = [(v or ZERO) * q for v, q in zip(values, quantities)]
        candidate = {
            "rate_id": rate.id, "source_row": rate.source_row_id,
            "agent": rate.agent.name, "carrier": carrier, "via": rate.via,
            "freight_lines": [str(v) for v in lines], "freight": str(sum(lines)),
            "documentation": str((rate.documentation or ZERO) * params.bl),
            "origin_cost": "0", "filling_cost": "0", "destination_cost": str(destination),
            "transit": rate.transit, "free_days": rate.free_days,
            "valid_from": str(rate.valid_from) if rate.valid_from else None,
            "valid_until": str(rate.valid_until) if rate.valid_until else None,
            "currency": rate.currency, "source_status": rate.source_status,
            "warnings": warnings, "charges": breakdown,
        }
        recalculate(candidate, params.snapshot())
        candidates.append(candidate)
    options, agent_count = rank_candidates(candidates, mode)
    return {"version": VERSION, "mode": mode, "input": params.snapshot(),
            "candidates": candidates, "options": options, "agent_count": agent_count,
            "excluded": excluded,
            "notice": "Estimacion provisional, no autorizacion comercial. Compatibilidad VBA solo para auditoria."}
