import hashlib
import json
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
from zipfile import BadZipFile, ZipFile
from xml.etree.ElementTree import ParseError

import openpyxl
from openpyxl.utils.cell import range_boundaries

from ..extensions import db
from ..models import Agent, Carrier, DestinationCharge, ImportBatch, Port, Rate, SourceRow
from .pricing import ValidationError, normalize

TABLES = {
    "TARIFARIOS_AGENTES": "TARIFARIOS_AGENTES",
    "Cargos_Destino": "tblCargosDestino",
    "Tarifas": "tblTarifas",
    "Historico": "tblHistorico",
}
REQUIRED = {
    "TARIFARIOS_AGENTES": {
        "AGENTE", "POL", "POD", "NAVIERA", "FLETE_20ST", "FLETE_40ST",
        "FLETE_40HC", "DOCUMENTACION_BL", "Fecha_Inicio", "Fecha_Vencimiento",
        "Moneda", "Source.Name", "TT", "DIAS_LIBRES", "VIA", "Estado Vigencia",
    },
    "Cargos_Destino": {
        "Agente", "Puerto_Destino", "Naviera_Aplica", "Concepto", "Base",
        "Importe", "Estado", "Vigencia_Hasta",
    },
}


def amount(value, issues, field):
    if value is None or str(value).strip() in {"", "-"}:
        issues.append(f"{field}: ausente")
        return None
    try:
        result = Decimal(str(value))
    except InvalidOperation:
        issues.append(f"{field}: no numerico ({value})")
        return None
    if not result.is_finite() or abs(result) >= Decimal("100000000000000"):
        issues.append(f"{field}: fuera de rango")
        return None
    if result.as_tuple().exponent < -4:
        raise ValidationError(f"{field}: precision superior a cuatro decimales; revise el archivo.")
    if result < 0:
        issues.append(f"{field}: negativo")
    return result


def day(value, issues, field):
    if value is None:
        issues.append(f"{field}: ausente")
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    issues.append(f"{field}: fecha Excel no valida ({value})")
    return None


def import_workbook(content, filename):
    if len(content) > 10 * 1024 * 1024:
        raise ValidationError("El archivo excede 10 MB.")
    digest = hashlib.sha256(content).hexdigest()
    if db.session.scalar(db.select(ImportBatch).where(ImportBatch.sha256 == digest)):
        raise ValidationError("Este archivo ya fue importado (SHA-256 identico).")
    try:
        with ZipFile(BytesIO(content)) as archive:
            if sum(i.file_size for i in archive.infolist()) > 100 * 1024 * 1024:
                raise ValidationError("El contenido descomprimido excede 100 MB.")
        book = openpyxl.load_workbook(BytesIO(content), data_only=False, keep_links=False)
    except (BadZipFile, KeyError, ValueError, ParseError) as exc:
        raise ValidationError(f"Archivo Excel no valido: {exc}") from exc
    for sheet, table in TABLES.items():
        if sheet not in book.sheetnames or table not in book[sheet].tables:
            raise ValidationError(f"Falta {sheet}/{table}; use el libro consolidado, no un tarifario individual.")
        if book[sheet].max_row > 50000:
            raise ValidationError("Limite del prototipo: 50.000 filas por hoja.")
    batch = ImportBatch(filename=filename[:255], sha256=digest, summary={})
    db.session.add(batch)
    db.session.flush()
    cache = {}

    def dimension(model, name):
        key = (model, name)
        if key not in cache:
            item = db.session.scalar(db.select(model).where(model.name == name))
            if item is None:
                item = model(name=name)
                db.session.add(item)
                db.session.flush()
            cache[key] = item
        return cache[key]

    summary = {"rates": 0, "charges": 0, "staged": 0, "rows_with_issues": 0,
               "legacy_history": 0, "quarantined_rates": 0}
    for sheet_name, table_name in TABLES.items():
        sheet = book[sheet_name]
        left, top, right, bottom = range_boundaries(sheet.tables[table_name].ref)
        headers = [sheet.cell(top, col).value for col in range(left, right + 1)]
        if not REQUIRED.get(sheet_name, set()).issubset(headers):
            raise ValidationError(f"Encabezados incompatibles en {sheet_name}.")
        for row_number in range(top + 1, bottom + 1):
            cells = [sheet.cell(row_number, col) for col in range(left, right + 1)]
            if any(c.data_type == "f" for c in cells):
                raise ValidationError(f"{sheet_name}!{row_number}: formulas en datos; entregue valores.")
            data = dict(zip(headers, (c.value for c in cells)))
            issues = []
            raw = SourceRow(batch_id=batch.id, sheet=sheet_name, row_number=row_number,
                            payload=json.loads(json.dumps(data, default=str)), issues=[])
            db.session.add(raw)
            db.session.flush()
            summary["staged"] += 1
            if sheet_name == "TARIFARIOS_AGENTES":
                agent, origin, destination = (normalize(data[k]) for k in ("AGENTE", "POL", "POD"))
                if not agent or not origin or not destination:
                    issues.append("Sin agente o ruta: conservada en cuarentena, no cotizable")
                    summary["quarantined_rates"] += 1
                else:
                    carrier = normalize(data["NAVIERA"])
                    rate = Rate(
                        source_row_id=raw.id, batch_id=batch.id,
                        agent_id=dimension(Agent, agent).id,
                        carrier_id=dimension(Carrier, carrier).id if carrier else None,
                        origin_id=dimension(Port, origin).id,
                        destination_id=dimension(Port, destination).id,
                        valid_from=day(data["Fecha_Inicio"], issues, "Fecha_Inicio"),
                        valid_until=day(data["Fecha_Vencimiento"], issues, "Fecha_Vencimiento"),
                        freight20=amount(data["FLETE_20ST"], issues, "FLETE_20ST"),
                        freight40st=amount(data["FLETE_40ST"], issues, "FLETE_40ST"),
                        freight40hc=amount(data["FLETE_40HC"], issues, "FLETE_40HC"),
                        documentation=amount(data["DOCUMENTACION_BL"], issues, "DOCUMENTACION_BL"),
                        currency=normalize(data["Moneda"]) or None,
                        via=str(data["VIA"] or ""),
                        transit=str(data["TT"]) if data["TT"] is not None else "",
                        free_days=str(data["DIAS_LIBRES"]) if data["DIAS_LIBRES"] is not None else "",
                        source_file=str(data["Source.Name"] or ""),
                        source_status=normalize(data["Estado Vigencia"]),
                    )
                    if rate.currency != "USD":
                        issues.append("Moneda ausente o distinta de USD")
                    if not carrier:
                        issues.append("Naviera ausente")
                    db.session.add(rate)
                    summary["rates"] += 1
            elif sheet_name == "Cargos_Destino":
                agent, destination = normalize(data["Agente"]), normalize(data["Puerto_Destino"])
                if not agent or not destination:
                    issues.append("Cargo sin agente o destino: cuarentena")
                else:
                    db.session.add(DestinationCharge(
                        source_row_id=raw.id, batch_id=batch.id,
                        agent_id=dimension(Agent, agent).id,
                        destination_id=dimension(Port, destination).id,
                        carrier_scope=normalize(data["Naviera_Aplica"]),
                        concept=str(data["Concepto"] or ""), basis=normalize(data["Base"]),
                        amount=amount(data["Importe"], issues, "Importe"),
                        status=normalize(data["Estado"]),
                        valid_until=day(data["Vigencia_Hasta"], issues, "Vigencia_Hasta"),
                    ))
                    issues.append("Excel no declara moneda del cargo")
                    summary["charges"] += 1
            elif sheet_name == "Historico" and any(c.value is not None for c in cells):
                issues.append("Historico con encabezados desalineados: requiere conciliacion; no se convierte automaticamente")
                summary["legacy_history"] += 1
            raw.issues = issues
            summary["rows_with_issues"] += bool(issues)
    if "Catalogos" in book.sheetnames:
        for row in book["Catalogos"].iter_rows(min_row=2, max_col=2):
            for cell, flag in zip(row, ("is_origin", "is_destination")):
                if cell.value:
                    setattr(dimension(Port, normalize(cell.value)), flag, True)
    book.close()
    batch.summary = summary
    db.session.flush()
    return batch
