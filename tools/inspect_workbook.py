"""Extract workbook evidence without executing VBA or external connections."""
import hashlib
import base64
import io
import json
import struct
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

import openpyxl
from openpyxl.formula.translate import Translator
from oletools.olevba import VBA_Parser

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "Cotizador_FCL_Automatico_Visual_Final_v6.xlsm"
OUTPUT = ROOT / "docs" / "evidence"


def main():
    OUTPUT.mkdir(parents=True, exist_ok=True)
    parser = VBA_Parser(str(SOURCE))
    modules = []
    for _, stream, name, code in parser.extract_macros():
        target = OUTPUT / Path(name).name
        target.write_text(code, encoding="utf-8")
        modules.append({"name": name, "stream": stream, "lines": len(code.splitlines())})
    parser.close()
    workbook = openpyxl.load_workbook(SOURCE, data_only=False)
    cached = openpyxl.load_workbook(SOURCE, data_only=True)
    summary = {
        "sha256": hashlib.sha256(SOURCE.read_bytes()).hexdigest(),
        "modules": modules,
        "defined_names": {key: str(value) for key, value in workbook.defined_names.items()},
        "sheets": [],
    }
    for sheet in workbook:
        cells = []
        formulas = []
        for row in sheet:
            for cell in row:
                if cell.value is not None:
                    value = cell.value
                    if cell.data_type == "f":
                        formulas.append({"cell": cell.coordinate, "formula": value,
                                         "cached": cached[sheet.title][cell.coordinate].value})
                    cells.append({"cell": cell.coordinate, "value": value,
                                  "type": cell.data_type, "format": cell.number_format})
        record = {
            "name": sheet.title, "state": sheet.sheet_state,
            "rows": sheet.max_row, "columns": sheet.max_column,
            "nonempty": len(cells), "formula_count": len(formulas),
            "errors": [c for c in cells if c["type"] == "e"],
            "merges": [str(r) for r in sheet.merged_cells.ranges],
            "validations": [str(v) for v in sheet.data_validations.dataValidation],
            "tables": [str(t) for t in sheet.tables.values()],
            "protection": str(sheet.protection),
            "hidden_columns": [key for key, dim in sheet.column_dimensions.items() if dim.hidden],
            "hidden_rows": [key for key, dim in sheet.row_dimensions.items() if dim.hidden],
            "conditional_formats": [str(c) for c in sheet.conditional_formatting],
            "sample": cells[:65],
        }
        summary["sheets"].append(record)
        (OUTPUT / f"{sheet.title}.json").write_text(
            json.dumps({"metadata": record, "cells": cells, "formulas": formulas},
                       ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    with ZipFile(SOURCE) as archive:
        names = archive.namelist()
        summary["package_parts"] = names
        for name in names:
            if any(part in name for part in (
                "vmlDrawing", "ctrlProp", "drawing1.xml", "connections.xml",
                "queryTable", "customXml/item1.xml", "workbook.xml",
                "webextension", "sheet5.xml.rels", "sheet8.xml.rels",
            )):
                (OUTPUT / name.replace("/", "__")).write_bytes(archive.read(name))
        mashup = base64.b64decode(ElementTree.fromstring(archive.read("customXml/item1.xml")).text)
        _, length = struct.unpack("<II", mashup[:8])
        with ZipFile(io.BytesIO(mashup[8:8 + length])) as package:
            scripts = [package.read(name).decode("utf-8-sig")
                       for name in package.namelist() if name.endswith(".m")]
            (OUTPUT / "PowerQuery.m").write_text("\n".join(scripts), encoding="utf-8")
    patterns = Counter()
    for row in workbook["Comparativo"]:
        for cell in row:
            if cell.data_type == "f":
                patterns[Translator(cell.value, origin=cell.coordinate)
                         .translate_formula(cell.column_letter + "5")] += 1
    active_rows = list(workbook["TARIFARIOS_AGENTES"].values)[1:]
    summary["profile"] = {
        "formula_families": dict(patterns),
        "duplicate_active_rows": len(active_rows) - len(set(active_rows)),
        "missing_route_rows": sum(not row[4] or not row[5] for row in active_rows),
        "missing_freight20": sum(row[8] is None for row in active_rows),
        "missing_documentation": sum(row[11] is None for row in active_rows),
        "missing_currency": sum(row[14] is None for row in active_rows),
        "agents": dict(Counter(row[1] for row in active_rows)),
    }
    (OUTPUT / "inventory.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(json.dumps({"sha256": summary["sha256"], "modules": modules,
                      "sheets": [{k: s[k] for k in ("name", "state", "rows", "columns",
                                                     "nonempty", "formula_count")}
                                 for s in summary["sheets"]]}, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
