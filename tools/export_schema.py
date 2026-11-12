"""Generate the initial PostgreSQL DDL from the same models used by Flask."""
from pathlib import Path
import sys

from sqlalchemy.dialects import postgresql
from sqlalchemy.schema import CreateIndex, CreateTable

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from app.extensions import db
from app import models


def main():
    dialect = postgresql.dialect()
    parts = ["-- PostgreSQL 16+. Generated with tools/export_schema.py.",
             "-- Initial empty database only. Use versioned migrations for subsequent changes.",
             "BEGIN;"]
    for table in db.metadata.sorted_tables:
        parts.append(str(CreateTable(table).compile(dialect=dialect)).strip() + ";")
        for index in sorted(table.indexes, key=lambda index: index.name):
            parts.append(str(CreateIndex(index).compile(dialect=dialect)).strip() + ";")
    parts.append("COMMIT;")
    target = Path(__file__).resolve().parents[1] / "sql" / "001_initial.sql"
    target.parent.mkdir(exist_ok=True)
    target.write_text("\n\n".join(parts) + "\n", encoding="utf-8")
    print(target)


if __name__ == "__main__":
    main()
