from datetime import datetime, timezone
from uuid import uuid4

from .extensions import db


def utcnow():
    return datetime.now(timezone.utc)


class ImportBatch(db.Model):
    __tablename__ = "import_batch"
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    sha256 = db.Column(db.String(64), nullable=False, unique=True)
    created_at = db.Column(db.DateTime(timezone=True), nullable=False, default=utcnow)
    summary = db.Column(db.JSON, nullable=False)


class SourceRow(db.Model):
    __tablename__ = "source_row"
    id = db.Column(db.Integer, primary_key=True)
    batch_id = db.Column(db.ForeignKey("import_batch.id"), nullable=False, index=True)
    sheet = db.Column(db.String(80), nullable=False)
    row_number = db.Column(db.Integer, nullable=False)
    payload = db.Column(db.JSON, nullable=False)
    issues = db.Column(db.JSON, nullable=False)
    __table_args__ = (db.UniqueConstraint("batch_id", "sheet", "row_number"),)


class Agent(db.Model):
    __tablename__ = "agent"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False, unique=True)


class Carrier(db.Model):
    __tablename__ = "carrier"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False, unique=True)


class Port(db.Model):
    __tablename__ = "port"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False, unique=True)
    is_origin = db.Column(db.Boolean, nullable=False, default=False)
    is_destination = db.Column(db.Boolean, nullable=False, default=False)


class Rate(db.Model):
    __tablename__ = "rate"
    id = db.Column(db.Integer, primary_key=True)
    source_row_id = db.Column(db.ForeignKey("source_row.id"), nullable=False, unique=True)
    batch_id = db.Column(db.ForeignKey("import_batch.id"), nullable=False, index=True)
    agent_id = db.Column(db.ForeignKey("agent.id"), nullable=False)
    carrier_id = db.Column(db.ForeignKey("carrier.id"))
    origin_id = db.Column(db.ForeignKey("port.id"), nullable=False)
    destination_id = db.Column(db.ForeignKey("port.id"), nullable=False)
    valid_from = db.Column(db.Date)
    valid_until = db.Column(db.Date)
    freight20 = db.Column(db.Numeric(18, 4))
    freight40st = db.Column(db.Numeric(18, 4))
    freight40hc = db.Column(db.Numeric(18, 4))
    documentation = db.Column(db.Numeric(18, 4))
    currency = db.Column(db.String(12))
    via = db.Column(db.String(200))
    transit = db.Column(db.String(120))
    free_days = db.Column(db.String(120))
    source_file = db.Column(db.String(255))
    source_status = db.Column(db.String(40))
    agent = db.relationship(Agent)
    carrier = db.relationship(Carrier)
    origin = db.relationship(Port, foreign_keys=[origin_id])
    destination = db.relationship(Port, foreign_keys=[destination_id])
    __table_args__ = (db.Index("ix_rate_route_batch", "origin_id", "destination_id", "batch_id"),)


class DestinationCharge(db.Model):
    __tablename__ = "destination_charge"
    id = db.Column(db.Integer, primary_key=True)
    source_row_id = db.Column(db.ForeignKey("source_row.id"), nullable=False, unique=True)
    batch_id = db.Column(db.ForeignKey("import_batch.id"), nullable=False, index=True)
    agent_id = db.Column(db.ForeignKey("agent.id"), nullable=False)
    destination_id = db.Column(db.ForeignKey("port.id"), nullable=False)
    carrier_scope = db.Column(db.String(120), nullable=False)
    concept = db.Column(db.String(255), nullable=False)
    basis = db.Column(db.String(40), nullable=False)
    amount = db.Column(db.Numeric(18, 4))
    status = db.Column(db.String(40), nullable=False)
    valid_until = db.Column(db.Date)
    agent = db.relationship(Agent)
    destination = db.relationship(Port)


class Quote(db.Model):
    __tablename__ = "quote"
    id = db.Column(db.String(36), primary_key=True, default=lambda: str(uuid4()))
    batch_id = db.Column(db.ForeignKey("import_batch.id"), nullable=False)
    created_at = db.Column(db.DateTime(timezone=True), nullable=False, default=utcnow)
    status = db.Column(db.String(20), nullable=False, default="draft")
    mode = db.Column(db.String(20), nullable=False)
    snapshot = db.Column(db.JSON, nullable=False)
    __table_args__ = (
        db.CheckConstraint("status IN ('draft', 'saved')", name="ck_quote_status"),
        db.CheckConstraint("mode IN ('strict', 'legacy')", name="ck_quote_mode"),
    )


class AuditEvent(db.Model):
    __tablename__ = "audit_event"
    id = db.Column(db.Integer, primary_key=True)
    quote_id = db.Column(db.ForeignKey("quote.id"), nullable=False, index=True)
    created_at = db.Column(db.DateTime(timezone=True), nullable=False, default=utcnow)
    action = db.Column(db.String(40), nullable=False)
    detail = db.Column(db.JSON, nullable=False)
