from copy import deepcopy
from datetime import date

from flask import Blueprint, abort, current_app, flash, redirect, render_template, request, url_for
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.orm import joinedload
from werkzeug.utils import secure_filename

from .extensions import db
from .models import AuditEvent, DestinationCharge, ImportBatch, Port, Quote, Rate, SourceRow
from .services.imports import import_workbook
from .services.pricing import (
    QuoteInput, ValidationError, calculate, decimal_input, rank_candidates, recalculate,
)

web = Blueprint("web", __name__)


@web.errorhandler(ValidationError)
def invalid_input(error):
    current_app.logger.warning("Validacion: %s", error)
    db.session.rollback()
    return render_template("error.html", message=str(error)), 422


@web.errorhandler(SQLAlchemyError)
def database_error(error):
    db.session.rollback()
    current_app.logger.exception("Fallo de persistencia")
    return render_template("error.html", message="No se pudo completar la operacion en la base de datos."), 500


@web.get("/")
def home():
    ports = db.session.scalars(db.select(Port).order_by(Port.name)).all()
    return render_template(
        "quote_form.html", ports=ports,
        batches=db.session.scalars(db.select(ImportBatch).order_by(ImportBatch.id.desc())).all(),
        today=date.today().isoformat(),
    )


@web.post("/quotes")
def create_quote():
    params = QuoteInput.parse(request.form)
    mode = request.form.get("mode", "strict")
    batch = db.get_or_404(ImportBatch, request.form.get("batch_id", type=int))
    rates = db.session.scalars(
        db.select(Rate).where(Rate.batch_id == batch.id)
        .options(joinedload(Rate.agent), joinedload(Rate.carrier),
                 joinedload(Rate.origin), joinedload(Rate.destination))
        .order_by(Rate.source_row_id)
    ).all()
    charges = db.session.scalars(
        db.select(DestinationCharge).where(DestinationCharge.batch_id == batch.id)
        .order_by(DestinationCharge.source_row_id)
    ).all()
    snapshot = calculate(params, rates, charges, mode)
    quote = Quote(batch_id=batch.id, mode=mode, snapshot=snapshot)
    db.session.add(quote)
    db.session.flush()
    db.session.add(AuditEvent(quote_id=quote.id, action="calculated",
                             detail={"version": snapshot["version"], "batch_sha256": batch.sha256}))
    db.session.commit()
    return redirect(url_for("web.quote_detail", quote_id=quote.id), code=303)


@web.get("/quotes/<quote_id>")
def quote_detail(quote_id):
    quote = db.get_or_404(Quote, quote_id)
    return render_template("quote_detail.html", quote=quote, report=quote.snapshot)


@web.post("/quotes/<quote_id>/save")
def save_quote(quote_id):
    quote = db.get_or_404(Quote, quote_id)
    if quote.mode != "strict":
        raise ValidationError("El modo VBA es solo de auditoria y no se guarda en el historico comercial.")
    if not quote.snapshot["options"] or decimal_input(quote.snapshot["options"][0]["total"], "Total") <= 0:
        raise ValidationError("Primero genere una cotizacion calculable con total positivo.")
    if quote.status != "saved":
        quote.status = "saved"
        db.session.add(AuditEvent(quote_id=quote.id, action="saved", detail={"provisional": True}))
        db.session.commit()
    flash("Estimacion guardada e inmutable; sigue pendiente de validacion comercial.", "success")
    return redirect(url_for("web.quote_detail", quote_id=quote.id), code=303)


@web.post("/quotes/<quote_id>/adjust")
def adjust_quote(quote_id):
    quote = db.get_or_404(Quote, quote_id)
    if quote.status != "draft" or quote.mode != "strict":
        abort(409, description="Solo se pueden ajustar borradores estrictos.")
    reason = request.form.get("reason", "").strip()
    if not 5 <= len(reason) <= 500:
        raise ValidationError("Indique el motivo del ajuste (5 a 500 caracteres).")
    snapshot = deepcopy(quote.snapshot)
    candidate = next((c for c in snapshot["candidates"]
                      if c["rate_id"] == request.form.get("rate_id", type=int)), None)
    if candidate is None:
        raise ValidationError("La oferta no pertenece a esta cotizacion.")
    old = {key: candidate[key] for key in ("origin_cost", "filling_cost")}
    for key in old:
        candidate[key] = str(decimal_input(request.form.get(key, ""), key))
    recalculate(candidate, snapshot["input"])
    snapshot["options"], snapshot["agent_count"] = rank_candidates(snapshot["candidates"], "strict")
    quote.snapshot = snapshot
    db.session.add(AuditEvent(quote_id=quote.id, action="adjusted", detail={
        "rate_id": candidate["rate_id"], "before": old,
        "after": {key: candidate[key] for key in old}, "reason": reason,
    }))
    db.session.commit()
    return redirect(url_for("web.quote_detail", quote_id=quote.id), code=303)


@web.get("/history")
def history():
    quotes = db.paginate(db.select(Quote).where(Quote.status == "saved")
                        .order_by(Quote.created_at.desc(), Quote.id.desc()), per_page=25, max_per_page=100)
    return render_template("history.html", quotes=quotes)


@web.route("/imports", methods=["GET", "POST"])
def imports():
    if request.method == "POST":
        upload = request.files.get("workbook")
        if upload is None or not upload.filename:
            raise ValidationError("Seleccione un libro consolidado XLSM o XLSX.")
        name = secure_filename(upload.filename)
        if not name.lower().endswith((".xlsm", ".xlsx")):
            raise ValidationError("Extension no permitida.")
        batch = import_workbook(upload.read(), name)
        db.session.commit()
        flash(f"Lote {batch.id} importado. Revise las incidencias antes de cotizar.", "success")
        return redirect(url_for("web.imports"), code=303)
    batches = db.session.scalars(db.select(ImportBatch).order_by(ImportBatch.id.desc())).all()
    return render_template("imports.html", batches=batches)


@web.get("/imports/<int:batch_id>/issues")
def import_issues(batch_id):
    db.get_or_404(ImportBatch, batch_id)
    # Kept as a bounded download endpoint for this small local prototype.
    rows = db.session.scalars(db.select(SourceRow).where(SourceRow.batch_id == batch_id)
                              .order_by(SourceRow.id)).all()
    return {"batch_id": batch_id, "issues": [
        {"sheet": row.sheet, "row": row.row_number, "issues": row.issues}
        for row in rows if row.issues
    ]}


@web.get("/catalog/<kind>")
def catalog(kind):
    specs = {
        "rates": (Rate, "Tarifas", ("id", "batch_id", "agent_id", "origin_id", "destination_id",
                                  "freight20", "freight40st", "freight40hc", "currency", "valid_until")),
        "charges": (DestinationCharge, "Cargos destino", ("id", "batch_id", "agent_id",
                    "destination_id", "concept", "basis", "amount", "carrier_scope", "status")),
        "ports": (Port, "Puertos (sin fusion automatica de alias)", ("id", "name", "is_origin", "is_destination")),
    }
    if kind not in specs:
        abort(404)
    model, title, columns = specs[kind]
    query = db.select(model).order_by(model.id)
    if kind != "ports" and request.args.get("batch_id", type=int):
        query = query.where(model.batch_id == request.args.get("batch_id", type=int))
    page = db.paginate(query, per_page=50, max_per_page=100)
    return render_template("catalog.html", title=title, columns=columns, page=page, kind=kind)
