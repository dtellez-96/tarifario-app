import os
import secrets
from pathlib import Path

import click
from flask import Flask, abort, request
from sqlalchemy import event
from sqlalchemy.engine import Engine
from flask_wtf.csrf import CSRFError

from .extensions import csrf, db


@event.listens_for(Engine, "connect")
def sqlite_foreign_keys(connection, _):
    if connection.__class__.__module__ == "sqlite3":
        cursor = connection.cursor()
        cursor.execute("PRAGMA foreign_keys=ON")
        cursor.close()


def create_app(test_config=None):
    app = Flask(__name__, instance_relative_config=True)
    app.config.from_mapping(
        SQLALCHEMY_DATABASE_URI=os.environ.get("DATABASE_URL", "sqlite:///tarifario.db"),
        SQLALCHEMY_TRACK_MODIFICATIONS=False,
        MAX_CONTENT_LENGTH=10 * 1024 * 1024,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        LOCAL_ONLY=True,
    )
    if test_config:
        app.config.update(test_config)
    Path(app.instance_path).mkdir(parents=True, exist_ok=True)
    if not app.config.get("SECRET_KEY"):
        key = os.environ.get("SECRET_KEY")
        if not key:
            path = Path(app.instance_path) / "secret-key"
            if not path.exists():
                try:
                    with path.open("x", encoding="ascii") as stream:
                        stream.write(secrets.token_hex(32))
                except FileExistsError:
                    pass
            key = path.read_text(encoding="ascii").strip()
        app.config["SECRET_KEY"] = key
    db.init_app(app)
    csrf.init_app(app)
    from . import models
    from .routes import web
    from .services.pricing import money
    app.register_blueprint(web)
    app.jinja_env.filters["money"] = money

    @app.before_request
    def local_only():
        if app.config["LOCAL_ONLY"] and request.remote_addr not in {"127.0.0.1", "::1"}:
            abort(403, description="Prototipo local: acceso remoto deshabilitado.")

    @app.after_request
    def headers(response):
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; style-src 'self'; script-src 'self'; img-src 'self' data:; "
            "frame-ancestors 'none'; form-action 'self'; base-uri 'self'"
        )
        response.headers["Cache-Control"] = "no-store"
        return response

    @app.errorhandler(CSRFError)
    def csrf_error(error):
        app.logger.warning("CSRF rechazado: %s", error.description)
        return {"error": "Formulario vencido o token CSRF invalido. Recargue la pagina."}, 400

    @app.cli.command("init-db")
    def init_db():
        """Create tables for a new local database; does not migrate existing schemas."""
        db.create_all()
        click.echo("Esquema creado. Para produccion use migraciones versionadas.")

    @app.cli.command("import-workbook")
    @click.argument("path", type=click.Path(exists=True, dir_okay=False, path_type=Path))
    def import_command(path):
        from .services.imports import import_workbook
        from .services.pricing import ValidationError
        from sqlalchemy.exc import SQLAlchemyError
        try:
            batch = import_workbook(path.read_bytes(), path.name)
            db.session.commit()
        except (ValidationError, SQLAlchemyError) as exc:
            db.session.rollback()
            app.logger.exception("Importacion CLI fallida")
            raise click.ClickException(str(exc)) from exc
        click.echo(f"Lote {batch.id}: {batch.summary}")

    return app
