from pathlib import Path

import pytest

from app import create_app
from app.extensions import db
from app.services.imports import import_workbook

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "Cotizador_FCL_Automatico_Visual_Final_v6.xlsm"


@pytest.fixture(scope="module")
def application():
    app = create_app({
        "TESTING": True, "SQLALCHEMY_DATABASE_URI": "sqlite://",
        "SECRET_KEY": "test-key-not-for-deployment", "WTF_CSRF_ENABLED": False,
    })
    with app.app_context():
        db.create_all()
        import_workbook(WORKBOOK.read_bytes(), WORKBOOK.name)
        db.session.commit()
        yield app
        db.session.remove()
        db.drop_all()


@pytest.fixture
def client(application):
    return application.test_client()
