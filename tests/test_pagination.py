from datetime import datetime, timezone
from html.parser import HTMLParser
from urllib.parse import parse_qs, urlsplit

import pytest

from app import create_app
from app.extensions import db
from app.models import ImportBatch, Quote


class TablePage(HTMLParser):
    def __init__(self, response):
        super().__init__()
        assert response.status_code == 200
        self.links = []
        self.disabled = []
        self.current = []
        self.rows = []
        self.text = []
        self.in_pagination = False
        self.in_body = False
        self.cell = None
        self.feed(response.get_data(as_text=True))

    def handle_starttag(self, tag, attributes):
        attrs = dict(attributes)
        if tag == "nav":
            self.in_pagination = attrs.get("class") == "table-pagination"
        if self.in_pagination:
            if tag == "a":
                self.links.append(attrs)
            if tag == "button" and "disabled" in attrs:
                self.disabled.append(attrs["aria-label"])
            if attrs.get("aria-current") == "page":
                self.current.append(attrs["aria-label"])
        if tag == "tbody":
            self.in_body = True
        if tag == "tr" and self.in_body:
            self.rows.append([])
        if tag == "td" and self.in_body:
            self.cell = []

    def handle_data(self, data):
        if self.in_pagination:
            self.text.append(data)
        if self.cell is not None:
            self.cell.append(data)

    def handle_endtag(self, tag):
        if tag == "nav":
            self.in_pagination = False
        if tag == "tbody":
            self.in_body = False
        if tag == "td" and self.cell is not None:
            self.rows[-1].append("".join(self.cell).strip())
            self.cell = None

    @property
    def summary(self):
        return " ".join(" ".join(self.text).split())


@pytest.mark.parametrize(("kind", "total"), [("rates", 360), ("charges", 100), ("ports", 23)])
def test_catalog_pagination_covers_all_records(client, kind, total):
    ids = []
    pages = (total + 49) // 50
    for number in range(1, pages + 1):
        table = TablePage(client.get(f"/catalog/{kind}?batch_id=1&page={number}"))
        start, end = (number - 1) * 50 + 1, min(number * 50, total)
        assert f"Mostrando {start} - {end} de {total} registros" in table.summary
        assert f"Pagina {number} de {pages}" in table.summary
        assert table.current == [f"Pagina {number}, actual"]
        assert len(table.rows) == end - start + 1
        assert ("Pagina anterior" in table.disabled) == (number == 1)
        assert ("Pagina siguiente" in table.disabled) == (number == pages)
        ids.extend(int(row[0]) for row in table.rows)
        for link in table.links:
            url = urlsplit(link["href"])
            query = parse_qs(url.query)
            assert url.path == f"/catalog/{kind}"
            assert url.fragment == "results-table"
            assert query["batch_id"] == ["1"]
            target = int(query["page"][0])
            assert 1 <= target <= pages and target != number
            if link.get("rel") == "prev":
                assert target == number - 1
            if link.get("rel") == "next":
                assert target == number + 1
    assert len(set(ids)) == total
    assert ids == sorted(ids)


def test_catalog_empty_and_ellipsis(client):
    empty = TablePage(client.get("/catalog/rates?batch_id=999999"))
    assert "0 registros para mostrar" in empty.summary
    assert not empty.links and not empty.current
    assert empty.rows == [["Sin registros para los filtros seleccionados."]]
    middle = TablePage(client.get("/catalog/rates?page=4"))
    assert "\u2026" in middle.summary
    targets = {int(parse_qs(urlsplit(link["href"]).query)["page"][0]) for link in middle.links}
    assert targets == {1, 3, 5, 8}
    assert all("batch_id" not in parse_qs(urlsplit(link["href"]).query) for link in middle.links)


@pytest.fixture
def history_client(request):
    app = create_app({
        "TESTING": True, "SQLALCHEMY_DATABASE_URI": "sqlite://",
        "SECRET_KEY": "pagination-test-key", "WTF_CSRF_ENABLED": False,
    })
    with app.app_context():
        db.create_all()
        batch = ImportBatch(filename="test.xlsx", sha256="0" * 64, summary={})
        db.session.add(batch)
        db.session.flush()
        for i in range(request.param + 1):
            db.session.add(Quote(
                id=f"00000000-0000-0000-0000-{i:012d}", batch_id=batch.id,
                status="saved" if i < request.param else "draft", mode="strict",
                created_at=datetime(2026, 11, 10, tzinfo=timezone.utc),
                snapshot={"input": {"quote_date": "2026-11-10", "origin": "SHANGHAI",
                                    "destination": "CARTAGENA"},
                          "options": [{"agent": "Agente de prueba", "total": "100"}]},
            ))
        db.session.commit()
        yield app.test_client(), request.param
        db.session.remove()
        db.drop_all()


@pytest.mark.parametrize("history_client", [0, 1, 25, 26, 56], indirect=True)
def test_history_boundaries_and_stable_order(history_client):
    client, total = history_client
    pages = (total + 24) // 25
    ids = []
    for number in range(1, max(pages, 1) + 1):
        table = TablePage(client.get(f"/history?page={number}"))
        if not total:
            assert "0 registros para mostrar" in table.summary
            assert table.rows == [["Sin cotizaciones guardadas."]]
            assert not table.links and not table.current
            continue
        start, end = (number - 1) * 25 + 1, min(number * 25, total)
        assert f"Mostrando {start} - {end} de {total} registros" in table.summary
        assert f"Pagina {number} de {pages}" in table.summary
        assert table.current == [f"Pagina {number}, actual"]
        assert len(table.rows) == end - start + 1
        assert ("Pagina anterior" in table.disabled) == (number == 1)
        assert ("Pagina siguiente" in table.disabled) == (number == pages)
        ids.extend(row[0] for row in table.rows)
        for link in table.links:
            url = urlsplit(link["href"])
            assert url.path == "/history" and url.fragment == "results-table"
            target = int(parse_qs(url.query)["page"][0])
            assert 1 <= target <= pages and target != number
    assert ids == [f"00000000-0000-0000-0000-{i:012d}" for i in reversed(range(total))]
    assert client.get(f"/history?page={max(pages, 1) + 1}").status_code == 404


@pytest.mark.parametrize("page", ["0", "-1", "abc", "999999"])
def test_invalid_catalog_page_is_rejected(client, page):
    assert client.get(f"/catalog/rates?page={page}").status_code == 404
