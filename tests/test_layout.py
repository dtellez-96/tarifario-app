from html.parser import HTMLParser

import pytest


class LayoutParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.tags = []

    def handle_starttag(self, tag, attrs):
        self.tags.append((tag, dict(attrs)))


@pytest.mark.parametrize(("path", "active"), [
    ("/", "/"),
    ("/history", "/history"),
    ("/imports", "/imports"),
    ("/catalog/rates", "/catalog/rates"),
    ("/catalog/charges", "/catalog/charges"),
    ("/catalog/ports", "/catalog/ports"),
    ("/catalog/rates?batch_id=1&page=2", "/catalog/rates"),
])
def test_shared_layout_and_active_navigation(client, path, active):
    response = client.get(path)
    assert response.status_code == 200
    parser = LayoutParser()
    parser.feed(response.get_data(as_text=True))
    for landmark in ("header", "main", "footer", "aside"):
        assert sum(tag == landmark for tag, _ in parser.tags) == 1
    assert sum(tag == "nav" and attrs.get("aria-label") == "Navegacion principal"
               for tag, attrs in parser.tags) == 1
    current_links = [
        attrs["href"] for tag, attrs in parser.tags
        if tag == "a" and attrs.get("aria-current") == "page"
    ]
    assert current_links == [active]
    ids = [attrs["id"] for _, attrs in parser.tags if "id" in attrs]
    assert len(ids) == len(set(ids))
    assert {"app-sidebar", "main-content", "open-navigation", "close-navigation",
            "toggle-sidebar"} <= set(ids)
    toggle = next(attrs for _, attrs in parser.tags if attrs.get("id") == "toggle-sidebar")
    assert toggle["aria-controls"] == "app-sidebar"
    assert toggle["aria-expanded"] == "true"
    navigation_labels = {"Cotizaciones", "Historico", "Importaciones", "Tarifas",
                         "Cargos destino", "Puertos"}
    assert navigation_labels <= {
        attrs["aria-label"] for tag, attrs in parser.tags
        if tag == "a" and attrs.get("aria-label") == attrs.get("title") and "aria-label" in attrs
    }
    links = {attrs.get("href") for tag, attrs in parser.tags if tag == "a"}
    assert {"/", "/history", "/imports", "/catalog/rates", "/catalog/charges",
            "/catalog/ports", "#main-content"} <= links
    assert b"PROTOTIPO LOCAL" in response.data


def test_quote_form_sections_and_fields(client):
    parser = LayoutParser()
    parser.feed(client.get("/").get_data(as_text=True))
    assert sum(tag == "fieldset" for tag, _ in parser.tags) == 3
    assert sum(tag == "legend" for tag, _ in parser.tags) == 3
    names = {attrs.get("name") for tag, attrs in parser.tags if tag in ("input", "select")}
    assert {"csrf_token", "quote_date", "batch_id", "mode", "origin", "destination",
            "quantity20", "quantity40st", "quantity40hc", "bl", "fob"} <= names


def test_validation_error_uses_shared_layout(client):
    response = client.post("/quotes", data={})
    assert response.status_code == 422
    assert b'id="app-sidebar"' in response.data
    assert b'<footer' in response.data
    assert b'role="alert"' in response.data
