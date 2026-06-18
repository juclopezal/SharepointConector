"""Tests de SharePointService: saneo de paths de subida, validación del campo
de filtro (anti-inyección OData) y paginación de Graph (@odata.nextLink).

Se instancia el servicio sin TokenManager y se sustituyen los helpers HTTP
(`_get`, `_put_bytes`) por dobles, de modo que no se toca la red.
"""

import pytest

from app.core.exceptions import GraphAPIError
from app.services.sharepoint import SharePointService


@pytest.fixture
def sp() -> SharePointService:
    return SharePointService(token_manager=None)


# ----------------------------------------------------------------------
# upload_file — saneo y encoding del path remoto
# ----------------------------------------------------------------------


@pytest.mark.parametrize(
    "filename",
    ["../escape.txt", "..", "a/b.txt", "a\\b.txt", "", "   ", "evil\nname.txt"],
)
async def test_upload_rejects_invalid_filename(sp, filename):
    with pytest.raises(GraphAPIError) as exc:
        await sp.upload_file("SITE", "DRIVE", "Folder", filename, b"x")
    assert exc.value.status_code == 400


@pytest.mark.parametrize("folder", ["a/../b", "..", "a\\b", "./x"])
async def test_upload_rejects_invalid_folder(sp, folder):
    with pytest.raises(GraphAPIError) as exc:
        await sp.upload_file("SITE", "DRIVE", folder, "ok.txt", b"x")
    assert exc.value.status_code == 400


async def test_upload_percent_encodes_path_segments(sp):
    captured = {}

    async def fake_put(url, data):
        captured["url"] = url
        return {"id": "1", "name": "n", "webUrl": "w"}

    sp._put_bytes = fake_put
    await sp.upload_file("SITE", "DRIVE", "Areas/Only Test", "informe #1.txt", b"x")

    assert captured["url"] == (
        "https://graph.microsoft.com/v1.0/sites/SITE/drives/DRIVE"
        "/root:/Areas/Only%20Test/informe%20%231.txt:/content"
    )


async def test_upload_without_folder_uses_root(sp):
    captured = {}

    async def fake_put(url, data):
        captured["url"] = url
        return {"id": "1", "name": "n", "webUrl": "w"}

    sp._put_bytes = fake_put
    await sp.upload_file("SITE", "DRIVE", "", "doc.txt", b"x")

    assert captured["url"].endswith("/root:/doc.txt:/content")


# ----------------------------------------------------------------------
# find_list_items_by_field — validación del campo y escape del valor
# ----------------------------------------------------------------------


@pytest.mark.parametrize(
    "field",
    ["Title eq 'x' or fields/Title", "fields/Other", "a-b", "a b", "", "a'b"],
)
async def test_find_rejects_invalid_field_name(sp, field):
    with pytest.raises(GraphAPIError) as exc:
        await sp.find_list_items_by_field("SITE", "LIST", field, "v")
    assert exc.value.status_code == 400


async def test_find_escapes_and_encodes_filter(sp):
    captured = {}

    async def fake_get(url, extra_headers=None):
        captured["url"] = url
        captured["headers"] = extra_headers
        return {"value": []}

    sp._get = fake_get
    await sp.find_list_items_by_field("SITE", "LIST", "Title", "O'Brien & Co")

    # comillas simples duplicadas (OData) y expresión percent-encodeada
    assert "$filter=fields%2FTitle%20eq%20%27O%27%27Brien%20%26%20Co%27" in captured["url"]
    assert captured["headers"]["Prefer"] == "HonorNonIndexedQueriesWarningMayFailRandomly"


# ----------------------------------------------------------------------
# find_list_items_for_upsert — filtro combinado clave (+ periodo)
# ----------------------------------------------------------------------


async def test_upsert_lookup_key_only_filter(sp):
    captured = {}

    async def fake_get(url, extra_headers=None):
        captured["url"] = url
        captured["headers"] = extra_headers
        return {"value": []}

    sp._get = fake_get
    await sp.find_list_items_for_upsert(
        "SITE", "LIST", key_field="TicketId", key_value="INC-1"
    )

    # solo la condición de clave, percent-encodeada, con la cabecera Prefer
    assert "$filter=fields%2FTicketId%20eq%20%27INC-1%27" in captured["url"]
    assert "ge%27" not in captured["url"]  # sin rango de fecha
    assert captured["headers"]["Prefer"] == "HonorNonIndexedQueriesWarningMayFailRandomly"


async def test_upsert_lookup_key_and_period_filter(sp):
    captured = {}

    async def fake_get(url, extra_headers=None):
        captured["url"] = url
        return {"value": []}

    sp._get = fake_get
    await sp.find_list_items_for_upsert(
        "SITE",
        "LIST",
        key_field="TicketId",
        key_value="INC-1",
        date_field="Created",
        period_start="2026-06-01T00:00:00Z",
        period_end="2026-07-01T00:00:00Z",
    )

    # las tres condiciones unidas con 'and' (percent-encodeado: %20and%20)
    assert "%20and%20" in captured["url"]
    assert "fields%2FCreated%20ge%20%272026-06-01T00%3A00%3A00Z%27" in captured["url"]
    assert "fields%2FCreated%20lt%20%272026-07-01T00%3A00%3A00Z%27" in captured["url"]


@pytest.mark.parametrize("bad_field", ["Title eq 'x'", "fields/Other", "a-b", ""])
async def test_upsert_lookup_rejects_invalid_key_field(sp, bad_field):
    with pytest.raises(GraphAPIError) as exc:
        await sp.find_list_items_for_upsert(
            "SITE", "LIST", key_field=bad_field, key_value="v"
        )
    assert exc.value.status_code == 400


async def test_upsert_lookup_rejects_invalid_date_field(sp):
    with pytest.raises(GraphAPIError) as exc:
        await sp.find_list_items_for_upsert(
            "SITE",
            "LIST",
            key_field="TicketId",
            key_value="v",
            date_field="Created eq 'x'",
            period_start="2026-06-01T00:00:00Z",
            period_end="2026-07-01T00:00:00Z",
        )
    assert exc.value.status_code == 400


# ----------------------------------------------------------------------
# Paginación — @odata.nextLink
# ----------------------------------------------------------------------


async def test_get_all_follows_next_link(sp):
    pages = {
        "page1": {"value": [{"id": "1"}, {"id": "2"}], "@odata.nextLink": "page2"},
        "page2": {"value": [{"id": "3"}], "@odata.nextLink": "page3"},
        "page3": {"value": [{"id": "4"}]},
    }
    calls = []

    async def fake_get(url, extra_headers=None):
        calls.append(url)
        return pages[url]

    sp._get = fake_get
    items = await sp._get_all("page1")

    assert [i["id"] for i in items] == ["1", "2", "3", "4"]
    assert calls == ["page1", "page2", "page3"]


async def test_list_site_lists_aggregates_pages(sp):
    first = {
        "value": [{"id": "L1"}],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/sites/SITE/lists?$skiptoken=x",
    }
    second = {"value": [{"id": "L2"}]}

    async def fake_get(url, extra_headers=None):
        return second if "skiptoken" in url else first

    sp._get = fake_get
    lists = await sp.list_site_lists("SITE")

    assert [l["id"] for l in lists] == ["L1", "L2"]
