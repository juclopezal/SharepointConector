"""Tests de SharePointService: saneo de paths de subida, validación del campo
de filtro (anti-inyección OData) y paginación de Graph (@odata.nextLink).

Se instancia el servicio sin TokenManager y se sustituyen los helpers HTTP
(`_get`, `_put_bytes`) por dobles, de modo que no se toca la red.
"""

import pytest

from app.core.exceptions import GraphAPIError
from app.services.sharepoint import SharePointService, _to_odata_literal


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


# ----------------------------------------------------------------------
# _to_odata_literal — traducción de tipos JSON a literal OData
# ----------------------------------------------------------------------


def test_odata_literal_bool_renders_as_1_0():
    # NUNCA true/false: Graph ignora en silencio esas cláusulas en columnas Sí/No
    assert _to_odata_literal(True) == "1"
    assert _to_odata_literal(False) == "0"
    assert _to_odata_literal(True) != "true"


def test_odata_literal_string_quoted_with_doubling():
    assert _to_odata_literal("O'Hara") == "'O''Hara'"
    assert _to_odata_literal("L02") == "'L02'"


def test_odata_literal_numbers_unquoted():
    assert _to_odata_literal(35) == "35"
    assert _to_odata_literal(3.5) == "3.5"


def test_odata_literal_int_one_is_not_bool_true():
    # bool es subclase de int: el chequeo de bool debe ir primero, y un int 1
    # debe seguir siendo el número 1 (no el literal booleano)
    assert _to_odata_literal(1) == "1"
    assert _to_odata_literal(True) == "1"
    # mismos literales aquí, pero la distinción se valida contra el esquema
    # (un int no pasa una columna boolean y viceversa) — ver tests de búsqueda


# ----------------------------------------------------------------------
# search_list_items — filtros validados contra el esquema de columnas
# ----------------------------------------------------------------------

_DSL_COLUMNS = [
    {"name": "Title", "text": {}},
    {"name": "Entorno", "text": {}},
    {"name": "_x00da_ltima", "boolean": {}},
    {"name": "Cliente_x002d_LIBSA", "lookup": {}},
    {"name": "hohl", "number": {}},
]


def _wire_search(sp, items_pages=None, columns=None):
    """Sustituye `_get` con un doble que sirve columnas e ítems y captura URLs."""
    captured = {"urls": [], "headers": None, "columns_calls": 0}
    pages = items_pages or [{"value": []}]

    async def fake_get(url, extra_headers=None):
        captured["urls"].append(url)
        if "/columns" in url:
            captured["columns_calls"] += 1
            return {"value": columns if columns is not None else _DSL_COLUMNS}
        captured["headers"] = extra_headers
        return pages.pop(0) if pages else {"value": []}

    sp._get = fake_get
    return captured


async def test_search_multi_filter_and_boolean_literal(sp):
    captured = _wire_search(sp)
    await sp.search_list_items(
        "SITE",
        "LIST",
        filters=[("Entorno", "L02"), ("_x00da_ltima", True)],
    )

    items_url = captured["urls"][-1]
    # AND percent-encodeado en una sola expresión; booleano como 1, sin comillas
    assert (
        "$filter=fields%2FEntorno%20eq%20%27L02%27"
        "%20and%20fields%2F_x00da_ltima%20eq%201" in items_url
    )
    assert captured["headers"]["Prefer"] == "HonorNonIndexedQueriesWarningMayFailRandomly"


async def test_search_string_value_escapes_quotes(sp):
    captured = _wire_search(sp)
    await sp.search_list_items("SITE", "LIST", filters=[("Title", "O'Hara")])

    assert "%27O%27%27Hara%27" in captured["urls"][-1]


async def test_search_lookup_id_value_unquoted(sp):
    captured = _wire_search(sp)
    await sp.search_list_items(
        "SITE", "LIST", filters=[("Cliente_x002d_LIBSALookupId", 35)]
    )

    assert "fields%2FCliente_x002d_LIBSALookupId%20eq%2035" in captured["urls"][-1]


async def test_search_unknown_field_rejected_naming_it(sp):
    _wire_search(sp)
    with pytest.raises(GraphAPIError) as exc:
        await sp.search_list_items("SITE", "LIST", filters=[("NoExiste", "x")])
    assert exc.value.status_code == 400
    assert "NoExiste" in exc.value.message


async def test_search_string_against_boolean_column_rejected(sp):
    # el caso que motivó la validación: "true" como string NO debe convertirse
    # en una consulta sin filtrar
    _wire_search(sp)
    with pytest.raises(GraphAPIError) as exc:
        await sp.search_list_items("SITE", "LIST", filters=[("_x00da_ltima", "true")])
    assert exc.value.status_code == 400
    assert "boolean" in exc.value.message


async def test_search_string_against_numeric_column_rejected(sp):
    _wire_search(sp)
    with pytest.raises(GraphAPIError) as exc:
        await sp.search_list_items("SITE", "LIST", filters=[("hohl", "35")])
    assert exc.value.status_code == 400
    assert "number" in exc.value.message


async def test_search_int_against_boolean_column_rejected(sp):
    # bool es subclase de int, pero un 1 entero no vale como booleano
    _wire_search(sp)
    with pytest.raises(GraphAPIError) as exc:
        await sp.search_list_items("SITE", "LIST", filters=[("_x00da_ltima", 1)])
    assert exc.value.status_code == 400


async def test_search_bool_against_numeric_column_rejected(sp):
    _wire_search(sp)
    with pytest.raises(GraphAPIError) as exc:
        await sp.search_list_items("SITE", "LIST", filters=[("hohl", True)])
    assert exc.value.status_code == 400


async def test_search_base_lookup_column_rejected_with_guidance(sp):
    # columna lookup referenciada sin el sufijo LookupId → 400 explícito
    _wire_search(sp)
    with pytest.raises(GraphAPIError) as exc:
        await sp.search_list_items(
            "SITE", "LIST", filters=[("Cliente_x002d_LIBSA", "35")]
        )
    assert exc.value.status_code == 400
    assert "LookupId" in exc.value.message


async def test_search_invalid_field_name_rejected_in_service(sp):
    # defensa en profundidad: aunque el schema ya lo valide, el servicio repite
    _wire_search(sp)
    with pytest.raises(GraphAPIError) as exc:
        await sp.search_list_items("SITE", "LIST", filters=[("a-b", "x")])
    assert exc.value.status_code == 400


async def test_search_without_filters_skips_columns_fetch(sp):
    captured = _wire_search(sp)
    await sp.search_list_items("SITE", "LIST", filters=None)

    assert captured["columns_calls"] == 0
    assert "$filter" not in captured["urls"][-1]


async def test_search_columns_cache_avoids_second_fetch(sp):
    captured = _wire_search(sp)
    await sp.search_list_items("SITE", "LIST", filters=[("Entorno", "L02")])
    await sp.search_list_items("SITE", "LIST", filters=[("Entorno", "L01")])

    assert captured["columns_calls"] == 1


async def test_search_orderby_default_asc_and_explicit_desc(sp):
    captured = _wire_search(sp)
    await sp.search_list_items("SITE", "LIST", order_by=("Created", "asc"))
    assert "$orderby=fields%2FCreated%20asc" in captured["urls"][-1]

    await sp.search_list_items("SITE", "LIST", order_by=("Created", "desc"))
    assert "$orderby=fields%2FCreated%20desc" in captured["urls"][-1]


async def test_search_orderby_threshold_error_propagates(sp):
    async def fake_get(url, extra_headers=None):
        raise GraphAPIError(
            400,
            "The request is not supported: field not usable for sorting "
            "(list view threshold exceeded)",
        )

    sp._get = fake_get
    with pytest.raises(GraphAPIError) as exc:
        await sp.search_list_items("SITE", "LIST", order_by=("Grupo", "asc"))
    assert "threshold" in exc.value.message


async def test_search_truncates_at_top_across_pages_with_has_more(sp):
    pages = [
        {"value": [{"id": "1"}, {"id": "2"}], "@odata.nextLink": "next1"},
        {"value": [{"id": "3"}, {"id": "4"}], "@odata.nextLink": "next2"},
        {"value": [{"id": "5"}]},
    ]
    _wire_search(sp, items_pages=pages)
    items, has_more = await sp.search_list_items("SITE", "LIST", top=3)

    assert [i["id"] for i in items] == ["1", "2", "3"]
    assert has_more is True


async def test_search_fewer_matches_than_top_has_more_false(sp):
    pages = [{"value": [{"id": "1"}, {"id": "2"}]}]
    _wire_search(sp, items_pages=pages)
    items, has_more = await sp.search_list_items("SITE", "LIST", top=100)

    assert len(items) == 2
    assert has_more is False


async def test_search_exact_top_without_next_link_has_more_false(sp):
    pages = [{"value": [{"id": "1"}, {"id": "2"}, {"id": "3"}]}]
    _wire_search(sp, items_pages=pages)
    items, has_more = await sp.search_list_items("SITE", "LIST", top=3)

    assert len(items) == 3
    assert has_more is False
