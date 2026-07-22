## 1. Schemas (`app/schemas/sharepoint.py`)

- [x] 1.1 Add `SearchFilter` model: `field: str` (pattern `^[A-Za-z0-9_]+$`),
      `value: str | bool | int | float` (reject null/object/array via 422; verify
      Pydantic v2 smart-mode keeps `bool` distinct from `int`).
- [x] 1.2 Add `OrderBy` model: `field: str` (same pattern), `direction: Literal["asc",
      "desc"] = "asc"`.
- [x] 1.3 Add `ListItemsSearchByUrlRequest`: `sharepoint_url: str`,
      `filters: list[SearchFilter] | None = None` (`max_length=15`, no `min_length` —
      an explicit `[]` is valid and behaves exactly like omitting the field: no
      `$filter` applied), `order_by: OrderBy | None`, `top: int = 100` (`ge=1,
      le=5000`). Set `model_config = ConfigDict(extra="forbid")`.
- [x] 1.4 Add `ListItemsSearchByUrlItem` (`id`, `webUrl: str | None`, `fields: dict`)
      and `ListItemsSearchByUrlResponse` (`total: int`, `items: list[...Item]`,
      `has_more: bool`, `site_id: str`, `list_id: str`).

## 2. Column schema fetch + cache (`app/services/sharepoint.py` or new module)

- [x] 2.1 Add `get_list_columns(site_id, list_id) -> list[dict]` calling
      `GET /sites/{site_id}/lists/{list_id}/columns` via existing `_get_all`.
- [x] 2.2 Add an in-process cache keyed by `(site_id, list_id)` storing
      `(columns, fetched_at)`, TTL-checked on read (default 300s), re-fetching on miss
      or staleness (per design.md decision 3 — no new cache dependency).
- [x] 2.3 Add a column-lookup helper that maps a requested filter `field` to its
      column definition, including the `{Base}LookupId` → base `lookup`-type column
      fallback described in design.md decision 4. A **direct** match on a `lookup`-type
      column (i.e. `field` given without the `LookupId` suffix) is a distinct outcome
      from "field not found" — surface it separately so step 3.2 can reject it with its
      own explicit message.
- [x] 2.4 Add a helper mapping a column's Graph type facet to the expected JSON type
      family (`boolean`→bool, `number`/`currency`→int|float, `text`/`note`/`choice`/
      `dateTime`/`personOrGroup`→str; `lookup` accessed via `LookupId` suffix→int).
      A direct (non-suffixed) `lookup` facet match does NOT fall through to this
      mapping — it's rejected in 2.3/3.2 before reaching here.

## 3. Filter validation + OData literal translation (`app/services/sharepoint.py`)

- [x] 3.1 Add `_to_odata_literal(value: str | bool | int | float) -> str`: bool→`"1"`/
      `"0"` (check `bool` before `int`/`float`), str→quoted with `'` doubled,
      int/float→unquoted `str(value)`.
- [x] 3.2 Add a filter-validation step: for each `SearchFilter`, resolve its column
      (task 2.3), reject with `GraphAPIError(400, ...)` if: the field doesn't exist;
      the field directly names a `lookup`-type column without the `LookupId` suffix
      (explicit message pointing the caller to the `LookupId` field); or `type(value)`
      doesn't match the expected type family (task 2.4) — no coercion, per design.md
      decision 4.
- [x] 3.3 Re-validate each filter's `field` against `_FIELD_NAME_RE` at the service
      layer (defense in depth, matching existing `find_list_items_by_field` /
      `find_list_items_for_upsert` pattern).

## 4. Search service method (`app/services/sharepoint.py`)

- [x] 4.1 Add `search_list_items(site_id, list_id, *, filters, order_by, top) ->
      tuple[list[dict], bool]`: builds the AND-joined `$filter` (reusing 3.1/3.2),
      optional `$orderby`, percent-encodes the full expression, sends
      `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`.
- [x] 4.2 Add bounded pagination (`_get_bounded` per design.md decision 5): follow
      `@odata.nextLink` until `top` items collected or exhausted; return
      `(items[:top], has_more)`.
- [x] 4.3 Let Graph's `notSupported` list-view-threshold error on `$orderby` propagate
      as-is via existing `raise_from_httpx` (no special handling/suppression).

## 5. Endpoint (`app/api/v1/endpoints/sharepoint.py`)

- [x] 5.1 Add `POST /list/items:search` on the existing `sharepoint` router →
      `ListItemsSearchByUrlResponse`, resolving the list via
      `resolver.resolve_list()` (unchanged) then calling `search_list_items`.
- [x] 5.2 Return `200` with `total = len(items)` (post-truncation), `has_more`,
      `site_id`, `list_id`; empty `filters`/no matches → `200` with `total: 0`.

## 6. Tests (`tests/test_sharepoint_service.py`, `tests/test_sharepoint_endpoints.py`)

- [x] 6.1 Unit tests for `_to_odata_literal`: bool `True`/`False` → `1`/`0` (not
      `true`/`false`), string quote-doubling (`O'Hara` → `'O''Hara'`), int/float
      unquoted, and a case proving JSON `1` (int) is treated distinctly from JSON
      `true` (bool).
- [x] 6.2 Unit tests for multi-filter AND `$filter` construction (2+ fields joined
      with `and`, percent-encoded as one expression).
- [x] 6.3 Unit tests for column-schema validation: unknown field → 400 naming the
      field; type mismatch (string vs boolean column, string vs numeric column) →
      400 naming the expected type; `LookupId`-suffixed field resolving against its
      base `lookup` column → accepted; the same base `lookup` column addressed
      **without** the `LookupId` suffix → 400 explaining the `LookupId` requirement
      (NOT silently accepted as a string).
- [x] 6.4 Unit test proving the column-schema cache avoids a second Graph call for a
      repeated `(site_id, list_id)` within the TTL window.
- [x] 6.5 Unit tests for `order_by` → `$orderby` construction (default `asc`, explicit
      `desc`) and for propagating a simulated `notSupported` threshold error.
- [x] 6.6 Unit tests for bounded pagination: fewer matches than `top` → `has_more:
      false`; more matches than `top` (spanning multiple simulated pages) → truncated
      to `top`, `has_more: true`.
- [x] 6.7 Schema tests: `filters` array with >15 entries → 422; explicit `"filters":
      []` → 200, unfiltered, same response as omitting `filters` entirely (NOT 422);
      unknown top-level field (e.g. `filter_by`) → 422; `value` as `null`/object/array
      → 422; `top` outside `[1, 5000]` → 422.
- [x] 6.8 Endpoint-level test (`TestClient` + `fake_sp` pattern from
      `tests/conftest.py`) covering a full request/response round trip, and
      `test_endpoints_present_in_openapi`-style assertion that `:search` appears in
      `app.openapi()["paths"]`.
- [x] 6.9 Run the full existing suite (`pytest`) and confirm no regressions.

## 7. Documentation

- [x] 7.1 README.md: add `POST /v1/sharepoint/list/items:search` row/section under
      "SharePoint (por URL)", with the request/response shape and the boolean→`1`/`0`
      rule called out explicitly.
- [x] 7.2 ARQUITECTURA.md: add a `#### POST /v1/sharepoint/list/items:search`
      subsection (mirroring the existing `:upsert` subsection), documenting the
      column-schema validation step, the cache, and the `order_by` threshold
      limitation.
- [x] 7.3 arquitecturasUML.md: update the PlantUML diagram to reflect the new
      endpoint/service calls (including the new `GET .../columns` Graph call).
- [x] 7.4 doc/CHANGELOG.md: add a `## v2.4.0 — <date>` entry, `### Feature: List
      Items Search by URL (SPEC-004)`, with context/solution prose, a JSON example,
      and the boolean/schema-validation/order_by-limitation notes.
- [x] 7.5 Bump `VERSION` to `2.4.0`.
- [x] 7.6 Add `requirements/SPEC-004_List_Items_Search_By_Url.md` following the
      existing SPEC-00N convention (see `requirements/00_SDD_GUIDELINES.md`).

## 8. Empirical verification (mandatory, against the real DSL list)

- [x] 8.1 Local Docker Desktop: `docker compose -f devops/docker-compose.yml up -d
      --build`; verify `http://localhost:8001/health` reports the new version and
      `/openapi.json` includes `/v1/sharepoint/list/items:search`.
- [x] 8.2 Against `https://latinia2com-portal8.sharepoint.com/Oper/Lists/DSL/Allitemsg.aspx`:
      `Entorno = "no-existe-xyz"` → `total: 0`.
- [x] 8.3 Same list: `_x00da_ltima = true` only → inspect every returned item's
      `fields._x00da_ltima` and confirm all are `true` (any `false` means the clause
      was ignored — treat as a blocking bug, not a data anomaly).
- [x] 8.4 Same list: `Cliente_x002d_LIBSALookupId = 35` AND `_x00da_ltima = true` →
      confirm real matching rows come back (Entorno "L02"); separately confirm that
      combining with historical `Entorno = "l01-2016"` rows (all `_x00da_ltima =
      false`) correctly yields `total: 0` (expected, not a bug).
- [x] 8.5 Same list: a string `"true"` against `_x00da_ltima` → confirm `400` from
      schema validation, not 55+ unfiltered rows.
- [x] 8.5b Same list: filter on `Cliente_x002d_LIBSA` (the base lookup column, no
      `LookupId` suffix) → confirm `400` explaining the `LookupId` requirement, not a
      `200` with unfiltered or incorrectly-filtered rows.
- [x] 8.6 Same list: `order_by` over the full (>5000-row) list → confirm the
      threshold `notSupported` error is returned with Graph's detail, not swallowed.
- [x] 8.7 Note as a manual follow-up (not an automated task): after merge, update
      `docker-ag` (172.29.x.x) via `git push origin main` + `./deploy.sh update
      --pull` over SSH, then verify `http://localhost:8001/health` on that host.
      Optionally update the GitHub remote URL to the double-`n`
      `SharepointConnector.git` name.

## 9. Select projection — schema + service (v2.5.0)

- [x] 9.1 Add `select: list[str] | None = None` to `ListItemsSearchByUrlRequest`
      (`max_length=50`, no `min_length` — `[]` ≡ omitted, mirroring `filters`; each
      entry pattern `^[A-Za-z0-9_]+$`).
- [x] 9.2 Service: validate each `select` name at the service layer
      (`_FIELD_NAME_RE` re-check) and against the cached column schema — nonexistent
      column → `GraphAPIError(400)` naming it. Accept BOTH lookup forms (base name
      → display value; `{Base}LookupId` → numeric ID), per design.md decision 9.
- [x] 9.3 Service: when `select` is non-empty, build
      `$expand=fields($select=Col1,Col2,...)` instead of `$expand=fields`
      (percent-encode as needed); omitted/empty → unchanged `$expand=fields`.

## 10. Select projection — tests

- [x] 10.1 Unit tests: `$expand=fields($select=...)` URL construction; omitted and
      `[]` produce plain `$expand=fields` (backward compat).
- [x] 10.2 Unit tests: nonexistent column in `select` → 400 naming it; base lookup
      name accepted; `{Base}LookupId` accepted.
- [x] 10.3 Schema tests: entry with invalid pattern → 422; >50 entries → 422;
      endpoint round trip passing `select` through to the service.
- [x] 10.4 Full suite (`pytest`) green, no regressions.

## 11. Select projection — docs & version

- [x] 11.1 README.md: document `select` in the items:search section (+ example).
- [x] 11.2 ARQUITECTURA.md: extend the items:search subsection (projection,
      validation, lookup asymmetry with filters, system metadata keys note).
- [x] 11.3 doc/CHANGELOG.md: `## v2.5.0` entry; bump `VERSION` to `2.5.0`; update
      version headers in README/ARQUITECTURA/arquitecturasUML.
- [x] 11.4 requirements/SPEC-004: append the select requirement + bitácora entry
      (no separate SPEC-005, per design.md decision 9).

## 12. Select projection — empirical verification (DSL list)

- [x] 12.1 Local Docker rebuild; `/health` reports 2.5.0.
- [x] 12.2 DSL list: `select: ["Title", "Entorno", "_x00da_ltima"]` with a filter →
      inspect the actual keys of `fields` in returned items: only those columns plus
      whatever system keys Graph includes (record the exact system-key set, e.g.
      `@odata.etag`, in SPEC-004/ARQUITECTURA).
- [x] 12.3 DSL list: nonexistent column in `select` → 400 naming it (NOT a 200 with
      the name silently dropped).
- [x] 12.4 DSL list: `select` with base lookup name (`Cliente_x002d_LIBSA`) returns
      display values; with `Cliente_x002d_LIBSALookupId` returns numeric IDs.
- [x] 12.5 Manual post-merge follow-up: push + `deploy.sh update --pull` on
      docker-ag; verify 2.5.0 on that host.
