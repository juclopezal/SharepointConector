## Why

Callers can create (`POST /list/item`), update-by-single-field (`PATCH /list/item`), and
upsert-by-key+period (`POST /list/item:upsert`) SharePoint list items by URL, but there is no
way to *read/search* items by more than one field at once without hand-rolling raw Graph
`$filter` syntax against `/v1/graph/...`. Consumers need to query lists (e.g. the DSL list)
by several business fields combined with AND — environment + "latest" flag + client lookup —
and get back the matching rows, not just a single record. Today's closest primitive
(`find_list_items_by_field`) only supports one field, and its OData literal handling is
naive (everything is quoted as a string), which silently produces wrong results for
boolean SharePoint columns (Graph ignores `eq true`/`eq false` clauses without erroring).

## What Changes

- New endpoint `POST /v1/sharepoint/list/items:search`: resolves a list by URL, applies 0–15
  AND-combined field filters, optional sort, and a bounded `top`, returning matching items.
- New typed `value` contract in filters (`str | bool | int | float`) with correct-by-default
  OData literal translation — notably booleans render as `1`/`0` (never `true`/`false`),
  the only literal form that actually filters SharePoint Yes/No columns via Graph.
- New schema-validation step: before building `$filter`, the list's column definitions are
  fetched (new Graph call, `GET /sites/{id}/lists/{id}/columns`) and cached per
  `site_id+list_id`, so an unknown field or a JSON-type/column-type mismatch fails fast with
  a `400` instead of Graph silently dropping the clause.
- New optional `order_by` (`$orderby`), with the list-view-threshold `notSupported` Graph
  error surfaced verbatim (with guidance to sort by an indexed column) rather than swallowed.
- New pagination-truncation contract: `top` bounds the *returned* count (not the Graph page
  size); the service keeps following `@odata.nextLink` until `top` items are collected or
  results are exhausted, and reports `has_more` when rows were left out.
- Request model uses `extra="forbid"`: no legacy `filter_by`-style field carries over from
  the update/upsert endpoints, and any unrecognized body field is a `422`, not a silent no-op.
- **Second iteration (v2.5.0), owner-approved 2026-07-22**: optional `select` field —
  array of column internal names projecting each returned item's `fields` to only those
  columns (Graph `$expand=fields($select=...)`). Omitted → all fields, fully backward
  compatible. Names are schema-validated like filter fields (Graph silently ignores unknown
  names in `$select`, hiding caller typos — same silent-failure mode this change eliminates).
- **BREAKING**: none — this is a net-new endpoint; no existing endpoint changes shape, and
  the `select` iteration is additive over v2.4.0.

## Capabilities

### New Capabilities
- `list-items-search-by-url`: multi-field AND-filtered search of SharePoint list items by
  list URL, including schema-validated filters, typed OData literal translation, optional
  ordering, and bounded/truncated pagination with `has_more`.

### Modified Capabilities
_None — existing `:upsert`/update/create endpoints and their delta specs are unaffected;
this change only adds a new read/search capability alongside them._

## Impact

- **New route**: `app/api/v1/endpoints/sharepoint.py` — `POST /list/items:search` on the
  existing `sharepoint` router (prefix `/sharepoint`).
- **New schemas**: `app/schemas/sharepoint.py` — `SearchFilter`, `OrderBy`,
  `ListItemsSearchByUrlRequest`, `ListItemsSearchByUrlResponse` (+ item shape).
- **New service code**: `app/services/sharepoint.py` — OData literal translation helper,
  schema-validated multi-clause filter builder, `$orderby` support, truncating pagination,
  and a new `get_list_columns()` Graph call.
- **New/changed cache**: a column-schema cache keyed by `site_id+list_id`, location and TTL
  mechanism to be decided in `design.md` (existing `SharePointResolver._site_cache` is an
  unbounded dict with no TTL — precedent to reconcile with, not necessarily copy).
- **No changes** to `app/services/resolver.py` (`resolve_list()` reused as-is) or to any
  existing endpoint/schema.
- **Docs**: README.md, ARQUITECTURA.md, arquitecturasUML.md, doc/CHANGELOG.md, VERSION
  (minor bump from 2.3.0), new `requirements/SPEC-004_*.md`. The `select` iteration bumps
  to 2.5.0 and extends SPEC-004 (same endpoint, no separate SPEC-005).
- **Deployment**: manual post-merge step in both local Docker Desktop and the remote
  `docker-ag` host — not automated by this change's tasks.
