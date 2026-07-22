## Context

The `/v1/sharepoint` router already builds Graph `$filter` expressions in two places
(`find_list_items_by_field`, `find_list_items_for_upsert` — `app/services/sharepoint.py:229-332`),
both single-purpose (one key field, optional date range) and both treat every value as an
OData string literal. This change needs a general N-field AND filter with typed values, which
means (a) generalizing clause-building into a reusable per-field literal formatter, and
(b) — new — validating each filter against the list's actual column schema, because Graph's
default behavior for a mismatched clause (wrong type, or a boolean compared with `true`/`false`)
is to silently return unfiltered/all rows rather than error. That silent-wrong-answer failure
mode is the core problem this design has to close off.

No existing code calls `GET /sites/{id}/lists/{id}/columns`, so the column-fetch-and-cache
path is new. The existing cache precedent (`SharePointResolver._site_cache`, a plain
unbounded dict) is not reused as-is: site IDs are permanent for a process lifetime, but list
columns can be added/renamed, so this needs an expiry.

## Goals / Non-Goals

**Goals:**
- One request → 0..15 AND-combined field filters, each independently type-checked against
  the list's real schema before touching Graph.
- Correct-by-default OData literal encoding per JSON type, in particular the non-obvious
  boolean → `1`/`0` rule (Graph silently ignores `eq true`/`eq false` on Yes/No columns).
- Bounded, truncating pagination (`top`, `has_more`) and optional `order_by`, with Graph's
  list-view-threshold error surfaced rather than swallowed.
- Reuse `SharePointResolver.resolve_list()` unchanged; reuse `GraphAPIError` for all error
  paths; keep the existing two-layer field-name validation pattern.

**Non-Goals:**
- Cursor-based pagination (no consumer yet; deferred, additive later).
- OR logic, nested groups, or operators beyond equality (`eq`) — AND-of-equalities only,
  matching what the proposal scopes.
- Changing `find_list_items_by_field` / `find_list_items_for_upsert` / their endpoints —
  they keep their current (string-only) behavior; this change does not retrofit them.
- Full OData type coverage for every SharePoint column type (e.g. `personOrGroup`,
  `term`/managed metadata, `hyperlinkOrPicture`) beyond what's needed to map JSON
  `str | bool | int | float` onto the column types the DSL list and similar lists use
  (text/note, boolean, number/currency, lookup-via-`LookupId` suffix, dateTime-as-text).

## Decisions

### 1. Endpoint path: `POST /list/items:search` (plural), not `item:search` — CONFIRMED
The existing `:upsert` precedent uses singular `item:upsert` because it acts on exactly one
item. This endpoint returns a collection, so plural `items:search` better matches REST
convention for a search/list operation — this does introduce the first pluralization
inconsistency on this router (`/list/item`, `/list/item:upsert` vs `/list/items:search`),
knowingly accepted. Alternative considered: `item:search` for naming consistency with
`:upsert` — rejected since the response shape is a collection. **Confirmed with the user.**

### 2. OData literal formatter: pure function, keyed off Python type after Pydantic parsing
A single `_to_odata_literal(value: str | bool | int | float) -> str` helper (service layer)
replaces ad-hoc string-escaping, dispatched on `type(value)` — **checking `bool` before
`int`/`float`**, since `bool` is an `int` subclass in Python and Pydantic v2's smart-mode
union preserves the caller's original JSON type (`true` stays `bool`, `35` stays `int`):
- `bool` → `"1"` / `"0"` (never `"true"`/`"false"` — verified empirically against the DSL
  list that Graph silently ignores the latter on Yes/No columns)
- `str` → `"'" + value.replace("'", "''") + "'"` (same escaping as existing code)
- `int` / `float` → `str(value)` unquoted
This function is independent of column type — the column-schema validation (decision 4) is
what guarantees the JSON type is appropriate for the column *before* this formatter runs, so
the formatter itself doesn't need to know SharePoint column types.

### 3. Column schema cache: in-process dict with a stored timestamp, TTL-checked on read — CONFIRMED
No new dependency: reuse the same "plain dict on the service/resolver object" pattern as
`SharePointResolver._site_cache`, but store `(columns, fetched_at)` per `(site_id, list_id)`
key and treat an entry as stale after a fixed TTL, **hardcoded at 300s, no config surface**
(confirmed with the user — add a config knob later only if it becomes a real operational
need). On a cache miss or stale entry, re-fetch via
`GET /sites/{site_id}/lists/{list_id}/columns` and overwrite.
Alternatives considered: (a) no TTL, matching the site cache exactly — rejected because
column additions/renames are a realistic mid-process event, unlike site IDs; (b) pull in
`cachetools` for a proper `TTLCache` — rejected as unnecessary weight for a single call site;
a manual timestamp check is a few lines and keeps the zero-new-dependency footprint.
This cache lives on `SharePointService` (or a small new collaborator it owns), not on
`SharePointResolver`, since it's about list *contents* (columns), which the resolver's own
docstring already says are deliberately not cached there.

### 4. Schema validation: reject-only, no implicit coercion
For each filter `{field, value}`:
1. Look up `field` in the cached column list by internal `name`.
   - Direct match on a **non-`lookup`** facet → proceed to step 2.
   - Direct match on a column whose facet **is** `lookup` (i.e. `field` is the *base*
     lookup column name, no `LookupId` suffix) → **`400`, explicitly rejected** — see
     the lookup special case below; this is a distinct, deliberate rejection, not a
     "field not found."
   - No direct match at all, and not a `LookupId`-suffixed match (see below) → `400`,
     naming the offending field as nonexistent.
2. Map the column's type facet to an expected JSON type: `boolean`→`bool`, `number`/
   `currency`→`int | float`, everything else usable here (`text`, `note`, `choice`,
   `dateTime`, `personOrGroup`) → `str`. (`lookup` never reaches this step — it's
   handled/rejected in step 1.)
3. `type(value)` (post-Pydantic, `bool` checked before `int`) must match the expected type
   family exactly → mismatch is `400`, explaining the expected type. **No coercion**
   (`"true"` staying a rejected string, `"35"` staying a rejected string against a numeric
   column) — the proposal left this open; reject-only was chosen because coercion means
   guessing caller intent (is `"35"` really meant as the number 35?) and this connector's
   existing philosophy (see `_FIELD_NAME_RE`, `extra="forbid"`) is fail-fast/explicit over
   lenient. Coercion can be added later without breaking existing callers if a real need
   shows up.
- **Lookup-column special case — CONFIRMED with the user**: SharePoint lookup columns are
  exposed to callers as a synthetic `{InternalName}LookupId` field (see requirement's own
  example, `Cliente_x002d_LIBSALookupId`) that does **not** appear as its own entry in
  `GET .../columns` — only the base lookup column (`Cliente_x002d_LIBSA`, type facet
  `lookup`) does. So: if `field` isn't found directly but ends in `LookupId`, strip the
  suffix and retry the lookup; if that matches a column with a `lookup` facet, the expected
  JSON type is `int` (lookup IDs are numeric) — this is required to make the requirement's
  own worked example (`Cliente_x002d_LIBSALookupId = 35`) validate and filter correctly.
  Conversely, if `field` matches a `lookup`-type column **directly** (no `LookupId` suffix
  — the caller tried to filter on the base lookup column itself), the request is rejected
  with `400` explaining that lookup columns must be filtered via their `{Base}LookupId`
  field. This was explicitly decided over silently treating it as a string: Graph does not
  reliably filter lookup base columns by equality the way it does text/number/boolean
  columns, so accepting the request would risk reintroducing the exact silent-wrong-answer
  failure mode (requirement 4 / this design) exists to eliminate.

### 5. Pagination/truncation: loop like `_get_all`, but stop and flag at `top`
New `_get_bounded(url, top, extra_headers) -> tuple[list[dict], bool]` mirroring `_get_all`'s
next-link loop, but breaking out once `len(items) >= top` (slicing to exactly `top`) and
returning `has_more=True` if either the break happened before `next_url` ran out, or the last
page fetched had more rows than needed. This keeps `_get_all` untouched for its existing
callers and adds a sibling rather than adding a branchy `top: int | None` parameter to it.

### 6. `order_by` and the list-view-threshold error: pass through, don't pre-empt
No client-side detection of "this list is probably too big" — Graph's `notSupported` error
(with its message about the list view threshold) is caught the same way every other Graph
error is (`raise_from_httpx` → `GraphAPIError`) and surfaced with its detail intact. The
`design` doesn't add special retry/backoff logic; the response documents the recommended
workaround (sort by `Created`/`Modified`/`ID`) rather than the code enforcing it.

### 7. Request validation: `extra="forbid"` + existing double-layer field-name check
`ListItemsSearchByUrlRequest` (schema layer) sets `model_config = ConfigDict(extra="forbid")`
so `filter_by` or any unrecognized key is a `422`. Each `SearchFilter.field` keeps the
existing `pattern=r"^[A-Za-z0-9_]+$"` (422 fast-path), and the service layer re-validates
with the same `_FIELD_NAME_RE` before interpolating into the filter string (defense in
depth, consistent with `find_list_items_by_field`/`find_list_items_for_upsert`). Note the
`LookupId`-suffix case in decision 4 still passes the same regex unchanged (it's alphanumeric
+ underscore), so no regex change is needed for it.

### 8. Response shape
`{ total, items: [{id, webUrl, fields}], has_more, site_id, list_id }` — `total` is
`len(items)` *after* truncation (matches proposal), not the count of all Graph-side matches
(which is never fully known when `has_more` is true, by design — no full-count query is made).

### 9. `select` projection (v2.5.0 iteration) — schema-validated, `[]` ≡ omitted
Added after v2.4.0 shipped, owner-approved 2026-07-22. Optional `select: list[str]`
(`max_length=50`, no `min_length`) projecting each item's `fields` to only the named
columns via `$expand=fields($select=Col1,Col2)`; omitted (or `[]`, mirroring the
`filters: []` decision) keeps today's `$expand=fields` — fully backward compatible.

- **Validation mirrors filters**: names pass `[A-Za-z0-9_]+` in both layers (schema 422 +
  service 400) and are checked against the cached column schema — a nonexistent column is
  a `400` naming it, because Graph *silently ignores* unknown names in `$select` (item
  returned, no error), which would hide caller typos: the same silent-failure mode
  requirement 4 eliminated for filters.
- **Lookup asymmetry with filters — deliberate**: in `select`, BOTH forms are valid —
  the base lookup column name (e.g. `Cliente_x002d_LIBSA`, Graph returns its display
  value) and the `{Base}LookupId` synthetic field (returns the numeric ID). This differs
  from filters, where the base name is rejected with `400`: *filtering* on a base lookup
  column is unreliable in Graph (silently wrong results), but *selecting* it works and is
  useful. The design flags this asymmetry explicitly so it reads as intentional, not an
  oversight.
- **System metadata keys**: Graph may still include bookkeeping keys (e.g. `@odata.etag`)
  inside `fields` even under `$select`. The contract is "only the requested columns plus
  Graph system metadata"; the exact key set is confirmed during empirical verification and
  documented, not guessed.
- **No response-shape change**: `items[].fields` simply has fewer keys; `total`/`has_more`
  semantics untouched. Version bumps to 2.5.0; SPEC-004 is extended (same endpoint) rather
  than opening a SPEC-005.

## Risks / Trade-offs

- **[Risk]** TTL cache can serve stale columns for up to the TTL window after a real schema
  change (new/renamed column) → **Mitigation**: short default TTL (300s); a stale-negative
  (column added, not yet visible) only affects that window and self-heals; document the TTL
  in README/ARQUITECTURA so operators know why a just-added column might 400 briefly.
- **[Risk]** Reject-only type validation may surprise callers used to loose typing (e.g.
  webhook sources that always send strings) → **Mitigation**: this is intentional per
  decision 4; error message names the expected type so the fix is a one-line caller change,
  and it's the same fail-fast philosophy already used for field-name validation.
- **[Risk]** Lookup-column special-casing (decision 4) is inferred from one worked example,
  not from exhaustively testing every SharePoint column type Graph exposes → **Mitigation**:
  the mandatory empirical verification task (against the real DSL list) explicitly exercises
  this exact field, and the base-lookup-column case is now an explicit `400` rather than a
  best-effort guess; if the `LookupId` path doesn't validate cleanly, this decision gets
  revisited before merge, not after.
- **[Risk]** `order_by` against a large list can hard-fail with `notSupported` and there's no
  workaround inside this endpoint (no auto-retry without sort) → **Mitigation**: documented
  limitation, recommended indexed-column workaround, matches requirement's explicit
  acceptance of this as a known limitation rather than something to engineer around.
- **[Trade-off]** No cursor pagination means a caller who needs "all 8000 rows" must raise
  `top` (max 5000) and still may get `has_more: true` with no way to fetch the rest via this
  endpoint yet → accepted per proposal (no real consumer today; additive later).

## Migration Plan

- Purely additive: new route, new schemas, new service methods, one new cache. No changes to
  existing endpoints, schemas, or stored data — nothing to migrate or backfill.
- Deploy as a minor version bump (2.3.0 → 2.4.0) in `VERSION`, `doc/CHANGELOG.md`.
  The `select` iteration (decision 9) is a further minor bump (2.4.0 → 2.5.0), equally
  additive: no migration, same rollout/rollback procedure.
- Rollout is manual and environment-specific (see proposal's Impact section): local Docker
  Desktop rebuild, then a separate commit+push+`deploy.sh update --pull` cycle on `docker-ag`.
  Rollback is reverting the commit and redeploying the previous image on `docker-ag`; no data
  migration means rollback has no cleanup step.

## Open Questions

_None outstanding._ All three questions raised during review are resolved and confirmed
with the user: endpoint naming (`items:search`, decision 1), column-cache TTL (hardcoded
300s, decision 3), and bare lookup-column filtering (explicit `400`, decision 4).
