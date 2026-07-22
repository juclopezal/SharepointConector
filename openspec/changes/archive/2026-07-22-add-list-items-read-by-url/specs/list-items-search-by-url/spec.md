## ADDED Requirements

### Requirement: Multi-field AND search by list URL
The system SHALL expose `POST /v1/sharepoint/list/items:search`, accepting a
`sharepoint_url` and an optional `filters` array of up to 15 conditions
(`{ "field": string, "value": string | boolean | integer | number }`), and SHALL
combine all provided conditions with logical AND when querying Microsoft Graph.
When `filters` is omitted, or is present but an empty array, the system SHALL
treat both cases identically and return the first `top` items of the list without
applying any filter.

#### Scenario: Single filter matches a subset of items
- **WHEN** the request has one filter `{ "field": "Entorno", "value": "L02" }`
- **THEN** the response contains only items whose `Entorno` field equals `"L02"`

#### Scenario: Multiple filters are combined with AND
- **WHEN** the request has filters `{ "field": "Entorno", "value": "L02" }` and
  `{ "field": "_x00da_ltima", "value": true }`
- **THEN** the response contains only items matching both conditions simultaneously,
  never items matching only one of them

#### Scenario: No filters returns the first page unfiltered
- **WHEN** the request omits `filters` and sets `top: 100`
- **THEN** the response contains up to 100 items from the list with no `$filter` applied

#### Scenario: Explicit empty filters array behaves the same as omitting it
- **WHEN** the request includes `"filters": []` (an explicit empty array, not an
  omitted key) and sets `top: 100`
- **THEN** the response is `200`, contains up to 100 items from the list with no
  `$filter` applied, and is treated identically to omitting `filters` entirely —
  it is NOT rejected with `422`

#### Scenario: A filter value matching no row returns an empty result, not an error
- **WHEN** a filter's `value` is well-typed for its column but matches zero rows
- **THEN** the response is `200` with `total: 0` and an empty `items` array

### Requirement: Typed filter values translate to correct OData literals
The system SHALL accept `value` as a JSON string, boolean, integer, or floating-point
number, and SHALL translate each to the OData literal that Microsoft Graph evaluates
correctly for the corresponding SharePoint column, without requiring the caller to
know OData syntax. In particular, boolean values SHALL be encoded as the unquoted
literals `1` or `0` — never `true`/`false` in either quoted or unquoted form — because
Microsoft Graph silently ignores `fields/{column} eq true` and `fields/{column} eq
'true'` clauses against SharePoint Yes/No columns, returning all rows as if the
condition were absent, whereas `eq 1` / `eq 0` filters correctly.

#### Scenario: Boolean true value filters correctly
- **WHEN** a filter is `{ "field": "_x00da_ltima", "value": true }`
- **THEN** every item in the response has `_x00da_ltima: true` in its `fields`, and no
  item with `_x00da_ltima: false` is returned

#### Scenario: Boolean false value filters correctly
- **WHEN** a filter is `{ "field": "_x00da_ltima", "value": false }`
- **THEN** every item in the response has `_x00da_ltima: false` in its `fields`

#### Scenario: String value is quoted with internal quotes doubled
- **WHEN** a filter is `{ "field": "Title", "value": "O'Hara" }`
- **THEN** the Graph `$filter` sent contains the literal `'O''Hara'` for that clause

#### Scenario: Numeric value is emitted unquoted
- **WHEN** a filter is `{ "field": "Cliente_x002d_LIBSALookupId", "value": 35 }`
- **THEN** the Graph `$filter` sent contains the unquoted literal `35` for that clause

#### Scenario: Integer true is distinguished from boolean true
- **WHEN** a filter is `{ "field": "Cliente_x002d_LIBSALookupId", "value": 1 }`
  (a JSON integer, not a JSON boolean)
- **THEN** the value is treated as the number `1`, not as a boolean, and validated
  against the column's expected type as an integer

#### Scenario: Unsupported value type is rejected before calling Graph
- **WHEN** a filter's `value` is `null`, a JSON object, or a JSON array
- **THEN** the request is rejected with `422` and never reaches Microsoft Graph

### Requirement: Filters are validated against the list's real column schema
Before building the Graph `$filter` expression, the system SHALL validate every
filter's `field` against the target list's actual columns (fetched via Microsoft
Graph and cached per `site_id` + `list_id`). A filter referencing a field that does
not exist in the list, or whose `value`'s JSON type does not correspond to that
column's type, SHALL be rejected explicitly rather than silently dropped or passed
through to Graph unfiltered.

#### Scenario: Nonexistent field is rejected
- **WHEN** a filter's `field` does not match any column on the resolved list
- **THEN** the response is `400`, and the error names the offending field

#### Scenario: Type mismatch against a Yes/No column is rejected
- **WHEN** a filter targets a Yes/No (boolean) column with a string value, e.g.
  `{ "field": "_x00da_ltima", "value": "true" }`
- **THEN** the response is `400` explaining the expected type, and the request never
  reaches Graph as an unfiltered or partially-filtered query

#### Scenario: Type mismatch against a numeric column is rejected
- **WHEN** a filter targets a numeric column with a string value, e.g.
  `{ "field": "Cliente_x002d_LIBSALookupId", "value": "35" }`
- **THEN** the response is `400` explaining the expected type

#### Scenario: Lookup column addressed via its LookupId field is recognized
- **WHEN** a filter's `field` is `{BaseColumnName}LookupId` (e.g.
  `Cliente_x002d_LIBSALookupId`) and no column by that exact name exists, but a
  lookup-type column named `{BaseColumnName}` does exist
- **THEN** the field is recognized as valid, its expected JSON type is treated as
  numeric, and a well-typed integer `value` is accepted

#### Scenario: Filtering directly on a base lookup column is rejected
- **WHEN** a filter's `field` matches an existing column whose type is a lookup
  column, addressed by its base name directly (i.e. NOT the `{BaseColumnName}LookupId`
  synthetic field — e.g. `Cliente_x002d_LIBSA` instead of
  `Cliente_x002d_LIBSALookupId`)
- **THEN** the response is `400`, explaining that lookup columns must be filtered via
  their `LookupId` field, rather than being accepted and silently mishandled by Graph

#### Scenario: Column schema is cached across requests
- **WHEN** two search requests target the same `site_id` + `list_id` within the
  cache's freshness window
- **THEN** the second request SHALL NOT trigger a second `GET .../lists/{id}/columns`
  call to Microsoft Graph

### Requirement: Optional result ordering
The system SHALL accept an optional `order_by` object (`{ "field": string,
"direction": "asc" | "desc" }`, `direction` defaulting to `"asc"`) and, when present,
SHALL translate it to a Graph `$orderby` parameter. When Microsoft Graph rejects the
sort because the list exceeds its list-view threshold without an indexing filter,
the system SHALL propagate the error (including Graph's detail) rather than
suppressing it or silently returning unsorted results.

#### Scenario: Ascending sort by default
- **WHEN** the request includes `order_by: { "field": "Created" }`
- **THEN** the Graph request includes `$orderby=fields/Created asc`

#### Scenario: Explicit descending sort
- **WHEN** the request includes `order_by: { "field": "Created", "direction": "desc" }`
- **THEN** the Graph request includes `$orderby=fields/Created desc`

#### Scenario: Sorting a list past the view threshold surfaces Graph's error
- **WHEN** `order_by` is applied to a list whose size exceeds SharePoint's list view
  threshold and no filter narrows the query enough to use an index
- **THEN** the response reflects Microsoft Graph's `notSupported` error with its
  detail message, rather than returning `200` with incorrect or unsorted data

### Requirement: Bounded and truncated result set
The system SHALL accept an optional `top` parameter (integer, 1–5000, default 100)
bounding the number of items returned. Because Microsoft Graph treats `$top` as a
page size rather than a total limit, the system SHALL follow `@odata.nextLink`
pagination internally until either `top` items have been collected or no matching
items remain, then truncate any excess in the response. The response SHALL include a
`has_more` boolean indicating whether matching items beyond `top` were left out.

#### Scenario: Fewer matches than top
- **WHEN** a filter matches fewer items than the requested `top`
- **THEN** the response contains all matching items, `total` equals that count, and
  `has_more` is `false`

#### Scenario: More matches than top
- **WHEN** a filter matches more items than the requested `top`
- **THEN** the response contains exactly `top` items, `total` equals `top`, and
  `has_more` is `true`

#### Scenario: top below the allowed minimum is rejected
- **WHEN** `top` is `0` or negative
- **THEN** the response is `422`

#### Scenario: top above the allowed maximum is rejected
- **WHEN** `top` is greater than `5000`
- **THEN** the response is `422`

### Requirement: Strict request validation
The system SHALL reject any request body field not defined by the search request
schema (including a `filter_by` field carried over from other endpoints) with `422`,
rather than ignoring it silently. Each filter's `field` SHALL match the internal
SharePoint column name pattern (letters, digits, underscore only); a `field` outside
that pattern SHALL be rejected at the schema layer, and independently re-validated
at the service layer before being interpolated into the Graph filter expression.

#### Scenario: Unknown top-level field is rejected
- **WHEN** the request body includes a field not defined by the schema, e.g.
  `filter_by`
- **THEN** the response is `422`, and the unknown field is never interpreted as a
  filter or otherwise acted upon

#### Scenario: Filter field name outside the allowed pattern is rejected
- **WHEN** a filter's `field` contains characters other than letters, digits, or
  underscore (e.g. `"Title eq 'x'"` or `"fields/Other"`)
- **THEN** the response is `422`

#### Scenario: Filters array above the allowed size is rejected
- **WHEN** `filters` contains more than `15` conditions
- **THEN** the response is `422`
- **NOTE**: an empty `filters` array (`0` conditions) is NOT rejected — see the
  "Explicit empty filters array" scenario above; only exceeding the maximum is a
  `422`

### Requirement: Field projection via select
The system SHALL accept an optional `select` array of up to 50 column internal
names. When present and non-empty, each returned item's `fields` SHALL contain
only the requested columns (plus any system metadata keys Microsoft Graph always
includes, such as `@odata.etag`), translated internally to
`$expand=fields($select=...)`. When `select` is omitted or an empty array, the
system SHALL return all fields, identical to the pre-projection behavior. Each
name SHALL be validated against the list's real column schema before calling
Graph: a nonexistent column in `select` SHALL be rejected with an explicit error
rather than silently ignored (Microsoft Graph's default for unknown `$select`
names). Unlike filtering, a base lookup column name IS valid in `select` (it
projects the display value), in addition to its `{Base}LookupId` form (numeric ID).

#### Scenario: Projection returns only the requested columns
- **WHEN** the request includes `select: ["Title", "Entorno"]`
- **THEN** every returned item's `fields` contains `Title` and `Entorno` (when set on
  the item) and no other list columns, beyond Graph system metadata keys

#### Scenario: Omitted or empty select returns all fields
- **WHEN** the request omits `select` or sends `"select": []`
- **THEN** each item's `fields` contains all columns, identical to the behavior
  before this capability existed

#### Scenario: Nonexistent column in select is rejected
- **WHEN** `select` includes a name that matches no column on the resolved list
- **THEN** the response is `400` naming the offending column, and the request never
  reaches Graph with a silently-dropped `$select` name

#### Scenario: Both lookup forms are selectable
- **WHEN** `select` includes a base lookup column name (e.g. `Cliente_x002d_LIBSA`)
  or its `{Base}LookupId` form
- **THEN** both are accepted — the base name projects the display value and the
  `LookupId` form the numeric ID — in contrast with filters, where the base name is
  rejected

#### Scenario: Invalid select name pattern is rejected at the schema layer
- **WHEN** a `select` entry contains characters other than letters, digits, or
  underscore
- **THEN** the response is `422`

#### Scenario: Select above the allowed size is rejected
- **WHEN** `select` contains more than `50` entries
- **THEN** the response is `422`

### Requirement: Search response shape
The system SHALL respond `200` with `{ total, items, has_more, site_id, list_id }`,
where `items` is an array of `{ id, webUrl, fields }` for each returned item, and
`total` is the count of items actually returned (after truncation), not the total
count of matches on the server.

#### Scenario: Response includes resolved identifiers
- **WHEN** a search request resolves successfully
- **THEN** the response includes the `site_id` and `list_id` that `sharepoint_url`
  resolved to
