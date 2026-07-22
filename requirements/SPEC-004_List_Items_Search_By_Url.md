# SPEC-004: Búsqueda multi-campo de ítems de lista por URL

> **Estado:** ✅ Implementada (v2.4.0). Todos los criterios de aceptación de
> código y tests cubiertos; verificación empírica contra la lista DSL registrada
> en la **Bitácora de IA** al final.

## 1. Contexto (El Problema)

El conector permite crear (`POST /list/item`), actualizar por un campo único
(`PATCH /list/item`) y hacer upsert por clave+periodo (`POST /list/item:upsert`)
sobre ítems de lista identificados por su URL, pero **no leer/buscar** ítems por
**varios campos a la vez**: el consumidor tendría que construir a mano un
`$filter` OData contra la API de bajo nivel `/v1/graph/...`.

Además, el primitivo de búsqueda existente (`find_list_items_by_field`) trata
todo valor como string. Con columnas **Sí/No** eso produce resultados
**silenciosamente incorrectos**: Graph **ignora sin error** las cláusulas
`fields/{col} eq true|false` (y también `eq 'true'` como string) sobre esas
columnas y devuelve TODAS las filas como si la condición no existiera. Solo
`eq 1`/`eq 0` filtra de verdad (verificado empíricamente contra la lista DSL el
2026-07-22). Ese modo de fallo — cláusula descartada en silencio, respuesta
plausible pero errónea — es exactamente lo que esta spec debe eliminar.

> **Alcance: genérico, no atado a una lista.** Igual que SPEC-003, nada del
> dominio viaja quemado en el código: la lista, los campos y los tipos van en el
> payload. La lista DSL es solo el caso motivador y el banco de pruebas.

## 2. Propuesta (La Solución)

Nuevo endpoint `POST /v1/sharepoint/list/items:search`:

```json
{
  "sharepoint_url": "https://host.sharepoint.com/Oper/Lists/DSL/Allitemsg.aspx",
  "filters": [
    { "field": "Entorno", "value": "L02" },
    { "field": "_x00da_ltima", "value": true }
  ],
  "order_by": { "field": "Created", "direction": "desc" },
  "top": 100
}
```

- **`filters`** (opcional, hasta 15): condiciones de igualdad combinadas con
  **AND**. Omitido o `[]` → primeros `top` ítems sin filtrar. Es la **única**
  forma de filtrado (decisión 2026-07-22 con el propietario: no se arrastra
  `filter_by`, que solo existe en el update con otra semántica); el request usa
  `extra="forbid"` → campo desconocido = `422`.
- **Valores tipados** (`str | bool | int | float`; null/objeto/array → `422`).
  El caller envía booleanos JSON puros (`true`, sin comillas); la traducción al
  literal OData es interna: string → `'texto'` (comillas duplicadas), boolean →
  **`1`/`0`** (nunca `true`/`false`), número → literal sin comillas.
- **Validación contra el esquema real** (`GET .../columns`, caché por
  `(site_id, list_id)` con TTL 300 s): campo inexistente → `400` nombrándolo;
  tipo JSON ≠ tipo de columna → `400` con el tipo esperado (**sin coerción**);
  columna lookup por su nombre base → `400` (usar `{Columna}LookupId`, entero).
  Valor bien tipado sin coincidencias → `200` con `total: 0` (conducto regular).
- **`order_by`** → `$orderby`. Limitación conocida: por encima del umbral de
  vista (~5000 filas) Graph responde `notSupported` con columnas no indexadas;
  se propaga con detalle (recomendar `Created`, `Modified`, `ID`).
- **`top`** (1–5000, default 100): Graph trata `$top` como tamaño de página; el
  conector sigue `@odata.nextLink` hasta reunir `top`, trunca y devuelve
  `has_more`. Paginación por cursor: pospuesta conscientemente.
- Garantías existentes: campo `[A-Za-z0-9_]+` en doble capa (`422`/`400`),
  percent-encoding de la expresión completa, cabecera
  `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`.
- **Respuesta:** `200` con
  `{ total, items[{id, webUrl, fields}], has_more, site_id, list_id }`
  (`total` = ítems devueltos tras el truncado).

## 3. Criterios de Aceptación

1. Multi-campo AND: `Entorno="L02"` + `_x00da_ltima=true` devuelve solo filas
   que cumplen ambas condiciones.
2. Booleano JSON puro filtra correctamente (literal interno `1`/`0`); todos los
   ítems devueltos traen el valor booleano pedido en `fields`.
3. Campo inexistente o tipo que no casa con la columna → `400` explícito con el
   campo/tipo ofendido; jamás una cláusula ignorada en silencio.
4. String `"true"` contra columna Sí/No → `400` (no N filas sin filtrar).
5. Columna lookup: `{Base}LookupId` + entero aceptado; nombre base directo → `400`.
6. `filter_by` u otro campo desconocido en el body → `422`; >15 filtros → `422`;
   `top` fuera de 1–5000 → `422`; `value` null/objeto/array → `422`.
7. `filters` omitido o `[]` → `200` con los primeros `top` ítems sin filtrar.
8. Truncado por `top` con `has_more` correcto siguiendo `@odata.nextLink`.
9. `order_by` genera `$orderby`; el error de umbral de Graph se propaga con detalle.
10. Esquema de columnas cacheado: segunda búsqueda sobre la misma lista dentro
    del TTL no repite `GET .../columns`.
11. Suite completa de `pytest` en verde y verificación empírica contra la lista
    DSL real (valores de `fields` inspeccionados, no solo `total`).

## 4. Bitácora de IA / Historial de Implementación

### 2026-07-22 — Agente LLM (Claude Code)

Implementación completa de la spec sobre el cambio OpenSpec
`add-list-items-read-by-url` (artefactos en `openspec/changes/`).

**Archivos modificados:**
- `app/services/sharepoint.py` — helpers puros `_to_odata_literal()` (bool
  comprobado ANTES que int: `bool` es subclase de `int`),
  `_search_field_expected_type()` (resolución de columna con el caso lookup:
  directo → `400`, sufijo `LookupId` → entero) y `_check_search_value_type()`
  (sin coerción); métodos `get_list_columns()` (caché dict `(site_id, list_id)`
  → `(columns, fetched_at)`, TTL fijo 300 s vía `time.monotonic()`, sin
  dependencia nueva), `_get_bounded()` (variante acotada de `_get_all` que corta
  en `top` y calcula `has_more`) y `search_list_items()`.
- `app/schemas/sharepoint.py` — `SearchFilter`, `OrderBy`,
  `ListItemsSearchByUrlRequest` (`extra="forbid"`, `filters` con `max_length=15`
  y sin `min_length`: `[]` explícito equivale a omitirlo),
  `ListItemsSearchByUrlItem`, `ListItemsSearchByUrlResponse`.
- `app/api/v1/endpoints/sharepoint.py` — endpoint
  `POST /v1/sharepoint/list/items:search` (plural `items`, decisión consciente
  frente al singular `item:upsert`: la respuesta es una colección).
- `tests/test_sharepoint_service.py` / `tests/test_sharepoint_endpoints.py` —
  30 tests nuevos; suite completa 114/114 en verde.
- `VERSION` (2.4.0), `README.md`, `ARQUITECTURA.md`, `arquitecturasUML.md`,
  `doc/CHANGELOG.md`.

**Decisiones clave (confirmadas con el propietario el 2026-07-22):**
- Validación **reject-only** (sin coerción de `"35"`/`"true"`): fail-fast
  coherente con `extra="forbid"` y la validación de nombres de campo.
- Columna lookup base sin sufijo → `400` con guía (no best-effort como string):
  Graph no filtra de forma fiable la columna base y aceptar reintroduciría el
  fallo silencioso.
- TTL del caché de columnas fijo en 300 s (sin knob de configuración).
- Nombre `items:search` (plural) pese a la inconsistencia con `item:upsert`.

**Sin desviaciones** respecto al diseño (`openspec/changes/add-list-items-read-by-url/design.md`).

**Verificación empírica (2026-07-22, contra la lista DSL real, Docker local
`localhost:8001`, v2.4.0 confirmada en `/health`):**
- `Entorno = "no-existe-xyz"` → `total: 0`. ✓
- `_x00da_ltima = true` (solo booleano, `top: 5000`) → 3.838 ítems, **todos**
  con `_x00da_ltima: true` en `fields` (la cláusula filtra; con el literal
  `true` de Graph habría devuelto también los `false`). ✓
- `Cliente_x002d_LIBSALookupId = 35` AND `_x00da_ltima = true` → filas reales
  `Entorno "L02"` con LookupId 35 y ultima true; el combinado con las filas
  históricas `Entorno = "l01-2016"` (todas `_x00da_ltima = false`) da
  `total: 0`, correcto. ✓
- String `"true"` contra la columna Sí/No → `400` "la columna espera boolean". ✓
- `Cliente_x002d_LIBSA` (lookup base, sin sufijo) → `400` con guía hacia
  `Cliente_x002d_LIBSALookupId`. ✓
- Campo inexistente → `400` nombrándolo; `filter_by` en el body → `422`. ✓
- `order_by` sobre la lista completa (>5000 filas, columna no indexada) →
  error de Graph propagado con detalle íntegro (`code: notSupported`,
  "supera el umbral de vista de lista", innerError con request-id). ✓
