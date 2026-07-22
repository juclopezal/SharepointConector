## v2.4.0 — 2026-07-22

### Feature: Búsqueda multi-campo de ítems de lista (SPEC-004)

**Contexto:** El conector permitía crear, actualizar (por un campo único) y hacer upsert (clave + periodo) de ítems por URL, pero no **leer/buscar** por varios campos a la vez sin construir a mano un `$filter` OData contra `/v1/graph/...`. Además, el primitivo más cercano (`find_list_items_by_field`) trata todo valor como string, lo que produce resultados **silenciosamente incorrectos** con columnas Sí/No: Graph ignora sin error las cláusulas `eq true|false` (y `eq 'true'`) sobre esas columnas y devuelve todas las filas como si el filtro no existiera.

**Solución:** Nuevo `POST /v1/sharepoint/list/items:search` — hasta 15 condiciones de igualdad combinadas con **AND**, valores **tipados** y validación previa contra el esquema real de la lista.

```json
{
  "sharepoint_url": "https://host.sharepoint.com/Oper/Lists/DSL/Allitemsg.aspx",
  "filters": [
    { "field": "Entorno", "value": "L02" },
    { "field": "_x00da_ltima", "value": true }
  ],
  "top": 100
}
```

- **Booleano → `1`/`0`, nunca `true`/`false`** (verificado empíricamente contra la lista DSL, 2026-07-22): el caller envía booleanos JSON puros y el conector los traduce internamente al único literal OData que filtra de verdad las columnas Sí/No. String → `'texto'` con comillas duplicadas; número → literal sin comillas.
- **Validación contra el esquema de columnas** (`GET .../columns`, caché por `(site_id, list_id)` con TTL 300 s): campo inexistente → `400` nombrándolo; tipo JSON ≠ tipo de columna → `400` con el tipo esperado (sin coerción); columna lookup por su nombre base → `400` (debe usarse `{Columna}LookupId`, tipo entero). Valor bien tipado sin coincidencias → `200` con `total: 0`.
- **`extra="forbid"`**: cualquier campo desconocido del body (p. ej. `filter_by`) → `422`, nunca ignorado.
- **`order_by`** (`$orderby`): en listas por encima del umbral de vista (~5000 filas) Graph responde `notSupported` al ordenar por columnas no indexadas; el error se propaga con detalle (usar `Created`, `Modified`, `ID`).
- **`top`** (1–5000, default 100): Graph trata `$top` como tamaño de página; el conector sigue `@odata.nextLink` hasta reunir `top` ítems, trunca el excedente y devuelve `has_more`. Paginación por cursor: pospuesta conscientemente (retrocompatible añadirla).
- Se mantienen las garantías existentes: validación de campo `[A-Za-z0-9_]+` en doble capa (`422` schema + `400` servicio), percent-encoding de la expresión completa, cabecera `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`.

**Archivos nuevos:**
- `requirements/SPEC-004_List_Items_Search_By_Url.md` — especificación de la feature

**Archivos modificados:**
- `app/services/sharepoint.py` — `_to_odata_literal()`, `_search_field_expected_type()`, `_check_search_value_type()`, `get_list_columns()` (caché TTL 300 s), `_get_bounded()` (paginación acotada), `search_list_items()`
- `app/schemas/sharepoint.py` — `SearchFilter`, `OrderBy`, `ListItemsSearchByUrlRequest` (extra="forbid"), `ListItemsSearchByUrlItem`, `ListItemsSearchByUrlResponse`
- `app/api/v1/endpoints/sharepoint.py` — endpoint `POST /v1/sharepoint/list/items:search`
- `tests/test_sharepoint_service.py`, `tests/test_sharepoint_endpoints.py` — 30 tests nuevos (literales OData, AND multi-campo, validación de esquema, lookup, caché, orderby, truncado/has_more, 422s)
- `VERSION`, `README.md`, `ARQUITECTURA.md`, `arquitecturasUML.md` — documentación a v2.4.0

---

## v2.3.0 — 2026-06-18

### Feature: Upsert genérico de ítems de lista (SPEC-003)

**Contexto:** Un servicio externo disparaba un flujo de Power Automate que, antes de insertar en una lista, **verificaba** si ya existía un registro equivalente (misma clave dentro del mes en curso) y actualizaba o creaba según el resultado. Se migra esa lógica "verificar-y-decidir" a este conector, manteniéndola **genérica**: ninguna lista, columna clave, campo de fecha ni esquema de datos queda quemado en el código — todo viaja en el payload, de modo que cualquier lista de SharePoint pueda reutilizar el mismo endpoint.

**Solución:** Se añade `POST /v1/sharepoint/list/item:upsert` (tercera vía: endpoint de *upsert* explícito, evitando el anti-patrón *flag-argument* sobre el `POST`/`PATCH` ya estables). Orquesta sobre primitivos existentes: resuelve la URL a `site_id`+`list_id` (resolver de SPEC-002), busca coincidencias y crea o actualiza.

```json
{
  "sharepoint_url": "https://host.sharepoint.com/Oper/Lists/Incidencias/View.aspx",
  "match": {
    "key_field": "TicketId",
    "key_value": "INC-12345",
    "date_field": "Created",
    "period": "current_month"
  },
  "data": { "Title": "Incidencia", "Estado": "Resuelto" }
}
```

- **Coincidencia** = igualdad exacta sobre `key_field` **y** (opcional) rango sobre `date_field`. `period` admite atajos con nombre (`current_day`, `current_week`, `current_month`, `current_year`) o rango explícito `{from,to}` (`to` inclusive). Sin `period`/`date_field`, solo por clave; `period` sin `date_field` se rechaza con `422`.
- **Periodo en la zona horaria del tenant:** los límites se calculan en `TENANT_TIMEZONE` (IANA, default `UTC`) y se convierten a UTC para el `$filter` de Graph (intervalo semiabierto `ge`/`lt`).
- **Desenlace:** `result="created"` si 0 coincidencias; `result="updated"` si ≥1 (se actualiza la **primera**). A diferencia del `PATCH`, varias coincidencias **no** son `409`: se emite un `warning` en logs y se devuelve `matched` (nº de coincidencias) para detectar duplicados preexistentes.

**Archivos nuevos:**
- `app/services/period.py` — `resolve_period()`: atajo con nombre / rango explícito → intervalo UTC en la zona del tenant
- `tests/test_period.py` — 10 tests (rango explícito, conversión TZ, atajos, validaciones)

**Archivos modificados:**
- `app/services/sharepoint.py` — nuevo método `find_list_items_for_upsert()` (filtro combinado clave + rango de fecha)
- `app/schemas/sharepoint.py` — modelos `ExplicitPeriod`, `UpsertMatch` (con validador period↔date_field), `UpsertListItemRequest`, `UpsertListItemResponse`
- `app/api/v1/endpoints/sharepoint.py` — endpoint `POST /v1/sharepoint/list/item:upsert`
- `app/core/config.py` — setting `tenant_timezone` (default `UTC`)
- `requirements.txt` — `tzdata` (zoneinfo en Windows)
- `tests/test_sharepoint_endpoints.py`, `tests/test_sharepoint_service.py` — 10 tests nuevos (created/updated/empate, validaciones, filtro)
- `VERSION`, `README.md`, `ARQUITECTURA.md`, `TECNOLOGIAS.md`, `devops/.env.example` — documentación y config a v2.3.0

---

## v2.2.1 — 2026-06-11

### Fix: Endurecimiento de seguridad y paginación de Graph

**Contexto:** Una revisión del código detectó tres grupos de problemas: (1) entradas del caller interpoladas sin validar en URLs de Graph (inyección OData vía `filter_by.field`, escape de carpeta vía `filename`/`folder` en subidas), (2) ausencia de límite de tamaño en uploads (el archivo se bufferizaba entero en memoria antes de fallar en Graph) y falta de seguimiento de la paginación de Graph (`@odata.nextLink`), que podía producir 404/409 falsos en sites con muchas listas/drives o listas grandes, y (3) higiene de configuración (ruta de logs personal hardcodeada como default, `VERSION` ausente de la imagen Docker → `/health` reportaba `0.0.0`).

**Cambios:**
- **Anti-inyección OData** — `filter_by.field` se valida contra `[A-Za-z0-9_]+` tanto en el schema (`422`) como en el servicio (`400`, defensa en profundidad); la expresión `$filter` completa se percent-encodea.
- **Saneo del path de subida** — `filename` y `folder` se rechazan con `400` si contienen `..`, separadores (`/` en filename, `\`) o caracteres de control; cada segmento se percent-encodea antes de construir la URL de Graph (`_encode_drive_path`).
- **Límite de subida** — los dos endpoints de upload leen como máximo 4 MB + 1 byte y responden `413` si se supera, sin bufferizar el archivo completo ni llamar a Graph (`app/core/uploads.py`).
- **Paginación** — nuevo helper `_get_all()` que sigue `@odata.nextLink`; aplicado a `list_sites`, `list_site_lists`, `list_site_drives`, `list_folder_children` y `find_list_items_by_field`. El término `search` de `/sites` también se percent-encodea.
- **Configuración** — `log_dir` ya no tiene una ruta personal hardcodeada como default (vacío = solo consola; configurable con `LOG_DIR`/`LOG_FILE`). El Dockerfile copia `VERSION`, de modo que `/health` reporta la versión real.

**Archivos nuevos:**
- `app/core/uploads.py` — `read_upload()` con `MAX_UPLOAD_BYTES`
- `tests/test_sharepoint_service.py` — 22 tests (saneo de paths, validación de campo, escape OData, paginación)

**Archivos modificados:**
- `app/services/sharepoint.py` — `_encode_drive_path`, `_get_all`, validación de `field`, encoding de `$filter` y `search`
- `app/schemas/sharepoint.py` — `FilterBy.field` con `pattern`
- `app/api/v1/endpoints/sharepoint.py`, `app/api/v1/endpoints/files.py` — uso de `read_upload()`
- `app/core/config.py` — default de `log_dir` vacío
- `devops/Dockerfile` — `COPY VERSION .`
- `devops/.env.example` — variables `LOG_DIR`/`LOG_FILE`
- `tests/test_sharepoint_endpoints.py`, `tests/test_files_endpoint.py` — tests de `422` y `413`
- `VERSION`, `README.md`, `ARQUITECTURA.md`, `TECNOLOGIAS.md` — documentación a v2.2.1

---

## v2.2.0 — 2026-06-10

### Feature: Actualización de ítems de lista por URL localizados por campo único

**Contexto:** Los endpoints by-URL de la v2.1.0 solo permitían **crear** ítems. Actualizar un registro existente exigía conocer su `item_id` interno de Graph, dato del que el caller no dispone: solo tiene la URL de la lista y algún valor de negocio que identifica la fila (p.ej. un código `LATSUP-0000`).

**Solución:** Se añade `PATCH /v1/sharepoint/list/item`, que localiza el registro mediante `filter_by` (`field` + `value`) y aplica `data`. El conector busca el ítem vía Graph (`$filter` sobre `fields/{field}`) para obtener su `id` interno y luego actualiza sus campos. El `filter_by` debe identificar un único registro.

```json
{
  "sharepoint_url": "https://host.sharepoint.com/Oper/Lists/Incidencias/View.aspx",
  "filter_by": { "field": "_x006c_dq4", "value": "LATSUP-0000" },
  "data": { "Title": "Actualización", "Atendida": false }
}
```

Comportamiento según las coincidencias de `filter_by`: `200` si hay exactamente una (ítem actualizado), `404` si ninguna, `409` si más de una (filtro no único — no se modifica nada). La búsqueda añade la cabecera `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly` porque las columnas custom no suelen estar indexadas; el valor se trata como texto y se escapa según OData.

**Archivos modificados:**
- `app/services/sharepoint.py` — helper `_patch`, `_get` admite `extra_headers`, y nuevos métodos `find_list_items_by_field()` y `update_list_item()`
- `app/api/v1/endpoints/sharepoint.py` — endpoint `PATCH /v1/sharepoint/list/item`
- `app/schemas/sharepoint.py` — modelos `FilterBy`, `ListItemUpdateByUrlRequest`, `ListItemUpdateByUrlResponse`
- `tests/test_sharepoint_endpoints.py` — 3 tests nuevos (éxito, 404, 409)
- `VERSION`, `README.md`, `ARQUITECTURA.md`, `TECNOLOGIAS.md` — documentación a v2.2.0

---

## v2.1.0 — 2026-06-10

### Feature: Endpoints orientados a usuario con resolución por URL (SPEC-002)

**Contexto:** La v2.0.0 exige que el caller conozca `site_id`, `list_id` y `drive_id` de Graph. Obtener esos identificadores es inviable para un usuario humano, que solo dispone de la URL de SharePoint tal como aparece en el navegador.

**Solución:** Se añade una capa de resolución (`SharePointResolver`) y dos endpoints nuevos que aceptan una URL "humana" y derivan internamente los identificadores de Graph:
- `POST /v1/sharepoint/list/item` — inserta un ítem en una lista a partir de su URL.
- `POST /v1/sharepoint/upload` — sube un archivo a una biblioteca/carpeta a partir de su URL (`multipart/form-data`).

El site se resuelve con `GET /sites/{host}:/{path}` recortando segmentos hacia la raíz (cubre raíz, managed paths como `/Oper`, `/sites/X`, `/teams/X` y subsites). Las listas se emparejan por `webUrl`/nombre y las bibliotecas por prefijo del `webUrl` del drive, con la carpeta destino tomada del parámetro `?id=` cuando está presente. Los IDs de site resueltos se cachean (son estables). Las claves de `data` siguen siendo los nombres internos de columna.

Los endpoints `/v1/graph/...` de la v2.0.0 se mantienen intactos como API de bajo nivel/discovery.

**Archivos nuevos:**
- `app/services/resolver.py` — `SharePointResolver` (parseo de URL, resolución de site/lista/drive/carpeta)
- `app/api/v1/endpoints/sharepoint.py` — endpoints `POST /v1/sharepoint/list/item` y `POST /v1/sharepoint/upload`
- `app/schemas/sharepoint.py` — modelos de request/response by-URL
- `tests/test_resolver.py`, `tests/test_sharepoint_endpoints.py` — 13 tests nuevos

**Archivos modificados:**
- `app/services/sharepoint.py` — nuevo método `get_site_by_path(hostname, path)`
- `app/core/dependencies.py` — singleton `get_resolver()`
- `app/api/v1/router.py` — registro del router `sharepoint`
- `VERSION`, `README.md`, `ARQUITECTURA.md` — documentación a v2.1.0

---

## v2.0.0 — 2026-06-09

### Refactor: Arquitectura dinámica multi-site con API versionada (SPEC-001)

**Contexto:** La v1.0.0 estaba acoplada a un site SharePoint fijo configurado mediante `SITE_URL`. Los endpoints eran planos (`POST /upload`, `POST /list`) sin versionado y la subida de archivos usaba Base64 en JSON. No había endpoints de discovery, ni descarga, ni metadatos. El logging era básico y los errores de autenticación se propagaban como 500 genérico.

**Solución:** Rediseño completo de la arquitectura hacia un servicio genérico y dinámico. El site, la lista y el drive se identifican por ID en cada llamada. Se introduce una API REST versionada bajo `/v1/graph/...` con 9 endpoints organizados en tres grupos (Discovery, List Items, Files). El logging pasa a JSON estructurado con `request_id` y `client_app_id` propagados a todo el stack.

**Archivos nuevos:**
- `app/core/auth.py` — `TokenManager` con manejo tipado de errores y logging de adquisición de token
- `app/core/config.py` — `Settings` con pydantic-settings, lee versión desde `VERSION`
- `app/core/context.py` — `ContextVar`s para `request_id` y `client_app_id`
- `app/core/dependencies.py` — singletons inyectables vía `lru_cache`
- `app/core/exceptions.py` — `GraphAPIError` tipado con handlers para FastAPI; `-> NoReturn`
- `app/core/logging.py` — `JSONFormatter` con whitelist de campos estructurados
- `app/api/v1/router.py` — router versionado bajo `/v1`
- `app/api/v1/endpoints/discovery.py` — 4 endpoints de exploración de sites/listas/drives/carpetas
- `app/api/v1/endpoints/list_items.py` — GET y POST de ítems de lista
- `app/api/v1/endpoints/files.py` — upload (`multipart/form-data`), metadata y download
- `app/schemas/discovery.py`, `files.py`, `list_items.py` — modelos Pydantic por dominio
- `VERSION` — fuente única de verdad para el número de versión
- `requirements-dev.txt`, `pytest.ini`, `tests/` — suite de tests unitarios (12 tests)

**Archivos eliminados (v1):**
- `app/auth.py`, `app/config.py`, `app/dependencies.py`, `app/models.py`
- `app/routers/upload.py`, `app/routers/list_item.py`

**Archivos modificados:**
- `app/main.py` — middleware con try/except para trazabilidad en 500; exception handlers registrados
- `app/services/sharepoint.py` — métodos de discovery, metadata, download; logging completo
- `devops/docker-compose.yml`, `devops/.env.example`, `devops/deploy.sh` — limpieza de config muerta
- `ARQUITECTURA.md`, `README.md` — documentación actualizada a v2.0.0

---

## v1.0.0 — 2026-05-19

### Feature: Microservicio inicial de integración con SharePoint

**Contexto:** La integración entre Jirito Newsletter y SharePoint se realizaba a través de Power Automate, con errores 408/429/5xx no visibles en logs y lógica distribuida entre código y flujos visuales externos.

**Solución:** Microservicio FastAPI con dos endpoints planos que reemplazaban los webhooks de Power Automate: `POST /upload` (archivo en Base64 + JSON) y `POST /list` (crear ítem en lista). Site único configurado en `SITE_URL`. `TokenManager` con caché de token OAuth2. Logging básico.

**Archivos:**
- `app/main.py`, `app/auth.py`, `app/config.py`, `app/dependencies.py`, `app/models.py`
- `app/routers/upload.py`, `app/routers/list_item.py`
- `app/services/sharepoint.py`
- `devops/Dockerfile`, `devops/docker-compose.yml`, `devops/.env.example`, `devops/deploy.sh`
