# Tecnologías utilizadas — SharePoint Connector

**Versión:** 2.3.0  
**Fecha:** 2026-06-18  
**Autor:** Juan Camilo López Alzate — Latinia  

---

## Índice

1. [Visión general del stack](#1-visión-general-del-stack)
2. [Python 3.12](#2-python-312)
3. [FastAPI](#3-fastapi)
4. [Pydantic y pydantic-settings](#4-pydantic-y-pydantic-settings)
5. [httpx](#5-httpx)
6. [Uvicorn](#6-uvicorn)
7. [Microsoft Graph API](#7-microsoft-graph-api)
8. [Docker y Docker Compose](#8-docker-y-docker-compose)

---

## 1. Visión general del stack

```
┌─────────────────────────────────────────────────────┐
│                  Docker Container                   │
│                                                     │
│  ┌──────────┐   ┌──────────┐   ┌─────────────────┐ │
│  │ Uvicorn  │ → │ FastAPI  │ → │ SharePointSvc   │ │
│  │  ASGI    │   │  Router  │   │  (httpx async)  │ │
│  └──────────┘   └──────────┘   └────────┬────────┘ │
│                                         │           │
│  ┌──────────────────────────────────────┘           │
│  │  Pydantic (modelos)  ·  pydantic-settings (env)  │
│  └──────────────────────────────────────────────────┘
└─────────────────────────────────────────────────────┘
         │ HTTPS
         ▼
┌─────────────────────────┐
│   Azure AD              │
│   OAuth2 token endpoint │
└────────────┬────────────┘
             │ Bearer token
             ▼
┌─────────────────────────┐
│   Microsoft Graph API   │
│   /sites / /drives      │
│   /lists / /items       │
└────────────┬────────────┘
             │
             ▼
┌─────────────────────────┐
│   SharePoint Online     │
└─────────────────────────┘
```

| Capa | Tecnología | Versión |
|---|---|---|
| Lenguaje | Python | 3.12 |
| Framework web | FastAPI | ≥ 0.115 |
| Validación de datos | Pydantic | ≥ 2.11 |
| Configuración | pydantic-settings | ≥ 2.7 |
| Cliente HTTP | httpx | ≥ 0.28 |
| Servidor ASGI | Uvicorn | ≥ 0.30 |
| API de integración | Microsoft Graph API | v1.0 |
| Autenticación | OAuth 2.0 Client Credentials | RFC 6749 |
| Contenerización | Docker + Docker Compose | — |

---

## 2. Python 3.12

### Qué es
Python es un lenguaje de programación interpretado, de tipado dinámico y sintaxis concisa. La versión 3.12 es la versión estable actual con mejoras de rendimiento respecto a versiones anteriores.

### Por qué se eligió
- Es el mismo lenguaje que usa Jirito Newsletter, lo que facilita el mantenimiento por el mismo equipo.
- El ecosistema de librerías para integración con APIs REST (httpx, Pydantic) es maduro y bien mantenido.
- Las anotaciones de tipo (`str | None`, `dict[str, Any]`) permiten validación estática sin sacrificar agilidad.

### Cómo se usa en el proyecto
- Todo el código del servicio está escrito en Python 3.12.
- Se usan type hints en todos los módulos para mejorar la legibilidad y la detección temprana de errores.

### Características de Python 3.12 aprovechadas
| Característica | Uso en el proyecto |
|---|---|
| `str \| None` (union types) | Campos opcionales en modelos Pydantic (p.ej. `webUrl` en las respuestas) |
| `dict[str, Any]` genérico | Tipado del campo `fields`/`data` en los schemas de lista (`CreateListItemRequest`, `ListItemByUrlRequest`) |
| `asyncio` nativo | Soporte async/await en toda la capa de servicio |

---

## 3. FastAPI

### Qué es
FastAPI es un framework web moderno para construir APIs REST con Python. Está basado en los estándares **OpenAPI** y **JSON Schema**, y usa `asyncio` de forma nativa para manejar concurrencia sin bloqueo.

**Repositorio oficial:** https://fastapi.tiangolo.com

### Por qué se eligió

| Criterio | FastAPI | Flask | Django REST |
|---|---|---|---|
| Soporte async nativo | Sí | Limitado | Limitado |
| Validación automática | Sí (Pydantic) | Manual | Parcial |
| Documentación automática | Sí (Swagger/Redoc) | No | No |
| Peso de la librería | Ligero | Ligero | Pesado |
| Curva de aprendizaje | Baja | Muy baja | Alta |

FastAPI genera automáticamente documentación interactiva en `/docs` (Swagger UI) y `/redoc`, lo que facilita las pruebas sin necesidad de herramientas externas como Postman.

### Cómo se usa en el proyecto

```python
# app/main.py
app = FastAPI(title=settings.app_name, version=settings.app_version)
app.include_router(v1_router)   # agrupa /v1/sharepoint/... y /v1/graph/...
```

- El título y la versión se leen de `Settings` (la versión, a su vez, del archivo `VERSION`), evitando duplicar el número en el código.
- Los endpoints se organizan en routers por dominio (`discovery`, `list_items`, `files`, `sharepoint`), agrupados bajo el router `/v1`; `/health` cuelga directamente de la app.
- FastAPI valida automáticamente el JSON de entrada contra los modelos Pydantic antes de llegar al handler.
- Si el JSON no cumple el esquema, devuelve `422 Unprocessable Entity` con detalle del campo inválido — sin escribir una sola línea de validación manual.

### Documentación automática

Una vez levantado el contenedor, la documentación interactiva está disponible en:
- `http://localhost:8003/docs` — Swagger UI
- `http://localhost:8003/redoc` — ReDoc

---

## 4. Pydantic y pydantic-settings

### Qué es
**Pydantic** es una librería de validación de datos basada en type hints de Python. Convierte y valida datos de entrada (JSON, dicts) en objetos Python tipados, rechazando datos inválidos con mensajes de error descriptivos.

**pydantic-settings** es una extensión que permite leer configuración desde variables de entorno y archivos `.env` con la misma mecánica de validación.

### Por qué se eligió
- FastAPI usa Pydantic internamente, por lo que no añade dependencias extra.
- Permite definir el esquema de la API y la validación en un solo lugar (el modelo), sin duplicar lógica.
- El campo `fields`/`data` de tipo `dict[str, Any]` en los schemas de lista es posible gracias a la flexibilidad de Pydantic con tipos genéricos.

### Cómo se usa en el proyecto

**Modelos de entrada (app/schemas/):**

Los modelos están organizados por dominio en `app/schemas/` (`discovery.py`, `list_items.py`, `files.py`, `sharepoint.py`). Los campos de lista se tipan como `dict[str, Any]` para aceptar cualquier columna de SharePoint:
```python
# app/schemas/list_items.py — API de bajo nivel
class CreateListItemRequest(BaseModel):
    fields: dict[str, Any]        # acepta cualquier tipo

# app/schemas/sharepoint.py — API por URL (actualización por campo único)
class FilterBy(BaseModel):
    field: str                    # nombre interno de columna
    value: str                    # valor que identifica el registro

class ListItemUpdateByUrlRequest(BaseModel):
    sharepoint_url: str
    filter_by: FilterBy
    data: dict[str, Any]
```

Pydantic acepta en `fields`/`data` valores de tipo `str`, `int`, `float`, `bool` y `list` sin configuración adicional — simplemente los pasa tal cual a Graph API. La validación anidada (`filter_by` como submodelo `FilterBy`) la resuelve Pydantic automáticamente.

**Configuración desde entorno (app/core/config.py):**
```python
class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", case_sensitive=False)

    # Azure AD / Microsoft Identity Platform
    tenant_id: str          # obligatorio — falla al arrancar si no está
    client_id: str          # obligatorio
    client_secret: str      # obligatorio

    # Application
    app_name: str = "SharePoint Connector"
    app_version: str = _read_version()   # leído del archivo VERSION
    log_level: str = "INFO"
```

Desde la v2.0.0 el servicio es **multi-site**: ya no existe un `SITE_URL` fijo en configuración; el site, la lista y la biblioteca se identifican en cada llamada (por ID en `/v1/graph` o resueltos desde la URL en `/v1/sharepoint`). Si una variable obligatoria (`tenant_id`/`client_id`/`client_secret`) no está definida en el entorno, el servicio **falla al arrancar** con un mensaje claro indicando qué falta — evitando arranques en estado inconsistente.

---

## 5. httpx

### Qué es
`httpx` es un cliente HTTP para Python que soporta tanto llamadas síncronas como **asíncronas** (async/await), con una API muy similar a la popular librería `requests`.

**Repositorio oficial:** https://www.python-httpx.org

### Por qué se eligió frente a `requests`

| Criterio | httpx | requests |
|---|---|---|
| Soporte async | Sí (`AsyncClient`) | No |
| HTTP/2 | Sí | No |
| API similar a requests | Sí | — |
| Soporte en FastAPI | Recomendado | Compatible pero bloqueante |

El servicio usa `async/await` en toda la capa de llamadas a Graph API. Usar `requests` (síncrono) dentro de un handler async bloquearía el event loop de Python, degradando el rendimiento bajo carga. `httpx.AsyncClient` resuelve esto de forma nativa.

### Cómo se usa en el proyecto

```python
# app/services/sharepoint.py
_TIMEOUT = httpx.Timeout(60.0)  # uploads pueden llegar a 4 MB sobre SharePoint

async def _get(self, url: str, extra_headers: dict | None = None) -> dict:
    headers = await self._auth_headers()
    if extra_headers:
        headers.update(extra_headers)   # p.ej. Prefer para $filter no indexado
    async with httpx.AsyncClient(timeout=_TIMEOUT) as c:
        r = await c.get(url, headers=headers)
        r.raise_for_status()   # lanza excepción en 4xx/5xx con el cuerpo del error
        return r.json()
```

El método `raise_for_status()` propaga el error HTTP original de Graph API (incluyendo su mensaje), que el router captura y devuelve al caller con código `502`. El parámetro opcional `extra_headers` permite añadir cabeceras puntuales por llamada (como `Prefer` en la búsqueda por campo) sin alterar el resto de peticiones. Para escritura existen los helpers análogos `_post`, `_patch` (actualización de ítems) y `_put_bytes` (subida de archivos).

---

## 6. Uvicorn

### Qué es
Uvicorn es un servidor ASGI (*Asynchronous Server Gateway Interface*) de alto rendimiento para Python. Es el servidor de producción recomendado para FastAPI.

**ASGI** es el estándar moderno para servidores web Python con soporte nativo de concurrencia asíncrona, sucesor de WSGI (usado por Flask/Django).

### Por qué se eligió
- Es la combinación estándar y recomendada con FastAPI.
- Soporta `asyncio` de forma nativa, aprovechando el modelo async del servicio.
- Arranque inmediato y sin configuración compleja.

### Cómo se usa en el proyecto

```dockerfile
# Dockerfile
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8003"]
```

Se ejecuta con un único worker porque el `TokenManager` y las cachés de IDs de SharePoint viven en memoria del proceso. Múltiples workers crearían instancias independientes, duplicando llamadas de resolución innecesarias. Para escalar horizontalmente, se añaden réplicas del contenedor completo.

---

## 7. Microsoft Graph API

### Qué es
Microsoft Graph API es la API REST unificada de Microsoft para acceder a datos y servicios de Microsoft 365: SharePoint, OneDrive, Teams, Outlook, Azure AD, entre otros.

**Documentación oficial:** https://learn.microsoft.com/en-us/graph/overview  
**Versión usada:** `v1.0` (versión estable)

### Por qué se eligió frente a la SharePoint REST API clásica

| Criterio | Graph API v1.0 | SharePoint REST (`_api/`) |
|---|---|---|
| Autenticación moderna | OAuth2 (Azure AD) | OAuth2 + cookies legacy |
| Endpoint unificado | `graph.microsoft.com` | Por tenant/site |
| Soporte futuro | Activo (Microsoft) | Mantenimiento mínimo |
| Compatibilidad con `Sites.Selected` | Sí | Parcial |
| Documentación | Extensa y actualizada | Dispersa |

### Endpoints utilizados

#### Resolución de site
```
GET https://graph.microsoft.com/v1.0/sites/{hostname}:{site-path}
```
Devuelve el `id` interno del site, necesario para todas las operaciones posteriores. Se cachea en memoria tras la primera llamada.

#### Resolución de biblioteca de documentos
```
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives
```
Lista las bibliotecas disponibles. Se filtra por nombre y se cachea el `id`.

#### Subida de archivo
```
PUT https://graph.microsoft.com/v1.0/sites/{site-id}/drives/{drive-id}/root:/{folder}/{filename}:/content
```
Cuerpo: bytes del archivo (`Content-Type: application/octet-stream`).  
Si el archivo ya existe, lo sobreescribe. Si no existe, lo crea junto con las carpetas intermedias que sean necesarias.

#### Resolución de lista
```
GET https://graph.microsoft.com/v1.0/sites/{site-id}/lists/{list-name}
```
Acepta el nombre visible de la lista o su ID. Se cachea tras la primera llamada.

#### Creación de ítem en lista
```
POST https://graph.microsoft.com/v1.0/sites/{site-id}/lists/{list-id}/items
Content-Type: application/json

{
  "fields": {
    "Title": "LATSUP-6585",
    "CustomField": "valor"
  }
}
```
Graph API espera nativamente el wrapper `{"fields": {...}}`, por lo que el diseño del endpoint `POST /v1/graph/.../items` del conector es un reflejo directo de la API de Microsoft.

#### Búsqueda de ítem por campo
```
GET https://graph.microsoft.com/v1.0/sites/{site-id}/lists/{list-id}/items?$expand=fields&$filter=fields/{campo} eq '{valor}'
Prefer: HonorNonIndexedQueriesWarningMayFailRandomly
```
Para localizar el ítem a actualizar a partir de un campo único, se usa `$filter` sobre `fields/{campo}`. Como las columnas *custom* de SharePoint no suelen estar **indexadas**, Graph rechazaría la consulta salvo que se envíe la cabecera `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`, que la autoriza explícitamente (Microsoft advierte que en listas muy grandes puede fallar de forma intermitente). El valor se escapa según OData (las comillas simples se duplican).

#### Actualización de ítem en lista
```
PATCH https://graph.microsoft.com/v1.0/sites/{site-id}/lists/{list-id}/items/{item-id}/fields
Content-Type: application/json

{
  "Title": "Prueba de inyección - actualización",
  "_x006c_dq4": "LATSUP-0001",
  "Atendida": false
}
```
A diferencia de la creación (`POST .../items`), el `PATCH` apunta al subrecurso `/fields` del ítem y el cuerpo es el conjunto de campos **sin** el wrapper `{"fields": {...}}`. Devuelve el `fieldValueSet` actualizado.

#### Búsqueda combinada para el upsert (clave + periodo)
```
GET .../items?$expand=fields&$filter=fields/{clave} eq '{valor}' and fields/{fecha} ge '{inicio}' and fields/{fecha} lt '{fin}'
Prefer: HonorNonIndexedQueriesWarningMayFailRandomly
```
El endpoint `:upsert` extiende el `$filter` anterior con un **rango de fecha** opcional sobre una columna de tipo Fecha (`Created`, `Modified` o una propia). Las columnas de fecha de Graph se almacenan y comparan en **UTC**, así que los límites del periodo se calculan primero en la **zona horaria del tenant** (`TENANT_TIMEZONE`, IANA, vía `zoneinfo`/`tzdata`) y luego se convierten a UTC con formato `YYYY-MM-DDTHH:MM:SSZ`. El intervalo es **semiabierto** (`ge` inicio, `lt` fin) para que periodos contiguos no se solapen. Los nombres de columna se validan contra `[A-Za-z0-9_]+` (anti-inyección OData) y la expresión completa se percent-encodea.

### Caché de IDs

```mermaid
sequenceDiagram
    participant App as SharePointService
    participant Graph as Graph API

    Note over App: Primera llamada
    App->>Graph: GET /sites/{host}:{path}
    Graph-->>App: site_id
    App->>App: _cached_site_id = site_id (caché)

    Note over App: Llamadas siguientes
    App->>App: return _cached_site_id (sin llamada a Graph)
```

Site ID, Drive ID y List ID se resuelven una sola vez por vida del proceso y se guardan en memoria. Esto reduce la latencia y el número de llamadas a Graph API.

---

## 8. Docker y Docker Compose

### Qué es
**Docker** es una plataforma de contenerización que empaqueta una aplicación junto con todas sus dependencias en una unidad portable llamada *contenedor*. Los contenedores son aislados del sistema operativo anfitrión y entre sí.

**Docker Compose** es una herramienta para definir y gestionar múltiples contenedores como un servicio único, usando un archivo YAML.

### Por qué se eligió
- Garantiza que el servicio se comporta igual en desarrollo, pruebas y producción.
- Elimina el problema de dependencias del sistema operativo anfitrión.
- Permite integrar el conector en una red Docker existente (p.ej. Jirito Newsletter) con comunicación interna por nombre de servicio.
- Cumple el requisito de Jimmy: el servicio debe ser independiente y portable.

### Estructura del Dockerfile

```dockerfile
FROM python:3.12-slim          # imagen base mínima (~50 MB vs ~900 MB de la completa)

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt  # dependencias en capa separada

COPY app/ ./app/               # código fuente en la última capa (más frecuentemente cambia)

EXPOSE 8003

HEALTHCHECK ...                # Docker comprueba que el servicio responde

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8003"]
```

El orden de las instrucciones está optimizado para **aprovechar la caché de capas de Docker**: las dependencias (que cambian poco) se instalan antes que el código fuente (que cambia frecuentemente). Así, un rebuild tras un cambio de código no reinstala las dependencias.

### Healthcheck

```yaml
healthcheck:
  test: ["CMD", "python", "-c", "import urllib.request; urllib.request.urlopen('http://localhost:8003/health')"]
  interval: 30s
  timeout: 10s
  start_period: 15s
  retries: 3
```

Docker comprueba cada 30 segundos que el servicio responde en `/health`. Si falla 3 veces consecutivas, Docker marca el contenedor como `unhealthy`, lo cual puede disparar alertas o reinicios según la política de orquestación.

### Integración con red Docker de Jirito

Cuando ambos servicios están en la misma red Docker, el conector es accesible por nombre de servicio sin exponer puertos al host:

```
jirito-app ──▶ http://sharepoint-connector:8003/upload
               (resolución DNS interna de Docker)
```

Esto evita exposición innecesaria del conector a la red exterior.

