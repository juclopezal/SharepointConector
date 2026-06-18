import logging
import re
from urllib.parse import quote

import httpx

from app.core.auth import TokenManager
from app.core.context import client_app_id_var, request_id_var
from app.core.exceptions import GraphAPIError, raise_from_httpx

logger = logging.getLogger(__name__)

GRAPH = "https://graph.microsoft.com/v1.0"
_TIMEOUT = httpx.Timeout(60.0)

# Nombres internos de columna de SharePoint: solo alfanuméricos y guion bajo.
# Evita que un `field` arbitrario inyecte operadores en la expresión $filter.
_FIELD_NAME_RE = re.compile(r"^[A-Za-z0-9_]+$")


def _encode_drive_path(folder: str, filename: str) -> str:
    """Construye el path remoto (carpeta + archivo) saneado y percent-encodeado.

    Rechaza nombres/segmentos que podrían escapar de la carpeta destino o
    corromper la URL de Graph (``..``, separadores, caracteres de control).
    """
    name = (filename or "").strip()
    if (
        not name
        or name in {".", ".."}
        or "/" in name
        or "\\" in name
        or any(ord(c) < 32 for c in name)
    ):
        raise GraphAPIError(400, f"Nombre de archivo inválido: {filename!r}")

    segments = [s for s in (folder or "").split("/") if s]
    if any(s in {".", ".."} or "\\" in s for s in segments):
        raise GraphAPIError(400, f"Ruta de carpeta inválida: {folder!r}")

    return "/".join(quote(s, safe="") for s in [*segments, name])


def _ctx() -> dict:
    """Return current request context fields for structured logging."""
    return {
        "request_id": request_id_var.get(),
        "client_app_id": client_app_id_var.get(),
    }


class SharePointService:
    def __init__(self, token_manager: TokenManager):
        self._tm = token_manager

    # ------------------------------------------------------------------
    # HTTP helpers
    # ------------------------------------------------------------------

    async def _auth_headers(self, content_type: str = "application/json") -> dict:
        return {
            "Authorization": f"Bearer {await self._tm.get_token()}",
            "Content-Type": content_type,
        }

    async def _get(self, url: str, extra_headers: dict | None = None) -> dict:
        logger.debug("Graph GET %s", url, extra={**_ctx(), "graph_url": url})
        headers = await self._auth_headers()
        if extra_headers:
            headers.update(extra_headers)
        async with httpx.AsyncClient(timeout=_TIMEOUT) as c:
            try:
                r = await c.get(url, headers=headers)
                r.raise_for_status()
                return r.json()
            except httpx.HTTPStatusError as e:
                raise_from_httpx(e)

    async def _get_all(self, url: str, extra_headers: dict | None = None) -> list[dict]:
        """GET paginado: sigue ``@odata.nextLink`` acumulando los ``value``.

        Graph devuelve los listados por páginas (~200 elementos); sin esto,
        los resultados más allá de la primera página se perderían en silencio.
        """
        items: list[dict] = []
        next_url: str | None = url
        while next_url:
            data = await self._get(next_url, extra_headers)
            items.extend(data.get("value", []))
            next_url = data.get("@odata.nextLink")
        return items

    async def _post(self, url: str, body: dict) -> dict:
        logger.debug("Graph POST %s", url, extra={**_ctx(), "graph_url": url})
        async with httpx.AsyncClient(timeout=_TIMEOUT) as c:
            try:
                r = await c.post(url, json=body, headers=await self._auth_headers())
                r.raise_for_status()
                return r.json()
            except httpx.HTTPStatusError as e:
                raise_from_httpx(e)

    async def _patch(self, url: str, body: dict) -> dict:
        logger.debug("Graph PATCH %s", url, extra={**_ctx(), "graph_url": url})
        async with httpx.AsyncClient(timeout=_TIMEOUT) as c:
            try:
                r = await c.patch(url, json=body, headers=await self._auth_headers())
                r.raise_for_status()
                return r.json()
            except httpx.HTTPStatusError as e:
                raise_from_httpx(e)

    async def _put_bytes(self, url: str, data: bytes) -> dict:
        logger.debug(
            "Graph PUT %s (%d bytes)", url, len(data), extra={**_ctx(), "graph_url": url}
        )
        headers = {
            "Authorization": f"Bearer {await self._tm.get_token()}",
            "Content-Type": "application/octet-stream",
        }
        async with httpx.AsyncClient(timeout=_TIMEOUT) as c:
            try:
                r = await c.put(url, content=data, headers=headers)
                r.raise_for_status()
                return r.json()
            except httpx.HTTPStatusError as e:
                raise_from_httpx(e)

    # ------------------------------------------------------------------
    # Discovery
    # ------------------------------------------------------------------

    async def list_sites(self, search: str = "*") -> list[dict]:
        sites = await self._get_all(f"{GRAPH}/sites?search={quote(search, safe='')}")
        logger.info("Listed %d SharePoint sites", len(sites), extra=_ctx())
        return sites

    async def get_site_by_path(self, hostname: str, site_path: str = "") -> dict:
        """Resolve a SharePoint site by its hostname and server-relative path.

        ``site_path`` is the path of the *web*, e.g. ``/sites/IADocs``,
        ``/teams/Soporte`` or a managed-path/root web like ``/Oper``. Pass an
        empty string to resolve the root site of the host.

        Graph reference:
        ``GET /sites/{hostname}:/{server-relative-path}`` (or ``GET /sites/{hostname}``
        for the root site).
        """
        clean = site_path.strip("/")
        if clean:
            url = f"{GRAPH}/sites/{hostname}:/{clean}"
        else:
            url = f"{GRAPH}/sites/{hostname}"
        data = await self._get(url)
        logger.debug(
            "Resolved site '%s' (path=%r) → %s",
            hostname,
            site_path,
            data.get("id"),
            extra={**_ctx(), "site_id": data.get("id")},
        )
        return data

    async def list_site_lists(self, site_id: str) -> list[dict]:
        lists = await self._get_all(f"{GRAPH}/sites/{site_id}/lists")
        logger.info(
            "Listed %d lists for site %s",
            len(lists),
            site_id,
            extra={**_ctx(), "site_id": site_id},
        )
        return lists

    async def list_site_drives(self, site_id: str) -> list[dict]:
        drives = await self._get_all(f"{GRAPH}/sites/{site_id}/drives")
        logger.info(
            "Listed %d drives for site %s",
            len(drives),
            site_id,
            extra={**_ctx(), "site_id": site_id},
        )
        return drives

    async def list_folder_children(
        self, site_id: str, drive_id: str, item_id: str | None = None
    ) -> list[dict]:
        if item_id:
            url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/items/{item_id}/children"
        else:
            url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/root/children"
        items = await self._get_all(url)
        logger.info(
            "Listed %d items in drive %s (parent=%s)",
            len(items),
            drive_id,
            item_id or "root",
            extra={**_ctx(), "site_id": site_id, "drive_id": drive_id},
        )
        return items

    # ------------------------------------------------------------------
    # List items
    # ------------------------------------------------------------------

    async def get_list_items(self, site_id: str, list_id: str, top: int = 20) -> list[dict]:
        url = f"{GRAPH}/sites/{site_id}/lists/{list_id}/items?$expand=fields&$top={top}"
        data = await self._get(url)
        items = data.get("value", [])
        logger.info(
            "Retrieved %d items from list %s",
            len(items),
            list_id,
            extra={**_ctx(), "site_id": site_id, "list_id": list_id},
        )
        return items

    async def create_list_item(self, site_id: str, list_id: str, fields: dict) -> dict:
        url = f"{GRAPH}/sites/{site_id}/lists/{list_id}/items"
        result = await self._post(url, {"fields": fields})
        logger.info(
            "Created list item id=%s in list %s / site %s",
            result.get("id"),
            list_id,
            site_id,
            extra={**_ctx(), "site_id": site_id, "list_id": list_id},
        )
        return result

    async def find_list_items_by_field(
        self, site_id: str, list_id: str, field: str, value: str
    ) -> list[dict]:
        """Busca ítems cuyo campo interno ``field`` sea igual a ``value``.

        Usa ``$filter=fields/{field} eq '{value}'`` contra Graph. Como las columnas
        custom de SharePoint no suelen estar indexadas, se envía la cabecera
        ``Prefer: HonorNonIndexedQueriesWarningMayFailRandomly``, que autoriza a Graph
        a resolver la consulta sobre columnas no indexadas (Microsoft advierte que en
        listas muy grandes puede fallar de forma intermitente).

        El valor se escapa según OData (las comillas simples se duplican) y la
        expresión completa se percent-encodea; el nombre del campo se valida
        contra ``[A-Za-z0-9_]+`` para impedir inyección de operadores OData.
        """
        if not _FIELD_NAME_RE.match(field):
            raise GraphAPIError(
                400,
                f"Nombre de campo inválido: {field!r}. Debe ser el nombre interno "
                "de la columna (solo letras, dígitos y '_').",
            )
        escaped = value.replace("'", "''")
        flt = quote(f"fields/{field} eq '{escaped}'", safe="")
        url = (
            f"{GRAPH}/sites/{site_id}/lists/{list_id}/items"
            f"?$expand=fields&$filter={flt}"
        )
        items = await self._get_all(
            url, extra_headers={"Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"}
        )
        logger.info(
            "Filtered list %s by %s=%r → %d match(es)",
            list_id,
            field,
            value,
            len(items),
            extra={**_ctx(), "site_id": site_id, "list_id": list_id},
        )
        return items

    async def find_list_items_for_upsert(
        self,
        site_id: str,
        list_id: str,
        *,
        key_field: str,
        key_value: str,
        date_field: str | None = None,
        period_start: str | None = None,
        period_end: str | None = None,
    ) -> list[dict]:
        """Busca coincidencias para el upsert: clave **y**, si se indica, periodo.

        Combina con ``and`` dos condiciones sobre la lista:

        1. **Clave** — igualdad exacta ``fields/{key_field} eq '{key_value}'``.
        2. **Periodo** (opcional) — rango sobre una columna de fecha:
           ``fields/{date_field} ge '{inicio}' and fields/{date_field} lt '{fin}'``.
           Solo se aplica si ``date_field`` y ambos límites están presentes; si no,
           la coincidencia es **solo por clave**.

        Los nombres de campo se validan contra ``[A-Za-z0-9_]+`` (anti-inyección
        OData), el valor de la clave se escapa según OData y la expresión completa
        se percent-encodea. Se envía ``Prefer: HonorNonIndexedQueriesWarningMayFailRandomly``
        porque las columnas clave/fecha pueden no estar indexadas.
        """
        if not _FIELD_NAME_RE.match(key_field):
            raise GraphAPIError(
                400,
                f"Nombre de campo clave inválido: {key_field!r}. Debe ser el nombre "
                "interno de la columna (solo letras, dígitos y '_').",
            )

        escaped = key_value.replace("'", "''")
        clauses = [f"fields/{key_field} eq '{escaped}'"]

        if date_field and period_start and period_end:
            if not _FIELD_NAME_RE.match(date_field):
                raise GraphAPIError(
                    400,
                    f"Nombre de campo de fecha inválido: {date_field!r}. Debe ser el "
                    "nombre interno de la columna (solo letras, dígitos y '_').",
                )
            clauses.append(f"fields/{date_field} ge '{period_start}'")
            clauses.append(f"fields/{date_field} lt '{period_end}'")

        flt = quote(" and ".join(clauses), safe="")
        url = (
            f"{GRAPH}/sites/{site_id}/lists/{list_id}/items"
            f"?$expand=fields&$filter={flt}"
        )
        items = await self._get_all(
            url, extra_headers={"Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly"}
        )
        logger.info(
            "Upsert lookup on list %s by %s=%r (period=%s) → %d match(es)",
            list_id,
            key_field,
            key_value,
            f"{period_start}..{period_end}" if period_start else "none",
            len(items),
            extra={**_ctx(), "site_id": site_id, "list_id": list_id},
        )
        return items

    async def update_list_item(
        self, site_id: str, list_id: str, item_id: str, fields: dict
    ) -> dict:
        """Actualiza los campos de un ítem existente (``PATCH .../items/{id}/fields``).

        El cuerpo es el conjunto de campos a modificar (claves = nombres internos),
        sin envolver en ``{"fields": ...}``. Devuelve el ``fieldValueSet`` resultante.
        """
        url = f"{GRAPH}/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"
        result = await self._patch(url, fields)
        logger.info(
            "Updated list item id=%s in list %s / site %s",
            item_id,
            list_id,
            site_id,
            extra={**_ctx(), "site_id": site_id, "list_id": list_id},
        )
        return result

    # ------------------------------------------------------------------
    # Files
    # ------------------------------------------------------------------

    async def upload_file(
        self,
        site_id: str,
        drive_id: str,
        folder: str,
        filename: str,
        data: bytes,
    ) -> dict:
        remote_path = _encode_drive_path(folder, filename)
        url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/root:/{remote_path}:/content"
        result = await self._put_bytes(url, data)
        logger.info(
            "Uploaded file '%s' (%d bytes) → drive %s / site %s",
            filename,
            len(data),
            drive_id,
            site_id,
            extra={
                **_ctx(),
                "site_id": site_id,
                "drive_id": drive_id,
                "file_name": filename,
            },
        )
        return result

    async def get_file_metadata(self, site_id: str, drive_id: str, item_id: str) -> dict:
        url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/items/{item_id}"
        result = await self._get(url)
        logger.info(
            "Retrieved metadata for item %s in drive %s / site %s",
            item_id,
            drive_id,
            site_id,
            extra={
                **_ctx(),
                "site_id": site_id,
                "drive_id": drive_id,
                "item_id": item_id,
            },
        )
        return result

    async def download_file_bytes(
        self, site_id: str, drive_id: str, item_id: str
    ) -> tuple[bytes, str, str]:
        metadata = await self._get(
            f"{GRAPH}/sites/{site_id}/drives/{drive_id}/items/{item_id}"
        )
        download_url: str = metadata.get("@microsoft.graph.downloadUrl", "")
        name: str = metadata.get("name", "file")
        mime: str = metadata.get("file", {}).get("mimeType", "application/octet-stream")

        logger.info(
            "Downloading file '%s' from drive %s / site %s",
            name,
            drive_id,
            site_id,
            extra={
                **_ctx(),
                "site_id": site_id,
                "drive_id": drive_id,
                "item_id": item_id,
                "file_name": name,
            },
        )

        async with httpx.AsyncClient(timeout=_TIMEOUT) as c:
            try:
                r = await c.get(download_url)
                r.raise_for_status()
            except httpx.HTTPStatusError as e:
                raise_from_httpx(e)
            content = r.content

        logger.info(
            "Downloaded file '%s' (%d bytes) from drive %s / site %s",
            name,
            len(content),
            drive_id,
            site_id,
            extra={
                **_ctx(),
                "site_id": site_id,
                "drive_id": drive_id,
                "item_id": item_id,
                "file_name": name,
            },
        )
        return content, name, mime
