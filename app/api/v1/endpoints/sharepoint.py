"""Endpoints orientados a usuario: el caller pasa una URL "humana" de SharePoint
y el conector resuelve internamente los identificadores de Microsoft Graph.

A diferencia de ``/v1/graph/...`` (que exige ``site_id``/``list_id``/``drive_id``
ya conocidos), aquí basta con copiar la URL de la lista o de la biblioteca desde
el navegador.
"""

import logging
from typing import Annotated

from fastapi import APIRouter, Depends, File, Form, UploadFile

from app.core.config import settings
from app.core.dependencies import get_resolver, get_sp
from app.core.exceptions import GraphAPIError
from app.core.uploads import read_upload
from app.schemas.sharepoint import (
    ExplicitPeriod,
    ListItemByUrlRequest,
    ListItemByUrlResponse,
    ListItemUpdateByUrlRequest,
    ListItemUpdateByUrlResponse,
    UpsertListItemRequest,
    UpsertListItemResponse,
    UploadByUrlResponse,
)
from app.services.period import resolve_period
from app.services.resolver import SharePointResolver
from app.services.sharepoint import SharePointService

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/sharepoint", tags=["SharePoint (by URL)"])


@router.post(
    "/list/item",
    response_model=ListItemByUrlResponse,
    status_code=201,
    summary="Insertar un ítem en una lista a partir de su URL",
)
async def create_list_item_by_url(
    payload: ListItemByUrlRequest,
    sp: SharePointService = Depends(get_sp),
    resolver: SharePointResolver = Depends(get_resolver),
):
    """Crea un ítem en la lista identificada por `sharepoint_url`.

    El conector resuelve la URL a `site_id` + `list_id` (vía Graph) y delega la
    inserción. Las claves de `data` deben ser los **nombres internos** de las
    columnas.

    Ejemplo de URL: `https://host.sharepoint.com/Oper/Lists/Incidencias/View.aspx`
    """
    logger.info("create_list_item_by_url → url=%r", payload.sharepoint_url)
    resolved = await resolver.resolve_list(payload.sharepoint_url)
    result = await sp.create_list_item(resolved.site_id, resolved.list_id, payload.data)
    return ListItemByUrlResponse(
        id=result["id"],
        webUrl=result.get("webUrl"),
        site_id=resolved.site_id,
        list_id=resolved.list_id,
    )


@router.patch(
    "/list/item",
    response_model=ListItemUpdateByUrlResponse,
    summary="Actualizar un ítem de una lista localizándolo por un campo único",
)
async def update_list_item_by_url(
    payload: ListItemUpdateByUrlRequest,
    sp: SharePointService = Depends(get_sp),
    resolver: SharePointResolver = Depends(get_resolver),
):
    """Actualiza un ítem de la lista identificada por `sharepoint_url`.

    El registro a modificar se localiza con `filter_by` (`field` + `value`), que
    debe identificar un **único** ítem: el conector lo busca vía Graph para obtener
    su `id` interno y luego aplica `data`. Tanto `filter_by.field` como las claves de
    `data` deben ser **nombres internos** de columna.

    Respuestas de error:
    - `404` si ningún registro coincide con `filter_by`.
    - `409` si coincide más de uno (el filtro no era único).
    """
    logger.info(
        "update_list_item_by_url → url=%r filter=%s=%r",
        payload.sharepoint_url,
        payload.filter_by.field,
        payload.filter_by.value,
    )
    resolved = await resolver.resolve_list(payload.sharepoint_url)
    matches = await sp.find_list_items_by_field(
        resolved.site_id,
        resolved.list_id,
        payload.filter_by.field,
        payload.filter_by.value,
    )

    if not matches:
        raise GraphAPIError(
            404,
            f"No se encontró ningún registro con {payload.filter_by.field}="
            f"'{payload.filter_by.value}' en la lista",
        )
    if len(matches) > 1:
        raise GraphAPIError(
            409,
            f"filter_by no es único: {len(matches)} registros coinciden con "
            f"{payload.filter_by.field}='{payload.filter_by.value}'. "
            "Refina el filtro para identificar un único ítem.",
        )

    item = matches[0]
    item_id = item["id"]
    await sp.update_list_item(
        resolved.site_id, resolved.list_id, item_id, payload.data
    )
    return ListItemUpdateByUrlResponse(
        id=item_id,
        webUrl=item.get("webUrl"),
        site_id=resolved.site_id,
        list_id=resolved.list_id,
    )


@router.post(
    "/list/item:upsert",
    response_model=UpsertListItemResponse,
    summary="Insertar o actualizar un ítem de lista (upsert genérico)",
)
async def upsert_list_item_by_url(
    payload: UpsertListItemRequest,
    sp: SharePointService = Depends(get_sp),
    resolver: SharePointResolver = Depends(get_resolver),
):
    """Inserta o actualiza un ítem de la lista identificada por `sharepoint_url`.

    El conector verifica primero si ya existe un registro que coincida con
    `match` y, según el resultado, **crea** uno nuevo o **actualiza** el existente
    (idempotencia). Es genérico: ni la columna clave, ni la de fecha, ni el
    esquema de `data` están quemados — todo viaja en el payload.

    Coincidencia (`match`):
    - `key_field` + `key_value`: igualdad exacta sobre la columna clave.
    - `date_field` + `period` (opcional): acota a un periodo. `period` admite un
      atajo con nombre (`current_day`, `current_week`, `current_month`,
      `current_year`) o un rango explícito `{ "from": "...", "to": "..." }`.
      Sin `period`/`date_field`, la coincidencia es solo por clave.

    Desenlace:
    - **0 coincidencias** → crea el ítem y responde `result="created"`.
    - **≥1 coincidencia** → actualiza el **primer** registro y responde
      `result="updated"`. Si hay más de uno, se emite un `warning` en los logs
      (no es error); `matched` indica cuántos coincidieron.
    """
    match = payload.match
    logger.info(
        "upsert_list_item_by_url → url=%r key=%s=%r date_field=%r period=%r",
        payload.sharepoint_url,
        match.key_field,
        match.key_value,
        match.date_field,
        match.period,
    )

    resolved = await resolver.resolve_list(payload.sharepoint_url)

    if isinstance(match.period, ExplicitPeriod):
        period_input: "str | tuple[str, str] | None" = (match.period.from_, match.period.to)
    else:
        period_input = match.period
    bounds = resolve_period(period_input, settings.tenant_timezone)
    period_start, period_end = bounds if bounds else (None, None)

    matches = await sp.find_list_items_for_upsert(
        resolved.site_id,
        resolved.list_id,
        key_field=match.key_field,
        key_value=match.key_value,
        date_field=match.date_field,
        period_start=period_start,
        period_end=period_end,
    )

    if not matches:
        created = await sp.create_list_item(resolved.site_id, resolved.list_id, payload.data)
        logger.info(
            "Upsert → created id=%s (%s=%r)",
            created.get("id"),
            match.key_field,
            match.key_value,
            extra={"site_id": resolved.site_id, "list_id": resolved.list_id},
        )
        return UpsertListItemResponse(
            result="created",
            id=created["id"],
            webUrl=created.get("webUrl"),
            site_id=resolved.site_id,
            list_id=resolved.list_id,
            matched=0,
        )

    if len(matches) > 1:
        logger.warning(
            "Upsert: %d registros coinciden con %s=%r (+periodo) — se actualiza el primero "
            "(posibles duplicados preexistentes)",
            len(matches),
            match.key_field,
            match.key_value,
            extra={"site_id": resolved.site_id, "list_id": resolved.list_id},
        )

    item = matches[0]
    item_id = item["id"]
    await sp.update_list_item(resolved.site_id, resolved.list_id, item_id, payload.data)
    logger.info(
        "Upsert → updated id=%s (%s=%r, matched=%d)",
        item_id,
        match.key_field,
        match.key_value,
        len(matches),
        extra={"site_id": resolved.site_id, "list_id": resolved.list_id},
    )
    return UpsertListItemResponse(
        result="updated",
        id=item_id,
        webUrl=item.get("webUrl"),
        site_id=resolved.site_id,
        list_id=resolved.list_id,
        matched=len(matches),
    )


@router.post(
    "/upload",
    response_model=UploadByUrlResponse,
    status_code=201,
    summary="Subir un archivo a una biblioteca a partir de su URL",
)
async def upload_file_by_url(
    sharepoint_url: Annotated[str, Form(description="URL de la biblioteca/carpeta destino")],
    file: Annotated[UploadFile, File()],
    sp: SharePointService = Depends(get_sp),
    resolver: SharePointResolver = Depends(get_resolver),
):
    """Sube un archivo a la biblioteca/carpeta identificada por `sharepoint_url`.

    El cuerpo es `multipart/form-data` con los campos `sharepoint_url` y `file`.
    El conector resuelve la URL a `site_id` + `drive_id` + carpeta destino; la
    carpeta se crea automáticamente si no existe. Tamaño máximo: 4 MB.

    La URL puede ser la de la biblioteca (`.../Forms/AllItems.aspx`) o incluir el
    parámetro `?id=` con la ruta de una subcarpeta concreta.
    """
    logger.info(
        "upload_file_by_url → url=%r filename=%r", sharepoint_url, file.filename
    )
    resolved = await resolver.resolve_upload_target(sharepoint_url)
    data = await read_upload(file)
    result = await sp.upload_file(
        site_id=resolved.site_id,
        drive_id=resolved.drive_id,
        folder=resolved.folder,
        filename=file.filename or "upload",
        data=data,
    )
    return UploadByUrlResponse(
        id=result["id"],
        name=result["name"],
        size=result.get("size"),
        webUrl=result["webUrl"],
        site_id=resolved.site_id,
        drive_id=resolved.drive_id,
        folder=resolved.folder,
    )
