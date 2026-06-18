"""Schemas de los endpoints orientados a usuario (resolución por URL)."""

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator


class ListItemByUrlRequest(BaseModel):
    sharepoint_url: str = Field(
        ...,
        description=(
            "URL de la lista tal como aparece en el navegador, p. ej. "
            "https://host.sharepoint.com/sitio/Lists/MiLista/AllItems.aspx"
        ),
    )
    data: dict[str, Any] = Field(
        ...,
        description=(
            "Campos del nuevo ítem. Las claves deben ser los nombres internos "
            "(internal name) de las columnas de la lista."
        ),
    )


class ListItemByUrlResponse(BaseModel):
    status: str = "created"
    id: str
    webUrl: str | None = None
    site_id: str
    list_id: str


class FilterBy(BaseModel):
    field: str = Field(
        ...,
        pattern=r"^[A-Za-z0-9_]+$",
        description=(
            "Nombre interno (internal name) de la columna que identifica de forma "
            "única el registro, p. ej. '_x006c_dq4'. Solo letras, dígitos y '_'."
        ),
    )
    value: str = Field(
        ...,
        description="Valor exacto que debe tener `field` en el registro a actualizar.",
    )


class ListItemUpdateByUrlRequest(BaseModel):
    sharepoint_url: str = Field(
        ...,
        description=(
            "URL de la lista tal como aparece en el navegador, p. ej. "
            "https://host.sharepoint.com/sitio/Lists/MiLista/AllItems.aspx"
        ),
    )
    filter_by: FilterBy = Field(
        ...,
        description=(
            "Campo único y valor con los que se localiza el registro a actualizar. "
            "Debe identificar un único ítem (si coincide más de uno se devuelve 409)."
        ),
    )
    data: dict[str, Any] = Field(
        ...,
        description=(
            "Campos a actualizar. Las claves deben ser los nombres internos "
            "(internal name) de las columnas de la lista."
        ),
    )


class ListItemUpdateByUrlResponse(BaseModel):
    status: str = "updated"
    id: str
    webUrl: str | None = None
    site_id: str
    list_id: str


class ExplicitPeriod(BaseModel):
    """Rango de fechas explícito para acotar el periodo del upsert.

    Ambos extremos son fechas ``YYYY-MM-DD``; ``to`` es **inclusive** (cubre todo
    ese día). Los límites se interpretan en la zona horaria del tenant.
    """

    model_config = ConfigDict(populate_by_name=True)

    from_: str = Field(
        ...,
        alias="from",
        description="Fecha inicial inclusive (YYYY-MM-DD), en la zona del tenant.",
    )
    to: str = Field(
        ...,
        description="Fecha final inclusive (YYYY-MM-DD), en la zona del tenant.",
    )


class UpsertMatch(BaseModel):
    """Criterio de coincidencia del upsert: clave y, opcionalmente, periodo."""

    key_field: str = Field(
        ...,
        pattern=r"^[A-Za-z0-9_]+$",
        description=(
            "Nombre interno de la columna clave (variable por lista). Solo letras, "
            "dígitos y '_'."
        ),
    )
    key_value: str = Field(
        ..., description="Valor exacto que debe tener la columna clave."
    )
    date_field: str | None = Field(
        default=None,
        pattern=r"^[A-Za-z0-9_]+$",
        description=(
            "Nombre interno de la columna de fecha que acota el periodo (p. ej. "
            "'Created', 'Modified' o una columna propia). Opcional."
        ),
    )
    period: str | ExplicitPeriod | None = Field(
        default=None,
        description=(
            "Alcance temporal. Atajo con nombre ('current_day', 'current_week', "
            "'current_month', 'current_year') o rango explícito {'from','to'}. "
            "Requiere `date_field`. Si se omite, la coincidencia es solo por clave."
        ),
    )

    @model_validator(mode="after")
    def _period_requires_date_field(self) -> "UpsertMatch":
        if self.period is not None and not self.date_field:
            raise ValueError(
                "Para acotar por periodo debe indicarse también `date_field`."
            )
        return self


class UpsertListItemRequest(BaseModel):
    sharepoint_url: str = Field(
        ...,
        description=(
            "URL de la lista tal como aparece en el navegador, p. ej. "
            "https://host.sharepoint.com/sitio/Lists/MiLista/AllItems.aspx"
        ),
    )
    match: UpsertMatch = Field(
        ..., description="Criterio para localizar el registro existente (clave + periodo)."
    )
    data: dict[str, Any] = Field(
        ...,
        description=(
            "Campos a escribir (crear o actualizar). Las claves deben ser los "
            "nombres internos de las columnas de la lista."
        ),
    )


class UpsertListItemResponse(BaseModel):
    result: Literal["created", "updated"]
    id: str
    webUrl: str | None = None
    site_id: str
    list_id: str
    matched: int = Field(
        default=0,
        description="Número de registros que coincidieron con el criterio.",
    )


class UploadByUrlResponse(BaseModel):
    status: str = "uploaded"
    id: str
    name: str
    size: int | None = None
    webUrl: str
    site_id: str
    drive_id: str
    folder: str
