# SPEC-003: Upsert de ítems de lista (verificar-y-decidir)

> **Estado:** Aprobada — lista para implementar. Decisiones de negocio cerradas;
> quedan puntos técnicos a validar en implementación (ver Riesgos). Código aún
> no escrito.

## 1. Contexto (El Problema)

Un servicio externo dispara hoy un flujo de **Power Automate** que inyecta
registros en una lista de SharePoint. Ese flujo no hace un alta ciega: primero
**verifica** si ya existe un registro equivalente y, según el resultado,
**actualiza** el existente o **crea** uno nuevo. Se busca migrar esa lógica de
Power Automate a este conector (código), manteniendo dos servicios desacoplados:
uno que produce el JSON y otro (este conector) que lo materializa en SharePoint.

> **Alcance: funcionalidad genérica, NO atada a una lista concreta.**
> Esta característica es **reutilizable por cualquier lista de SharePoint**, no
> exclusiva del caso ticket/issue que la motiva. Distintas listas tienen
> **estructuras de datos diferentes** (columnas, nombres internos y campos clave
> distintos), y todas deben poder usar el mismo endpoint. Por tanto, **nada del
> dominio "issue" se quema en el código**: ni nombres de columna, ni el campo
> clave, ni el campo de fecha, ni el conjunto de datos. Todo parámetro específico
> de la lista viaja en el *payload* de cada llamada. El flujo de tickets descrito
> abajo es **solo el ejemplo motivador**, no el límite de la funcionalidad.

**Ejemplo motivador — comportamiento del flujo de tickets actual (alto nivel):**

1. Recibe vía HTTP los datos de un *issue* (organización, responsable,
   resolución, identificador de ticket, etc.).
2. Busca en la lista un registro que coincida por **dos condiciones a la vez**:
   - el **identificador de ticket** es igual, y
   - el registro pertenece al **mes en curso**.
3. Si **no** existe coincidencia → **crea** un ítem nuevo con todos los datos.
4. Si **existe** → **actualiza** ese registro con los datos recibidos.

Otras listas reutilizarán el mismo mecanismo con **su propia clave, su propio
campo de fecha y su propia estructura de datos** (p. ej. una lista de activos
identificada por `NumeroSerie`, una de contratos por `CodigoContrato`, etc.).

**Restricciones de dominio que condicionan el diseño:**

- **El campo clave es variable por lista.** No se puede quemar un nombre de
  columna concreto: cada lista identifica su registro con su propia columna
  (nombre interno distinto). → El campo de coincidencia viaja en el *payload*.
- **El campo de fecha también es variable.** La columna que acota el periodo
  puede ser la **fecha de creación**, la de **modificación**, o una columna de
  fecha propia de la lista. Tampoco se puede quemar. → El campo de fecha usado
  para acotar el periodo viaja en el *payload*.
- **La estructura de datos (`data`) es libre y propia de cada lista.** El conector
  no asume ningún esquema; reenvía a SharePoint las columnas que reciba (por
  nombre interno), sea cual sea la lista destino.
- La semántica de "mes en curso" es opcional y parametrizable: cuando aplica,
  el mismo valor clave puede generar un **registro nuevo cada periodo** (p. ej.
  cadencia mensual); dentro del mismo periodo se actualiza, en uno nuevo se crea.

## 2. Propuesta (La Solución)

Implementar la verificación **en este conector** (no en el servicio productor del
JSON), porque es quien conoce el estado real de la lista y debe ser la única
autoridad sobre idempotencia. El servicio productor permanece sin cambios.

Frente a las dos vías evaluadas — (1) añadir un *flag* al endpoint de inserción
actual, o (2) crear una API nueva genérica de inserción — se opta por una
**tercera vía: un endpoint de _upsert_ explícito** con semántica propia de
"insertar-o-actualizar". Razones:

- Evita el anti-patrón *flag-argument* (un mismo endpoint con dos contratos según
  un booleano), que ensucia pruebas, documentación y la API ya estable.
- No duplica lógica de escritura: **reutiliza los primitivos ya existentes** del
  `SharePointService` y la capa de resolución por URL de SPEC-002.
- Expresa la intención real (idempotencia), no "otra inserción".

**Endpoint propuesto (capa orientada a usuario, por URL — consistente con SPEC-002):**

```
POST /v1/sharepoint/list/item:upsert
```

**Cuerpo (borrador):**

```jsonc
{
  "sharepoint_url": "https://host.sharepoint.com/Oper/Lists/Incidencias/AllItems.aspx",
  "match": {
    "key_field": "TicketId",        // nombre interno; variable por lista
    "key_value": "INC-12345",
    "date_field": "Created",         // nombre interno del campo de fecha; variable (Created | Modified | columna propia)
    "period": "current_month"        // alcance temporal; opcional. Ver formas abajo.
  },
  "data": { /* campos a escribir; claves = nombres internos */ }
}
```

`period` admite un **atajo con nombre** (`"current_day"`, `"current_week"`,
`"current_month"`, `"current_year"`, …) o un **rango explícito**:

```jsonc
"period": { "from": "2026-01-01", "to": "2026-06-30" }
```

Si se omite `period` (o `date_field`), la coincidencia es **solo por clave**.

**Decisiones de comportamiento (acordadas con el negocio):**

| # | Decisión | Acuerdo |
|---|----------|---------|
| 1 | ¿Campo clave fijo o variable? | **Variable por lista** → va en el payload (`match.key_field`). |
| 2 | ¿Qué hacer si la búsqueda devuelve >1 registro? | **Actualizar el primero** y registrar un `warning` (no se devuelve 409, a diferencia del PATCH actual). El warning permite detectar duplicados preexistentes. |
| 3 | ¿Distinguir creación de actualización? | **Sí**, para registrarlo en los logs y devolverlo en la respuesta (`result: "created" \| "updated"`). |

**Orquestación (sin escritura nueva de bajo nivel):**

1. Resolver `sharepoint_url` → `site_id` + `list_id` (resolver de SPEC-002).
2. Buscar coincidencias por **clave + mes en curso** sobre el `date_field` indicado.
3. Si **0** → `create_list_item` → responder `result="created"`.
4. Si **≥1** → tomar el **primero**, `update_list_item` → responder `result="updated"`.
5. Loguear en ambos ramos el desenlace (`created`/`updated`), el `id` y la clave.

**Regla del filtro de coincidencia (decisión de diseño):**

La búsqueda combina **dos condiciones** sobre la lista, unidas con `and`:

1. **Campo clave** — igualdad exacta: `fields/{key_field} eq '{key_value}'`.
2. **Periodo** — el `date_field` es una columna de **tipo Fecha**, por lo que se
   filtra por **rango**: `fields/{date_field} ge '{inicio}' and fields/{date_field} lt '{fin}'`.
   El periodo es **opcional**: si no se indica `date_field`, la coincidencia es
   solo por clave.

   El periodo **no se limita a "mes en curso"**. `period` admite:
   - **Rango explícito** — `from`/`to` proporcionados por el caller; el conector
     filtra ese intervalo tal cual. Es la forma más genérica y sirve a cualquier lista.
   - **Atajos con nombre** — valores cómodos que el conector traduce a un rango en
     la zona horaria del tenant: `current_day`, `current_week`, `current_month`,
     `current_year` (extensible). `current_month` es solo uno más, no el único.

`Created`/`Modified` son campos de fecha de sistema y se filtran igual por rango.

**Zona horaria del periodo (decisión):** los límites del "mes en curso" se
calculan en la **zona horaria del tenant de SharePoint** (no la del servidor ni
UTC), de modo que el corte de mes coincida con cómo el usuario ve las fechas en
SharePoint. Los valores enviados a Graph en el `$filter` se ajustan en
consecuencia (Graph almacena/compara en UTC, así que el rango local del tenant se
convierte a UTC antes de consultar).

**Riesgos y cuestiones abiertas (a resolver antes de implementar):**

- **Columnas no indexadas en listas grandes.** Si el `date_field` o el `key_field`
  no están indexados, Graph puede rechazar el `$filter` en listas con muchos miles
  de ítems. Ya se envía la cabecera `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly`
  (que Microsoft advierte puede fallar de forma intermitente). Recomendación
  operativa: **indexar** las columnas clave y de fecha en las listas que usen upsert.
- **Obtención de la zona horaria del tenant.** Confirmar cómo se obtiene de forma
  fiable (configuración del sitio vía Graph, o parámetro de configuración del
  conector) para hacer la conversión local→UTC del rango de fechas.
- **Concurrencia.** Dos *upserts* simultáneos de la misma clave podrían crear dos
  registros (no hay unicidad garantizada en la lista). Evaluar si es un riesgo real
  según la cadencia del servicio productor.
- **Forma del contrato.** Confirmar nombres definitivos de las claves del payload
  (`match`/`key_field`/`date_field`/`period`) y el catálogo final de atajos con
  nombre de `period`. La estructura ya está acordada: `period` admite rango
  explícito (`from`/`to`) **y** atajos con nombre.

## 3. Criterios de Aceptación

- [ ] El endpoint es **genérico**: funciona contra **cualquier lista** de SharePoint sin código específico de dominio. No hay nombres de columna, campo clave, campo de fecha ni esquema de datos quemados; todo viaja en el payload.
- [ ] Se demuestra con **al menos dos listas de estructura distinta** (claves y columnas diferentes) usando el mismo endpoint sin cambios de código.
- [ ] `POST /v1/sharepoint/list/item:upsert` crea el ítem si no existe coincidencia.
- [ ] Actualiza el ítem si existe coincidencia por **clave (+ periodo, si se indica)**.
- [ ] El **campo clave** se toma del payload (`match.key_field`), no está quemado.
- [ ] El **campo de fecha** que acota el periodo se toma del payload (`match.date_field`), admite al menos `Created` y `Modified`.
- [ ] El acotado por periodo es **opcional**: si no se indica `date_field`/`period`, la coincidencia es solo por clave.
- [ ] `period` admite **rango explícito** (`from`/`to`) **y** atajos con nombre (`current_day`, `current_week`, `current_month`, `current_year`), no solo `current_month`.
- [ ] Cuando se indica periodo, la coincidencia exige **ambas** condiciones simultáneamente (clave Y periodo).
- [ ] La estructura de `data` se reenvía tal cual a SharePoint (nombres internos), sin asumir esquema.
- [ ] Ante **múltiples** coincidencias, se actualiza el **primer** registro (sin error) y se emite un `warning` en los logs indicando cuántos registros coincidieron.
- [ ] La respuesta indica `result: "created"` o `"updated"`, e incluye `id`, `site_id`, `list_id`.
- [ ] El desenlace (`created`/`updated`) queda registrado en los logs estructurados.
- [ ] Los endpoints existentes (`POST` y `PATCH /v1/sharepoint/list/item`, y todo `/v1/graph/...`) permanecen sin cambios de comportamiento.
- [ ] El servicio productor del JSON no requiere cambios para la verificación.
- [ ] Suite de tests cubre los tres desenlaces: alta, actualización y empate múltiple.

## 4. Bitácora de IA / Historial de Implementación

_Pendiente. Aún no implementado. Cada agente o desarrollador que aborde esta
spec debe añadir aquí su entrada (fecha, rol, archivos modificados, decisiones
clave y desviaciones) al completar la implementación._
