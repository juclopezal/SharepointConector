# spec-mas-de-un-campo

Prompt listo para lanzar la extensión del endpoint de búsqueda a filtros con varios campos.
Cuando quieras hacerla, copia el bloque siguiente y pégalo tal cual en Claude Code.

---

## Prompt

Quiero extender el endpoint `POST /v1/sharepoint/list/items:search` del SharePoint Connector
para poder filtrar por **más de un campo a la vez**. Hazlo con OpenSpec, actualizando el cambio
existente `add-list-items-read-by-url` (`/opsx:update`): primero revisa proposal, design, spec
y tasks, enséñame los artefactos y **espera mi aprobación antes de tocar código**.

Requisitos funcionales:

1. Nuevo campo opcional `filters` en el body: un array de condiciones
   `{ "field": "...", "value": ... }` que se combinan con **AND** en el `$filter` de Graph
   (generalizando lo que ya hace el upsert con clave + fecha).

   Ejemplo de petición:

   ```json
   {
     "sharepoint_url": "https://latinia2com-portal8.sharepoint.com/Oper/Lists/DSL/Allitemsg.aspx",
     "filters": [
       { "field": "Entorno", "value": "l01-2016" },
       { "field": "_x00da_ltima", "value": true }
     ],
     "top": 100
   }
   ```

2. Compatibilidad hacia atrás: el `filter_by` actual (un solo campo) debe seguir funcionando
   igual. Decidir en el design si `filters` lo sustituye (deprecándolo) o conviven, y qué pasa
   si llegan los dos a la vez.

3. Tratamiento de tipos en `value`: además de texto, soportar **booleanos** (columnas como
   `_x00da_ltima`, que hoy fallarían comparadas como texto) y valorar números. El design debe
   explicar cómo se traduce cada tipo al `$filter` de OData.

4. Mantener las garantías actuales: validación de nombres de campo (`[A-Za-z0-9_]+`,
   anti-inyección OData), sin exigir unicidad (0 coincidencias → `200` con `total: 0`),
   `top` 1–5000 con truncado, misma forma de respuesta.

5. Tests nuevos (varios campos, booleanos, retrocompatibilidad de `filter_by`, validaciones)
   y toda la suite existente en verde.

6. Actualizar documentación: README, ARQUITECTURA.md, arquitecturasUML.md y doc/CHANGELOG.md.

7. Al terminar la implementación aprobada: reconstruir y redesplegar el contenedor
   (`docker compose -f devops/docker-compose.yml up -d --build`) y verificar en
   `http://docker-ag:8001/health` y el OpenAPI que el cambio está desplegado.

Contexto útil: la petición se prueba desde Postman contra `http://localhsot:8001`
(POST + body raw JSON). La lista de pruebas es la DSL:
`https://latinia2com-portal8.sharepoint.com/Oper/Lists/DSL/Allitemsg.aspx`,
con campos como `Entorno` (texto), `_x00da_ltima` (booleano), `Title`, `Grupo`, `hohl`.
