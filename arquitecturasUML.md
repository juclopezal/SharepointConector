# Diagramas de Arquitectura (PlantUML): SharePoint Connector

**Versión:** 2.5.0
**Fecha:** 2026-07-22
**Autor:** Juan Camilo López Alzate — Latinia

> Versión en **PlantUML** de los diagramas de [`ARQUITECTURA.md`](ARQUITECTURA.md).
> Cada bloque puede renderizarse en [plantuml.com](https://www.plantuml.com/plantuml),
> con la extensión *PlantUML* de VS Code, o con el CLI `plantuml diagrama.puml`.

---

## 2. Arquitectura actual

```plantuml
@startuml arquitectura-actual
left to right direction
skinparam componentStyle rectangle

package "Caller (cualquier sistema)" {
  component "HTTP Client" as C
}

package "sharepoint-connector (Docker)" {
  component "Middleware\nRequest ID · X-App-ID" as MW
  component "API Layer\n/v1/graph/..." as API
  component "Dependencies\nlru_cache singletons" as DEP
  component "SharePointService\nGraph API client" as SRV
  component "TokenManager\nOAuth2 cache" as AUTH
  component "Settings\npydantic-settings" as CFG
}

package "Azure AD" {
  component "Token Endpoint\nclient_credentials" as TK
}

package "Microsoft Graph API v1.0" {
  component "/sites — Discovery" as GS
  component "/sites/{id}/lists — List Items" as GL
  component "/sites/{id}/drives — Files" as GF
}

package "SharePoint" as SP {
  database "Biblioteca\nDocumentos" as DL
  database "Lista" as LST
}

C --> MW : HTTP + X-App-ID
MW --> API
API --> DEP
DEP --> SRV
DEP --> AUTH
AUTH --> TK : client_credentials
TK --> AUTH : Bearer token
SRV --> GS
SRV --> GL
SRV --> GF
GS --> SP
GL --> LST
GF --> DL
CFG --> AUTH
CFG --> SRV
@enduml
```

---

## 3. Componentes internos

```plantuml
@startuml componentes-internos
skinparam componentStyle rectangle

package "app/api/v1/endpoints" {
  component "discovery.py\nGET /sites\nGET /sites/{id}/lists\nGET /sites/{id}/drives\nGET /sites/{id}/drives/{id}/items" as EP_DISC
  component "list_items.py\nGET  /sites/{id}/lists/{id}/items\nPOST /sites/{id}/lists/{id}/items" as EP_LIST
  component "files.py\nPOST /sites/{id}/drives/{id}/files\nGET  /sites/{id}/drives/{id}/items/{id}\nGET  /sites/{id}/drives/{id}/items/{id}/download" as EP_FILE
  component "sharepoint.py (by URL)\nPOST  /sharepoint/list/item\nPATCH /sharepoint/list/item\nPOST  /sharepoint/list/item:upsert\nPOST  /sharepoint/list/items:search\nPOST  /sharepoint/upload" as EP_SP
}

package "app/core" {
  component "auth.py\nTokenManager" as AUTH
  component "config.py\nSettings" as CFG
  component "context.py\nContextVars" as CTX
  component "dependencies.py\nget_sp()" as DEP
  component "exceptions.py\nGraphAPIError" as EXC
  component "logging.py\nJSONFormatter" as LOG
}

package "app/schemas" {
  component "discovery.py" as SCH_D
  component "files.py" as SCH_F
  component "list_items.py" as SCH_L
  component "sharepoint.py" as SCH_SP
}

package "app/services" {
  component "sharepoint.py\nSharePointService" as SRV
  component "resolver.py\nSharePointResolver" as RES
}

EP_DISC --> DEP
EP_LIST --> DEP
EP_FILE --> DEP
EP_SP --> DEP
EP_DISC --> SCH_D
EP_FILE --> SCH_F
EP_LIST --> SCH_L
EP_SP --> SCH_SP
EP_SP --> RES
RES --> SRV
DEP --> SRV
DEP --> AUTH
SRV --> CTX
SRV --> EXC
AUTH --> CFG
SRV --> CFG
@enduml
```

---

## 4. Capa de resolución de URL (`SharePointResolver`)

```plantuml
@startuml resolucion-url
left to right direction
skinparam componentStyle rectangle

component "URL SharePoint\n(navegador)" as URL
component "SharePointResolver" as R
component "site_id" as SID
component "list_id" as LID
component "drive_id + carpeta" as DID
component "SharePointService\ncreate_list_item / update_list_item /\nsearch_list_items (+ get_list_columns) / upload_file" as SRV

URL --> R
R --> SID : GET /sites/host:/path\nrecorte hacia raíz
R --> LID : match webUrl / nombre
R --> DID : match prefijo webUrl drive
SID --> SRV
LID --> SRV
DID --> SRV
@enduml
```

---

## 5. Autenticación y seguridad

```plantuml
@startuml autenticacion
participant Caller
participant Connector
participant "Azure AD" as AAD
participant "Microsoft Graph" as Graph

Caller -> Connector : HTTP request + X-App-ID
Connector -> Connector : Asigna request_id (UUID)
Connector -> Connector : ¿token en caché válido?
alt token expirado o no existe
    Connector -> AAD : POST /oauth2/v2.0/token\ngrant_type=client_credentials
    AAD --> Connector : access_token (1h TTL)
    Connector -> Connector : guarda token en memoria
end
Connector -> Graph : Llamada con Bearer token
Graph --> Connector : 200 / error
Connector --> Caller : respuesta + X-Request-ID
@enduml
```
