# Guía de Desarrollo para Agentes AI (AGENTS.md)

Este documento contiene las reglas y la metodología de trabajo que cualquier Agente de Inteligencia Artificial debe seguir al analizar, planificar o codificar dentro de este repositorio: **SharePoint Connector**, un microservicio REST en **Python + FastAPI** que reemplaza flujos de Power Automate y opera sobre SharePoint a través de **Microsoft Graph API**.

> El servicio es **solo backend**: no hay frontend en este repositorio. Expone una API HTTP versionada (`/v1/...`) y se despliega como contenedor.

## 1. Patrones Arquitectónicos y Estilo
- **Arquitectura por capas:** Mantener la separación que ya existe en `app/`:
  1. **Endpoints / Routers** (`app/api/v1/endpoints/`) — puntos de entrada HTTP. Validan el contrato (schemas Pydantic) y delegan; no contienen lógica de negocio ni llamadas HTTP directas a Graph.
  2. **Services / Lógica** (`app/services/`) — orquestación y acceso a la API externa. `SharePointService` (cliente de Graph), `SharePointResolver` (URL → IDs de Graph) y módulos de apoyo como `period.py`.
  3. **Core** (`app/core/`) — infraestructura transversal: configuración, autenticación, logging, contexto, excepciones, dependencias.
  4. **Schemas** (`app/schemas/`) — modelos Pydantic de request/response, organizados por dominio.
- **OOP donde aporta:** la lógica con estado o dependencias (servicios, resolver, token manager) se encapsula en clases. Las utilidades puras (p. ej. resolución de periodos, saneo de paths) pueden ser funciones de módulo.
- **Inyección de Dependencias (DI):** Nunca instanciar servicios acoplados dentro de los endpoints. Las dependencias se definen como singletons en `app/core/dependencies.py` (`@lru_cache`) y se inyectan en los endpoints con `Depends(...)` de FastAPI. Esto además permite sustituirlas por dobles en los tests (`app.dependency_overrides`).
- **Guard Clauses:** Programación lineal y predecible. Devolver el control (`return`/`raise` temprano) ante condiciones inválidas o estados negativos al inicio de la función. Evitar el anidamiento profundo de condicionales.
- **Genericidad:** El conector es **dinámico y multi-site**. No quemar en código sites, listas, columnas ni esquemas concretos: todo identificador específico viaja en la petición (path params en `/v1/graph/...`, o el payload/URL en `/v1/sharepoint/...`).
- **Seguridad de entradas:** Toda entrada del caller que se interpole en una URL o `$filter` de Graph debe validarse y escaparse. Patrones ya establecidos que hay que respetar y reutilizar:
  - Nombres de columna interna validados contra `^[A-Za-z0-9_]+$` (anti-inyección OData), tanto en el schema (`pattern`, responde `422`) como en el servicio (defensa en profundidad, responde `400`).
  - Valores de filtro escapados según OData (comillas simples duplicadas) y expresión `$filter` completa percent-encodeada.
  - Rutas/nombres de archivo saneados contra `..`, separadores y caracteres de control antes de construir la URL de Graph.

## 2. Metodología de Implementación (SDD — Spec Driven Development)
Este proyecto usa la variante **Spec-Anchored**: la especificación es la fuente única de verdad. Ver `requirements/00_SDD_GUIDELINES.md`.
- **Las especificaciones primero:** Antes de implementar una funcionalidad nueva o un cambio de alto impacto, debe existir (o proponerse) un documento `requirements/SPEC-XXX_NombreDescriptivo.md` con sus 4 secciones (Contexto, Propuesta, Criterios de Aceptación, Bitácora de IA).
- **Satisfacer los Criterios de Aceptación:** La implementación se considera completa solo cuando todos los criterios de la SPEC se cumplen y hay tests que los cubren.
- **Evolución sin destrucción:** Si una SPEC ya implementada necesita enmiendas, no reescribir la historia original; anexar al final del documento un bloque nuevo describiendo el cambio.

### Flujo al finalizar la implementación de una SPEC
1. **Versionado:** Actualizar el número de versión en el archivo **`VERSION`** (raíz del repo) siguiendo SemVer. Es la fuente única de verdad: `app/core/config.py` lo lee y `/health` lo reporta.
2. **Changelog:** Añadir una entrada nueva (agrupada bajo la versión) al inicio de **`doc/CHANGELOG.md`**, con contexto, solución, archivos nuevos/modificados.
3. **Estado de la SPEC:** Cambiar el encabezado de estado de la SPEC a `✅ Implementada` y **rellenar la sección de Bitácora de IA** con: fecha y rol, resumen de archivos modificados y para qué, decisiones clave y desviaciones respecto al diseño original.
4. **Documentación viva:** Mantener coherentes `README.md`, `ARQUITECTURA.md` y `TECNOLOGIAS.md` (tablas de endpoints, variables de entorno, número de versión) cuando el cambio los afecte.
5. **Tests:** Añadir/actualizar la suite `pytest` (`tests/`) y dejarla en verde antes de dar por cerrado el trabajo.

## 3. Logs y Observabilidad
- **Logging estructurado JSON:** El proyecto emite logs en JSON mediante `JSONFormatter` (`app/core/logging.py`). No introducir `print()` ni formatos de texto ad-hoc.
- **Logger por módulo:** Usar `logger = logging.getLogger(__name__)` en cada módulo.
- **Campos estructurados vía `extra`:** Adjuntar contexto con `extra={...}` usando únicamente los campos de la whitelist de `_EXTRA_FIELDS` (`request_id`, `client_app_id`, `method`, `path`, `status_code`, `duration_ms`, `site_id`, `list_id`, `drive_id`, `item_id`, `file_name`, `graph_url`, `graph_status`). Si se necesita un campo estructurado nuevo, añadirlo a esa whitelist.
- **Contexto de petición:** `request_id` y `client_app_id` se propagan por toda la pila mediante `ContextVar` (`app/core/context.py`), inicializados en el middleware de `app/main.py`. Recupéralos con el helper `_ctx()` de cada módulo cuando registres dentro de un servicio.
- **Errores controlados:** Los fallos de Graph se traducen a `GraphAPIError` tipado (`app/core/exceptions.py`) con su `status_code`; los handlers registrados en `main.py` los convierten en respuestas JSON consistentes con `request_id`. Un error en una integración no debe romper el flujo de forma silenciosa: captúralo, regístralo y propágalo tipado.

## 4. Control de Entorno y Servidor
- **Dependencias declaradas:** Toda librería nueva debe declararse en `requirements.txt` (runtime) o `requirements-dev.txt` (desarrollo/tests). Nunca instalar paquetes en el Python global del host: se asume ejecución en `venv` o en el contenedor Docker (`devops/`).
- **Configuración por entorno:** Los parámetros configurables se definen en `Settings` (`app/core/config.py`, pydantic-settings) y se documentan en `devops/.env.example`. No hardcodear credenciales, rutas personales ni valores cambiantes en el código.
- **Despliegue:** El servicio se conteneriza (`devops/Dockerfile`, `docker-compose.yml`). Si añades un archivo que deba existir en runtime (p. ej. `VERSION`), asegúrate de que el `Dockerfile` lo copie.
