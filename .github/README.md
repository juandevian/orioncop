# Estandares de la carpeta .github

Este archivo define que debe vivir en `.github` dentro de `orioncop`.

## 1. Proposito de `.github`

`.github` es una carpeta de gobierno del repositorio y automatizacion; no es documentacion canonica de producto.

## 2. Contenido permitido

1. Workflows de CI/CD en `.github/workflows/`.
2. Plantillas de PR/issues.
3. Configuracion de seguridad y dependencias.
4. Instrucciones operativas para automatizaciones.

## 3. Fuente canonica de documentacion

1. `docs/PLAN_MAESTRO.md`
2. `docs/ESTRATEGIA_REPOSITORIOS_GITHUB.md`

Si se requiere referencia desde `.github`, usar enlaces y no duplicados.

## 4. Reglas de mantenimiento

1. Evitar duplicar documentos estrategicos en `.github`.
2. Actualizar primero `docs/` ante cambios de decision.
3. Mantener consistencia de nombres de componentes.

## 5. Plantilla de PR obligatoria

Se usa `pull_request_template.md` para asegurar:

1. evaluacion de impacto cross-repo;
2. evidencia de validacion;
3. plan de rollback antes de merge.
