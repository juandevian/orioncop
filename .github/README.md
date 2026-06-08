# Estandares de la carpeta .github

Este archivo define que debe vivir en `.github` dentro de `orioncop`.

## 1. Proposito de `.github`

`.github` es una carpeta de gobierno del repositorio y automatización; no es documentación canónica de producto.

## 2. Contenido permitido

1. Workflows de CI/CD en `.github/workflows/`.
2. Plantillas de PR/issues.
3. Configuración de seguridad y dependencias.
4. Instrucciones operativas para automatizaciones.
5. Agent necesarios para el código
6. Skills específicos para .net 4.7

## 3. Fuente canónica de documentación

1. `docs/PLAN_ORIONCOP_V1.md`
2. `docs/ESTRATEGIA_REPOSITORIOS_GITHUB.md`

Si se requiere referencia desde `.github`, usar enlaces y no duplicados.

## 4. Reglas de mantenimiento

1. Evitar duplicar documentos estratégicos en `.github`.
2. Actualizar primero `docs/` ante cambios de decision.
3. Mantener consistencia de nombres de componentes.

## 5. Plantilla de PR obligatoria

Se usa `pull_request_template.md` para asegurar:

1. evaluacion de impacto cross-repo;
2. evidencia de validación;
3. plan de rollback antes de merge.
