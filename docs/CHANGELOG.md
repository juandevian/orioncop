# Changelog de orioncop

## Release actual: `v17.39.452.1452`

**Fecha de referencia:** 2026-09-14  
**Proyecto principal ejecutable:** `OrionCopIU`  
**Línea base técnica:** `17.39.452.1452`

## Resumen

Esta release fija la versión de referencia del módulo principal de `orioncop` a la versión `v17.39.452.1452` para documentar la línea base operativa y la trazabilidad del software que se ejecuta en el repositorio.

## Evidencia técnica

- `OrionCopIU/My Project/AssemblyInfo.vb` -> `AssemblyVersion("17.39.452.1452")`
- `OrionCopIU/mOrionCopIU.vb` -> `GstrVersionApp = My.Application.Info.Version.ToString`
- Esta versión corresponde al ensamblado principal de la interfaz de usuario y es la que se usa como referencia para la aplicación en runtime.

## Versiones asociadas en la solución

| Proyecto | Versión |
| --- | --- |
| `OrionCopIU` | `17.39.452.1452` |
| `OrionCopL` | `17.52.352.1452` |
| `RepOriCop` | `7.30.177.1448` |
| `OriIntCon` | `5.21.119.1446` |

## Trazabilidad de cambios para esta versión

1. Se revisaron los `AssemblyInfo.vb` de la solución para confirmar la versión del ensamblado principal.
2. Se validó que la aplicación toma la versión en runtime desde `My.Application.Info.Version.ToString` en `OrionCopIU/mOrionCopIU.vb`.
3. Se actualizó la documentación canónica de `docs` para reflejar la base técnica actual y el criterio de release.
4. Se registró la línea base en `docs/PLAN_MAESTRO.md`, `docs/ESTRATEGIA_REPOSITORIOS_GITHUB.md`, `docs/PLAN_ORIONCOP_V1.md` y `docs/HOJA_DE_RUTA_ORIONCOP_V1.md`.

## Nota estratégica

El repositorio `orioncop` no tiene un único `AssemblyVersion` unificado para todos los proyectos. La versión de referencia del producto ejecutable queda establecida en `v17.39.452.1452` para la capa de interfaz y la línea base operativa del módulo principal.
