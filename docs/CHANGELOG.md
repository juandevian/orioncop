# Changelog de orioncop

## Release actual: `v17.39.453.1454`

**Fecha de referencia:** 2026-09-14  
**Proyecto principal ejecutable:** `OrionCopIU`  
**Línea base técnica:** `17.39.453.1454`

## Resumen

Esta release ajusta la versión del ensamblado principal de `OrionCopIU` a `17.39.453.1454` y documenta los cambios reales incluidos en la base operativa actual: validación de `TasaContribucion`, ajuste de dependencia `Newtonsoft.Json` en `OriIntCon` y eliminación de `mDefPubRepOrion.vb` del proyecto `RepOriCop`.

## Evidencia técnica

- `OrionCopIU/My Project/AssemblyInfo.vb` -> `AssemblyVersion("17.39.453.1454")`
- `OrionCopIU/OrionCopIU.vbproj` -> `ApplicationVersion>17.39.453.1454</ApplicationVersion>`
- `OrionCopL/clsSectorModulo.vb` -> validación `TasaContribucion` con máximo `2`
- `OriIntCon/OriIntCon.vbproj` + `OriIntCon/packages.config` -> `Newtonsoft.Json` ajustado a `13.0.3`
- `RepOriCop/RepOriCop.vbproj` -> eliminación de `mDefPubRepOrion.vb`

## Versiones asociadas en la solución

| Proyecto | Versión |
| --- | --- |
| `OrionCopIU` | `17.39.453.1454` |
| `OrionCopL` | `17.52.353.1454` |
| `RepOriCop` | `7.30.178.1454` |
| `OriIntCon` | `5.21.119.1446` |

## Cambios reales incluidos

1. Se corrigió la validación de `TasaContribucion` en `OrionCopL` para aceptar un máximo de `2` en lugar de `1`.
2. Se ajustó la dependencia `Newtonsoft.Json` en `OriIntCon` a la versión `13.0.3`.
3. Se eliminó la inclusión del archivo `mDefPubRepOrion.vb` del proyecto `RepOriCop`.
4. Se actualizó el número de versión del ensamblado principal `OrionCopIU` a `17.39.453.1454`.

## Trazabilidad de la release

1. Se revisaron los `AssemblyInfo.vb` y `vbproj` de la solución para confirmar la versión actual del ensamblado principal.
2. Se identificaron los cambios funcionales reales y de dependencia que acompaña la release.
3. Se actualizó la documentación canónica de `docs` para dejar la línea base técnica y de release consistente con el código actual.

## Nota estratégica

La solución `orioncop` no tiene un único `AssemblyVersion` unificado para todos los proyectos. La versión de referencia del producto ejecutable queda establecida en `v17.39.453.1454` para la capa principal de interfaz y la base operativa del repositorio.
