# Changelog de orioncop

## En desarrollo (sin liberar): herramienta IBC (interés bancario corriente)

Estado: en integración (PR de `orioncop` y `comunes`). Requiere el esquema de BD 268 de `comunes`.

- Consulta de los certificados del interés bancario corriente (IBC) de la Superfinanciera desde la API propia (`OriIntCon/clsIbcApiCliente.vb`) y almacenamiento local en la tabla `OriIbcCertificados`.
- Parametrización del interés de mora por centro de utilidad: **Fijo** o **Variable** (IBC x factor definido por el usuario, máximo 1,5). Modo *Sin parametrizar* (por defecto) mantiene el comportamiento anterior.
- Tope legal de 1,5 x IBC (Ley 675 de 2001, art. 30): no se puede parametrizar un interés superior; la tasa fija deseada se conserva y la tasa efectiva se recalcula con cada IBC.
- El Cierre de mes actualiza `OriTasasMora` antes de causar intereses; si no se puede (API caída y sin certificado local vigente) no se cierra el mes ni se causan intereses.
- Herramientas > *Parametrizar interés de mora (IBC)* y *Consulta IBC*.
- Pruebas unitarias MSTest en `tests/OrionCopL.Tests` (fuera de la solución).
- Archivos principales: `OrionCopL/clsIbcCalculo.vb`, `clsIbcSincroniza.vb`, `clsIbcCertificado.vb`, `clsIbcTasasMoraBd.vb`, `clsTasaMora.vb` (`SRegistreTasaAnual`), `clsCentroUtilidadOriCop.vb`, `clsOrionCop.vb`, `OrionCopIU/WinParametrizacionInteres.xaml`, `WinConsultaIbc.xaml`, `winCausaMora.xaml.vb`.

## Release actual: `v17.39.454.1455`

**Fecha de referencia:** 2026-09-16  
**Proyecto principal ejecutable:** `OrionCopIU`  
**Línea base técnica:** `17.39.454.1455`

## Resumen

Esta release eleva la versión del ensamblado principal de `OrionCopIU` a `17.39.454.1455` e incluye mejoras funcionales en cálculo de pagos y validaciones de e-factura, ajustes de UI para notas crédito y una mejora de portabilidad para compilar correctamente después de clonar el repositorio.

## Evidencia técnica

- `OrionCopIU/My Project/AssemblyInfo.vb` -> `AssemblyVersion("17.39.454.1455")`
- `OrionCopIU/OrionCopIU.vbproj` -> `ApplicationVersion>17.39.454.1455</ApplicationVersion>`
- `OrionCopL/clsItemFactura.vb` -> inclusión de retenciones y reversos en el cálculo de pagos.
- `OrionCopL/clsCentroUtilidadOriCop.vb` -> cuenta `IDCtaImpAsumidos` obligatoria cuando aplica autorización de e-factura.
- `OrionCopIU/winNotasCr.xaml.vb` -> `txtValorDctoNuevo` se inicializa con el valor del ítem.
- `OrionCopIU/winCentroUtilidadOriCop.xaml.vb` -> clase marcada como obsoleta para retiro en futuras versiones.
- `OriIntCon/OriIntCon.vbproj` -> `IntermediateOutputPath` cambia de ruta local temporal a rutas relativas `obj/*`.

## Versiones asociadas en la solución

| Proyecto | Versión |
| --- | --- |
| `OrionCopIU` | `17.39.454.1455` |
| `OrionCopL` | `17.52.354.1455` |
| `RepOriCop` | `7.30.178.1454` |
| `OriIntCon` | `5.21.119.1446` |

## Cambios reales incluidos

1. Se ajustó `ClsItemFactura` para incorporar retenciones y reversos de retenciones en el valor de pago acumulado.
2. Se ajustó `ClsCentroUtilidadOriCop` para exigir la cuenta de impuestos asumidos cuando la e-factura está autorizada.
3. Se completó la carga del valor de descuento nuevo en la UI de notas crédito.
4. Se incrementó la versión de ensamblados (`OrionCopIU`, `OrionCopL`) y versión de publicación de `OrionCopIU`.
5. Se normalizaron rutas intermedias en `OriIntCon` para evitar dependencias a rutas temporales del equipo local.
6. Se marcó `WinCentroUtilidadOriCop` como obsoleta para migración y eliminación controlada en próximas versiones.

## Nota estratégica

La solución `orioncop` mantiene versionado por proyecto. La referencia operativa principal para ejecución y release queda en `OrionCopIU` versión `v17.39.454.1455`.
