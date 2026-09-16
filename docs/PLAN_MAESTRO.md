# Plan Maestro de orioncop

> **Estado:** vigente para ejecución inicial (2026/06/07)  
> **Ámbito:** solo este repositorio (`orioncop`)  
> **Documento complementario:** `docs/ESTRATEGIA_REPOSITORIOS_GITHUB.md`  
> **Versión base operativa:** `v17.39.454.1455`  
> **Proyecto principal de ejecución:** `OrionCopIU`  
> **Evidencia:** `OrionCopIU/My Project/AssemblyInfo.vb` declara `AssemblyVersion("17.39.454.1455")`; `OrionCopIU/OrionCopIU.vbproj` usa `ApplicationVersion>17.39.454.1455</ApplicationVersion>`.

---

## 1. Propósito

Este plan define **qué ejecutar** en `orioncop` para estabilizar el módulo principal de operación del negocio con bajo riesgo de regresión.

---

## 2. Objetivo de 5 días

Dejar `orioncop` en estado "estable y gobernable" para iteraciones semanales.

Resultado esperado al día 5:

1. Build reproducible en CI.
2. Flujo de cambios por PR con trazabilidad completa.
3. Contratos cross-repo documentados en `docs`.
4. Releases pequeños y reversibles.

---

## 3. Alcance real del plan

Incluye:

1. Estabilización técnica local de `orioncop`.
2. Orden documental en `docs`.
3. Gobierno de cambios y releases.
4. Coordinación de impactos con `comunes`, `adminorion`, `orionpcorreo` y `orion-installer`.

No incluye:

1. Reescritura total del sistema.
2. Migración tecnológica completa en esta etapa.
3. Cambios funcionales grandes sin baseline técnico.

---

## 4. Líneas de trabajo

### Línea A. Estabilidad técnica

1. Definir configuraciones de build soportadas.
2. Checklist de build local y CI.
3. Registro y mitigación de fallas repetitivas.

### Línea B. Trazabilidad y gobierno

1. PR template con impacto funcional y técnico.
2. Convención de ramas/commits aplicada.
3. Política de tags y releases operativa.

### Línea C. Contratos cross-repo

1. Dependencias con `comunes` documentadas.
2. Entradas/salidas de configuración con `adminorion`.
3. Integración operativa con `orionpcorreo`.
4. Artefactos para `orion-installer`.

### Línea D. Base para evolución

1. Mapa de zonas frágiles.
2. Priorización de encapsulaciones de bajo riesgo.
3. Backlog técnico con estimación corta.

---

## 5. Fases y calendario sugerido

### Fase 1 (Día 1-2): Recuperar control

1. Validar pipeline mínimo.
2. Confirmar reglas de ramas y PR.
3. Consolidar documentos canónicos en `docs`.

### Fase 2 (Día 3-4): Estabilizar para iterar

1. Corregir puntos de falla recurrentes.
2. Formalizar contratos con repos relacionados.
3. Establecer rutina de release pequeño.

### Fase 3 (Día 5): Preparar aceleración

1. Ejecutar mejoras de desacoplamiento.
2. Medir impacto en velocidad y riesgo.
3. Definir siguiente plan.

---

## 6. Indicadores de seguimiento

1. % de builds exitosos en CI.
2. Tiempo promedio desde PR a merge.
3. Número de hotfix en `main`.
4. Cambios con impacto cross-repo documentado.
5. Tiempo de recuperación ante rollback.

---

## 7. Riesgos abiertos y mitigación

1. **Conocimiento tácito del sistema.**  
   Mitigación: decision log y checklist de cambios.
2. **Acoplamiento con librerías compartidas.**  
   Mitigación: versionado explícito y matriz de dependencias.
3. **Regresiones en cambios urgentes.**  
   Mitigación: PR corta, CI mínima y rollback por revert.

---

## 8. Criterio de éxito

Este plan se considera cumplido cuando:

1. `orioncop` soporta iteraciones semanales sin bloqueos sistemáticos.
2. Existe trazabilidad completa de cambios y releases.
3. Los impactos cross-repo se gestionan por contrato.
4. El siguiente trimestre puede enfocarse en mejora evolutiva.

---

## 9. Línea base técnica de la versión actual

La versión base actualmente documentada para la aplicación principal es `v17.39.454.1455`, equivalente al ensamblado `OrionCopIU`.

Trazabilidad de referencia:

- `OrionCopIU/My Project/AssemblyInfo.vb` -> `AssemblyVersion("17.39.454.1455")`
- `OrionCopIU/OrionCopIU.vbproj` -> `ApplicationVersion>17.39.454.1455</ApplicationVersion>`
- `OrionCopL/clsItemFactura.vb` -> cálculo de pagos incluye retenciones y reversos de retenciones.
- `OrionCopL/clsCentroUtilidadOriCop.vb` -> `IDCtaImpAsumidos` pasa a ser requerido cuando `ObjAutorizaEFacBln` es `True`.
- `OrionCopIU/winNotasCr.xaml.vb` -> sincronización de `txtValorDctoNuevo` con valor del ítem.
- `OrionCopIU/winCentroUtilidadOriCop.xaml.vb` -> marcada como obsoleta para eliminación futura.
- `OriIntCon/OriIntCon.vbproj` -> rutas `IntermediateOutputPath` convertidas a rutas relativas portables.

Esta versión representa la línea base operativa efectiva del módulo principal ejecutable del repositorio `orioncop`.
