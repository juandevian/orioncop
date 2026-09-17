# Estrategia de Repositorio GitHub para orioncop

> **Estado:** vigente para este repositorio  
> **Alcance:** solo `orioncop`  
> **Objetivo:** operar y evolucionar `orioncop` con bajo riesgo, trazabilidad alta y releases controlados.  
> **Versión base actual:** `v17.39.454.1455`  
> **Proyecto principal ejecutable:** `OrionCopIU`

---

## 1. Propósito de este repositorio

`orioncop` es el módulo principal de operación del negocio dentro de la plataforma Orion.

### Lo que SÍ pertenece a `orioncop`

- UI y flujos principales de operación.
- Casos de uso de negocio del módulo.
- Integraciones necesarias para ejecución operativa.

### Lo que NO pertenece a `orioncop`

- Administración operativa de `adminorion`.
- Evolución de librerías compartidas (`comunes`).
- Lógica especializada de correo de `orionpcorreo`.
- Orquestación de instalación/release final (`orion-installer`).

---

## 2. Relación con otros repositorios

1. **`comunes`**: proveedor de componentes compartidos.
2. **`adminorion`**: proveedor/consumidor de configuraciones administrativas.
3. **`orionpcorreo`**: integración para notificaciones y procesos de correo.
4. **`orion-installer`**: empaquetado de artefactos versionados.

Principio: `orioncop` evoluciona de forma independiente sin romper contratos compartidos sin versionado y plan de transición.

---

## 3. Principios de decisión

1. Continuidad operativa primero.
2. Cambios pequeños y reversibles.
3. Compatibilidad explícita ante impactos cross-repo.
4. Trazabilidad total: issue -> rama -> PR -> tag -> release.

---

## 4. Flujo de ramas y cambios

- `main`, `develop`, `feature/*`, `hotfix/*`, `release/*`

Reglas:

1. No commit directo a `main`.
2. Todo cambio por PR.
3. PR con descripción de impacto técnico y operativo.
4. Impacto cross-repo documentado en `docs`.

---

## 5. Versionado y releases

Semver (`vX.Y.Z`):

- `X`: cambio mayor o ruptura.
- `Y`: mejora compatible.
- `Z`: corrección.

Cada release debe indicar:

1. Versión de `comunes` esperada.
2. Impactos en `adminorion`/`orionpcorreo` (si aplica).
3. Artefacto para `orion-installer`.

### Línea base actual documentada

La línea base actual efectiva del repositorio queda fijada en `v17.39.454.1455` para el módulo principal ejecutable `OrionCopIU`.

Trazabilidad:

- `OrionCopIU/My Project/AssemblyInfo.vb` -> `AssemblyVersion("17.39.454.1455")`
- `OrionCopIU/OrionCopIU.vbproj` -> `ApplicationVersion>17.39.454.1455</ApplicationVersion>`
- `OrionCopL/clsItemFactura.vb` -> cálculo de pagos incluye retenciones y reversos.
- `OrionCopL/clsCentroUtilidadOriCop.vb` -> validación obligatoria de cuenta de impuestos asumidos cuando aplica e-factura.
- `OrionCopIU/winNotasCr.xaml.vb` -> actualización del campo `txtValorDctoNuevo`.
- `OrionCopIU/winCentroUtilidadOriCop.xaml.vb` -> marcado como obsoleto.
- `OriIntCon/OriIntCon.vbproj` -> rutas intermedias normalizadas para portabilidad en clonación.

---

## 6. CI mínima obligatoria

1. Restore de dependencias.
2. Compilación.
3. Validaciones estáticas disponibles.
4. Empaquetado de artefacto (si aplica).

---

## 7. Rollback

1. Preferir `git revert`.
2. Generar release correctiva de parche.
3. Actualizar notas de release con causa/mitigación.
4. Sin reescritura de historia publicada de `main`.

---

## 8. Decisión estratégica vigente

Para `orioncop` se adopta:

1. Gestión independiente alineada al modelo multi-repo.
2. Integración controlada con `comunes`, `adminorion`, `orionpcorreo` y `orion-installer`.
3. Releases pequeños, trazables y reversibles.
4. Prioridad en estabilidad operativa del módulo principal.

---

## 9. Traza de la versión actual

La versión base de referencia para esta etapa es `v17.39.454.1455`.

Se documenta con evidencia en `AssemblyInfo.vb` y `vbproj` del módulo principal, más cambios funcionales y de mantenibilidad orientados a operación estable y build reproducible al clonar el repositorio.
