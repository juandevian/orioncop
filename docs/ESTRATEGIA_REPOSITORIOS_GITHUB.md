# Estrategia de Repositorio GitHub para orioncop

> **Estado:** vigente para este repositorio  
> **Alcance:** solo `orioncop`  
> **Objetivo:** operar y evolucionar `orioncop` con bajo riesgo, trazabilidad alta y releases controlados.  
> **Versión base actual:** `v17.39.453.1454`  
> **Proyecto principal ejecutable:** `OrionCopIU`

---

## 1. Proposito de este repositorio

`orioncop` es el modulo principal de operacion del negocio dentro de la plataforma Orion.

### Lo que SI pertenece a `orioncop`

- UI y flujos principales de operacion.
- Casos de uso de negocio del modulo.
- Integraciones necesarias para ejecucion operativa.

### Lo que NO pertenece a `orioncop`

- Administracion operativa de `adminorion`.
- Evolucion de librerias compartidas (`comunes`).
- Logica especializada de correo de `orionpcorreo`.
- Orquestacion de instalacion/release final (`orion-installer`).

---

## 2. Relacion con otros repositorios

1. **`comunes`**: proveedor de componentes compartidos.
2. **`adminorion`**: proveedor/consumidor de configuraciones administrativas.
3. **`orionpcorreo`**: integracion para notificaciones y procesos de correo.
4. **`orion-installer`**: empaquetado de artefactos versionados.

Principio: `orioncop` evoluciona de forma independiente sin romper contratos compartidos sin versionado y plan de transicion.

---

## 3. Principios de decision

1. Continuidad operativa primero.
2. Cambios pequenos y reversibles.
3. Compatibilidad explicita ante impactos cross-repo.
4. Trazabilidad total: issue -> rama -> PR -> tag -> release.

---

## 4. Flujo de ramas y cambios

- `main`, `develop`, `feature/*`, `hotfix/*`, `release/*`

Reglas:

1. No commit directo a `main`.
2. Todo cambio por PR.
3. PR con descripcion de impacto tecnico y operativo.
4. Impacto cross-repo documentado en `docs`.

---

## 5. Versionado y releases

Semver (`vX.Y.Z`):

- `X`: cambio mayor o ruptura.
- `Y`: mejora compatible.
- `Z`: correccion.

Cada release debe indicar:

1. Version de `comunes` esperada.
2. Impactos en `adminorion`/`orionpcorreo` (si aplica).
3. Artefacto para `orion-installer`.

### Línea base actual documentada

La línea base actual efectiva del repositorio queda fijada en `v17.39.453.1454` para el módulo principal ejecutable `OrionCopIU`.

Trazabilidad:

- `OrionCopIU/My Project/AssemblyInfo.vb` -> `AssemblyVersion("17.39.453.1454")`
- `OrionCopIU/OrionCopIU.vbproj` -> `ApplicationVersion>17.39.453.1454</ApplicationVersion>`
- `OrionCopL/clsSectorModulo.vb` -> validación de `TasaContribucion` ajustada a máximo `2`
- `OriIntCon` -> ajuste de dependencia `Newtonsoft.Json` a `13.0.3`
- `RepOriCop/RepOriCop.vbproj` -> eliminación de `mDefPubRepOrion.vb`

---

## 6. CI minima obligatoria

1. Restore de dependencias.
2. Compilacion.
3. Validaciones estaticas disponibles.
4. Empaquetado de artefacto (si aplica).

---

## 7. Rollback

1. Preferir `git revert`.
2. Generar release correctiva de parche.
3. Actualizar notas de release con causa/mitigacion.
4. Sin reescritura de historia publicada de `main`.

---

## 8. Decision estrategica vigente

Para `orioncop` se adopta:

1. Gestion independiente alineada al modelo multi-repo.
2. Integracion controlada con `comunes`, `adminorion`, `orionpcorreo` y `orion-installer`.
3. Releases pequenos, trazables y reversibles.
4. Prioridad en estabilidad operativa del modulo principal.

---

## 9. Traza de la versión actual

La versión base de referencia para esta etapa es `v17.39.453.1454`.

Se documenta con evidencia en `AssemblyInfo.vb` y `vbproj` del módulo principal, además del cambio funcional y de dependencias que acompaña a la release.
