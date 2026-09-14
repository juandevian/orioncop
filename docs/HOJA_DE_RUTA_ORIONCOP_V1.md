# Hoja de Ruta v1 - orioncop

> **Estado:** vigente  
> **Documento base:** `docs/PLAN_ORIONCOP_V1.md`  
> **Versión base actual:** `v17.39.452.1452`  
> **Proyecto principal ejecutable:** `OrionCopIU`

## Baseline técnica actual

La referencia operativa del repositorio queda fijada en la versión `v17.39.452.1452` del ensamblado principal `OrionCopIU`.

Trazabilidad:

- `OrionCopIU/My Project/AssemblyInfo.vb` -> `AssemblyVersion("17.39.452.1452")`
- `OrionCopIU/mOrionCopIU.vb` -> `My.Application.Info.Version.ToString`
- Esta línea base se mantiene como despliegue principal del repositorio para la ejecución de la aplicación.

## Tramo 1 - Diagnostico funcional
**Entregables**
1. Mapa de flujos críticos y dependencias con `comunes`.
2. Canon de reglas legacy por flujo.
**Criterio de salida:** trazabilidad funcional validada.

## Tramo 2 - Estabilización minima
**Entregables**
1. Runbook técnico de build y despliegue actual.
2. Smoke tests de flujo critico.
**Criterio de salida:** ejecución repetible y menos frágil.

## Tramo 3 - Arquitectura puente
**Entregables**
1. Contratos para extraer lógica de negocio.
2. Definición de puntos de automatización AI.
**Criterio de salida:** backlog ejecutable para migración incremental.

## Tramo 4 - Implementación inicial
**Entregables**
1. Capacidades nuevas en paralelo con datos legacy.
2. Plan de rollback por flujo.
**Criterio de salida:** primer flujo critico con automatización AI asistida en operación controlada.

