# Plan v1 - orioncop

> **Estado:** activo  
> **Fecha:** 2026-06-07  
> **Objetivo macro:** aportar al nuevo sistema AI-first para automatizar procesos administrativos de Orion.

## 1. Objetivo del repo
Convertir `orioncop` en el núcleo funcional priorizado para migración híbrida: mantener continuidad legacy mientras se extraen reglas de negocio y flujos críticos hacia capacidades automatizables con AI.

## 2. Estado actual resumido
1. Modulo core de operación con alto volumen de código y alta dependencia de `comunes`.
2. Flujo actual sensible a despliegue manual y configuraciones de entorno.
3. Validaciones predominantemente manuales.

## 3. Alcance v1
1. Identificar y documentar 3 flujos críticos de negocio.
2. Definir contratos iniciales para extraer lógica reusable.
3. Preparar integración controlada con componentes nuevos sin romper operación.
4. Dejar base de eventos/datos para automatización AI asistida.

## 4. Backlog inmediato
1. Inventario de módulos y pantallas criticas por flujo.
2. Mapa de dependencias `orioncop -> comunes`.
3. Matriz regla legacy -> capacidad objetivo AI.
4. Smoke tests mínimos sobre flujos prioritarios.
5. Definición de puntos de inserción para automatización AI.

## 5. Riesgos y mitigación
1. **Acoplamiento alto con comunes** -> encapsular acceso y contratos.
2. **Conocimiento tácito** -> canon funcional por flujo.
3. **Interrupción operativa** -> despliegues fuera de horario + rollback.

## 6. Criterios de éxito v1
1. Flujos críticos identificados y trazados extremo a extremo.
2. Contratos iniciales definidos para migración incremental.
3. Evidencia de al menos un flujo con automatización AI asistida en paralelo.

