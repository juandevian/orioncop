---
name: build-orionplus
description: Esta skill ejecuta el build nativo por defecto de OrionPlus y valida el runtime en `build\releases\orion-plus`; mantiene Inno como flujo secundario QA.
version: 1.0.0
last_updated: 2026-06-08
status: Activo
target_agent: Agente AVV
category: Build / CI-CD / Operación Técnica
repository: juandevian/orioncop
---

## Objetivo
1. Ejecutar por defecto el build nativo (`scripts\build.ps1`) y producir `OrionPlus.exe` en `build\releases\orion-plus`.
2. Capturar salida del script y validar artefacto + runtime requerido.
3. Mantener Inno Setup como opción secundaria para QA (`scripts\build-installer.ps1`).

## Prerrequisitos
- Repositorio `orioncop` clonado junto con `..\Comunes`.
- NuGet y MSBuild disponibles.
- Estructura local con solución `.\OrionCop.Net.sln`.

## Flujo recomendado
1. Build nativo por defecto:
   ```powershell
   .\scripts\build.ps1
   ```
2. Verificar salida esperada:
   - `.\build\releases\orion-plus\OrionPlus.exe`
   - DLL/config runtime en `.\build\releases\orion-plus\`
3. Opcional QA (Inno):
   ```powershell
   .\scripts\build-installer.ps1
   ```
   - Instalador QA: `.\build\installer\qa\OrionPlus-QA-Setup.exe`

## Criterios de éxito
- Compilación completa sin errores.
- `OrionPlus.exe` presente en la ruta de salida esperada.
- Trazabilidad del cambio documentada en PR (impacto cross-repo y rollback).