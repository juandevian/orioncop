# Guía: generar instalador `.exe` con Inno Setup para OrionPlus

## Objetivo
Generar un instalador ejecutable QA (`OrionPlus-QA-Setup.exe`) que instale OrionPlus en:

`C:\Program Files\OptimuSoft\`

## Archivos agregados
1. `scripts\installer\OrionPlusInstaller.iss`
2. `scripts\build-installer.ps1`

## Requisitos
1. Tener compilado OrionPlus en `build\releases\orion-plus` (o dejar que el script lo compile).
2. Tener Inno Setup 6 instalado (`ISCC.exe`).

## Opción rápida (recomendada)
Desde la raíz del repositorio:

```powershell
.\scripts\build-installer.ps1
```

Esto ejecuta:
1. `.\scripts\build.ps1` para generar runtime completo en `build\releases\orion-plus`.
2. `ISCC.exe` sobre `scripts\installer\OrionPlusInstaller.iss`.
3. Producción del instalador QA en `build\installer\qa\OrionPlus-QA-Setup.exe`.

## Opción manual
1. Construir runtime:
   ```powershell
   .\scripts\build.ps1
   ```
2. Compilar instalador:
   ```powershell
   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" .\scripts\installer\OrionPlusInstaller.iss
   ```

## Resultado esperado
1. Instalador QA: `.\build\installer\qa\OrionPlus-QA-Setup.exe`
2. Instalación por defecto en:  
   `C:\Program Files\OptimuSoft\`

## Notas
1. Este flujo Inno es secundario y orientado a QA.
2. El instalador requiere permisos de administrador.
3. Se crea acceso directo en menú inicio y opcional en escritorio.
