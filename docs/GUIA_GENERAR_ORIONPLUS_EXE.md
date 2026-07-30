# Guía paso a paso: compilación nativa por defecto de `OrionPlus.exe`

## Propósito
Generar un binario compilable de OrionCop con dependencias de `..\Comunes`, entregando como resultado `OrionPlus.exe` en plataforma `Release|x64`.
Este es el flujo principal por defecto.

## Referencias de gobierno
- `docs/PLAN_ORIONCOP_V1.md`
- `docs/ESTRATEGIA_REPOSITORIOS_GITHUB.md`
- `.github/PLAN_MAESTRO.md` (puntero histórico)
- `.github/ESTRATEGIA_REPOSITORIOS_GITHUB.md` (puntero histórico)

## Qué se implementó en el repositorio
1. Se mantiene `AssemblyName` como `OrionCopIU`.
2. Se agrega un target MSBuild en `OrionCopIU\OrionCopIU.vbproj` que copia el runtime completo de salida a `build\releases\orion-plus` y además genera `OrionPlus.exe`.
3. CI compila explícitamente en `Release|x64` y publica artefacto nativo desde `build\releases\orion-plus`.

## Paso a paso manual (local)
1. Abrir terminal en la raíz del repo:
   ```powershell
   cd C:\FuentesPanorama.Net\Trunk\orioncop
   ```
2. Restaurar paquetes:
   ```powershell
   nuget restore .\OrionCop.Net.sln
   ```
3. Compilar solución completa:
   ```powershell
   "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" .\OrionCop.Net.sln /p:Configuration=Release /p:Platform=x64 /m /nologo
   ```
4. Confirmar salida del ejecutable final:
   - `.\build\releases\orion-plus\OrionPlus.exe`
   - En la misma carpeta deben quedar también las DLL/config necesarias del runtime.

## Automatización local sin agentes
1. Ejecutar script:
   ```powershell
   .\scripts\build\build.ps1
   ```
2. Opciones útiles:
   ```powershell
   .\scripts\build\build.ps1 -Configuration Release -Platform x64
   .\scripts\build\build.ps1 -SkipRestore
   .\scripts\build\build.ps1 -MsBuildPath "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
   ```

## Verificación rápida
1. Verificar existencia del archivo:
   ```powershell
   Test-Path .\build\releases\orion-plus\OrionPlus.exe
   ```
2. (Opcional) Ver hora de modificación:
   ```powershell
   Get-Item .\build\releases\orion-plus\OrionPlus.exe | Select-Object FullName, Length, LastWriteTime
   ```

## Impacto cross-repo
- No se modifican repositorios externos en esta entrega.
- `orion-installer` debe consumir el nuevo artefacto `OrionPlus.exe` según su proceso de empaquetado.

## Nota de ejecución
Ejecuta siempre el archivo desde `build\releases\orion-plus\OrionPlus.exe` conservando los archivos acompañantes de esa carpeta.  
Si se copia solo el `.exe` a otra ubicación sin DLL/config, puede cerrarse al iniciar sin mostrar login.

## Rollback
1. Revertir commit que agrega el target `CreateOrionPlusExecutable`.
2. Revertir ajustes de `.github/workflows/ci.yml`.
3. Recompilar con el flujo anterior.
