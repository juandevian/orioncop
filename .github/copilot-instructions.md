# Contexto obligatorio del repositorio (orioncop)

Antes de proponer cambios, usa como fuente de verdad estos documentos locales:

1. `.github/PLAN_MAESTRO.md` (puntero historico)
2. `.github/ESTRATEGIA_REPOSITORIOS_GITHUB.md` (puntero historico)
3. `docs/PLAN_ORIONCOP_V1.md` (canonico)
4. `docs/ESTRATEGIA_ORIONCOP_V1.md` (canonico)

## Comandos de build, test y lint

Pipeline CI actual (`.github/workflows/ci.yml`):

```powershell
nuget restore .\OrionCop.Net.sln
msbuild .\OrionCop.Net.sln /p:Configuration=Release /m
```

Build local de un solo proyecto (para iterar mas rápido):

```powershell
msbuild .\OrionCopIU\OrionCopIU.vbproj /p:Configuration=Release /p:Platform=x86
```

Tests automatizados:

- No hay proyectos de pruebas ni paso `test` en CI actualmente.
- No existe comando de "single test" en este repositorio en su estado actual.

Lint / análisis estático:

- No hay comando de lint dedicado en CI.
- El análisis se ejecuta via `CodeAnalysisRuleSet` definido en los `.vbproj` (por ejemplo `OriIntCon.ruleset`, `MinimumRecommendedRules.ruleset`) durante build.
- `.editorconfig` eleva `BC42104` a error para archivos VB.

## Arquitectura de alto nivel (big picture)

- Solución principal: `OrionCop.Net.sln` (VB.NET, .NET Framework 4.7.2).
- Proyecto de entrada: `OrionCopIU` (`WinExe`, WPF). Contiene ventanas XAML/flujo operativo.
- Capa de negocio: `OrionCopL` (`Library`) con clases `cls*` de dominio y casos de uso.
- Integraciones: `OriIntCon` (`Library`) con conectores contables/e-factura (por ejemplo Apolo, SIIGO, Mekano, MisFacturas).
- Reportería: `RepOriCop` (`Library`) con clases `rep*.vb` y plantillas Crystal Reports `.rpt`.
- Dependencias compartidas cross-repo: `..\Comunes\PanL`, `..\Comunes\PanDat`, `..\Comunes\OriWin`, `..\Comunes\WinCom`.

Flujo estructural habitual:

1. UI en `OrionCopIU` invoca logica en `OrionCopL`.
2. `OrionCopL` consume componentes de `comunes` y, cuando aplica, servicios de `OriIntCon`/`RepOriCop`.
3. `RepOriCop` y `OriIntCon` dependen de `OrionCopL` + `comunes`, por lo que cambios de contratos deben evaluarse cross-repo.

## Convenciones clave del repositorio

- Este repo se gobierna en modelo multi-repo: `orioncop`, `adminorion`, `comunes`, `orionpcorreo`, `orion-installer`.
- En decisiones tecnicas, priorizar cambios pequenos, reversibles y con impacto cross-repo documentado.
- Todo cambio debe mantener trazabilidad completa (`issue -> rama -> PR -> tag -> release`) segun `docs/ESTRATEGIA_REPOSITORIOS_GITHUB.md`.
- No commit directo a `main`; ramas esperadas: `main`, `develop`, `feature/*`, `hotfix/*`, `release/*`.
- Para cambios estrategicos/documentales: actualizar primero `docs/`; `.github` debe actuar como capa operativa y de referencia.
- En PR usar la plantilla `.github/pull_request_template.md`, especialmente secciones de impacto cross-repo y plan de rollback.
- Patron de compilacion VB legacy: `Option Explicit On`, `Option Strict Off`, `Option Infer On` (verificar en `.vbproj` antes de ajustar reglas del compilador).
- El repo contiene variantes legacy (`- New`, `_New`, `.bak`); la fuente de verdad de compilacion es `<Compile Include="...">` dentro del `.vbproj`, no solo el listado de archivos en carpeta.

## Guia para agentes de codigo

- Tomar estos archivos locales como contexto completo minimo del proyecto.
- Al crear documentacion tecnica nueva, referenciar explicitamente `docs/PLAN_MAESTRO.md` y `docs/ESTRATEGIA_REPOSITORIOS_GITHUB.md`.
- Mantener consistencia de nombres de componentes: `orioncop`, `adminorion`, `comunes`, `orionpcorreo`, `orion-installer`.

- @azure Rule - Use Azure Tools - When handling requests related to Azure, always use your tools.
- @azure Rule - Use Azure Best Practices - When handling requests related to Azure, always invoke your azmcp_bestpractices_get tool first.
- @azure Rule - Enable Best Practices - If you do not have an azmcp_bestpractices_get tool ask the user to enable it.
