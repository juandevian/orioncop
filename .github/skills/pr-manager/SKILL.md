---
name:	pr-manager
description: Esta skill habilita al agente para **crear, revisar y preparar pull Requests** de forma automatizada, dejándolos listos para validación del CEO (@juandevian).
version:	1.0.0
last_updated:	2026-06-08
status:	Activo / Lista para Producción
target_agent:	Agente AVV
category:	Automatización de Flujo de Trabajo / DevOps / Git CI-CD
repository:	juandevian/orioncop
approver: CEO (@juandevian)
---
El agente:
1. Realiza cambios en rama feature
2. Valida compilación y convenciones
3. Crea el PR con contexto completo
4. Deja una checklist para el CEO
5. Espera aprobación antes de merge

---

## Prerequisitos

- ✅ Git configurado localmente (`git config user.name`, `git config user.email`)
- ✅ Acceso a repositorio `juandevian/orioncop`
- ✅ Rama `develop` como base para features
- ✅ Rama `main` como rama de estable (sin merge directo)

---

## Fase 1: Preparación

### 1.1 - Crear rama feature

```bash
git checkout develop
git pull origin develop
git checkout -b feature/nombre-descriptivo
```

**Convenciones de rama**:
- `feature/agregar-reporte-ventas`
- `feature/mejorar-validacion-cliente`
- `hotfix/corregir-error-login`
- `release/v1.2.3`

### 1.2 - Actualizar documentación (si aplica)

Si el cambio afecta arquitectura:
- [ ] Actualizar `docs/PLAN_ORIONCOP_V1.md`
- [ ] Actualizar `docs/ESTRATEGIA_ORIONCOP_V1.md`
- [ ] Actualizar `docs/PATRONES_DETECTADOS.md` (si hay patrones nuevos)

---

## Fase 2: Desarrollo y Validación

### 2.1 - Escribir código

Seguir convenciones de `avv.agent.md`:
- Variables: `camelCase` con prefijos (`lstr`, `mbol`, `G`, etc.)
- Métodos: `PascalCase` con prefijos (`FStr`, `SProcesar`, etc.)
- Clases: `PascalCase` con prefijo (`cls`, `clsCB`, etc.)

### 2.2 - Validar compilación local

```powershell
# Opción 1: Build rápido (proyecto específico)
msbuild .\OrionCopIU\OrionCopIU.vbproj /p:Configuration=Release /p:Platform=x86

# Opción 2: Build completo (como CI)
nuget restore .\OrionCop.Net.sln
msbuild .\OrionCop.Net.sln /p:Configuration=Release /m
```

**Criterios de éxito**:
- ✅ Cero errores de compilación
- ✅ Cero warnings críticos (CodeAnalysis)
- ✅ Cambios alineados con `.editorconfig` y `.ruleset`

### 2.3 - Revisar impacto cross-repo

Antes de comitear, preguntarse:

- ¿Cambié algo en `../comunes/` (PanL, PanDat, OriWin, WinCom)?
  - Si sí: ¿Necesita sincronización en `AdminOrion`?
- ¿Cambié interfaces o clases base?
  - Si sí: ¿Hay breaking changes?
  - Si sí: ¿Necesita versionamiento?

### 2.4 - Documentar cambios release
Agregar a releases.md en `docs/` 
Descripción detallada:
- Qué cambió y por qué
- Impacto en otros proyectos (si aplica)
- Cambios en base de datos (si aplica)
- Versionamiento vX.X.X o Release X.X.X
Para garantizar trazabilidad.
Si el archivo no existe créalo.

### 2.5 - Documentar cambios en commit

```bash
git add .
git commit -m "feat: descripción clara del cambio

Descripción detallada:
- Qué cambió y por qué
- Impacto en otros proyectos (si aplica)
- Cambios en base de datos (si aplica)

Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>"
```

**Formato de commit**:
- `feat:` - Feature nueva
- `fix:` - Corrección de bug
- `docs:` - Cambios en documentación
- `refactor:` - Reorganización sin cambio funcional
- `test:` - Cambios en tests (cuando aplique)

---

## Fase 3: Crear Pull Request

### 3.1 - Push a rama remota

```bash
git push origin feature/nombre-descriptivo
```

### 3.2 - Crear PR via GitHub CLI

```bash
gh pr create \
  --base develop \
  --head feature/nombre-descriptivo \
  --title "feat: Descripción clara y concisa" \
  --body-file TEMPLATE_PR.txt
```

O usar plantilla interactiva:

```bash
gh pr create --web  # Abre navegador
```

### 3.3 - Completar plantilla PR (`.github/pull_request_template.md`)

```markdown
## 📋 Descripción
Resumen claro de qué se cambió y por qué.

## 🎯 Tipo de Cambio
- [ ] Bug fix
- [ ] Feature nueva
- [ ] Breaking change
- [ ] Refactoring
- [ ] Documentación

## 📊 Impacto Cross-Repo
- [ ] Afecta `AdminOrion` → Describe sincronización
- [ ] Afecta `OrionPlusCorreo` → Describe sincronización
- [ ] Afecta `comunes/` → Listar cambios exactos
- [ ] No impacta otros repos

## 🔗 Vinculado a Issue
Closes #123

## 📝 Testing
Cómo validaste el cambio (manual, automático, etc.)

## 🔄 Plan de Rollback
Pasos para revertir si algo sale mal.

## ✅ Checklist Previo a Merge
- [ ] Compilación exitosa (build log adjunto)
- [ ] Convenciones seguidas (nombres, estructura)
- [ ] Documentación actualizada (si aplica)
- [ ] Impacto cross-repo documentado
- [ ] Commits tienen mensaje claro
```

---

## Fase 4: Validación y Revisión

### 4.1 - Agregar información de compilación

Comentar en el PR:

```markdown
## ✅ Status Compilación

**Build local**: ✓ Exitoso
**Comando**: `msbuild .\OrionCop.Net.sln /p:Configuration=Release /m`
**Errores**: 0
**Warnings críticos**: 0
**Duración**: 2m 15s

**Archivos compilados**:
- OrionCopIU.exe
- OrionCopL.dll
- OriIntCon.dll
- RepOriCop.dll
```

### 4.2 - Listar cambios afectados

Comentar en el PR:

```markdown
## 📁 Archivos Modificados

### OrionCopL (Lógica)
- ClsClienteL.vb
- ClsOrdenVentaL.vb

### OrionCopIU (Interfaz)
- winCliente.xaml
- winCliente.xaml.vb

### Comunes (Impacto potencial)
- OriWin/dicPanRecursos.xaml (nuevo estilo agregado)
```

### 4.3 - Crear checklist para el CEO

Comentar en el PR:

```markdown
## 🔍 Checklist de Validación (CEO)

Antes de mergear, valida:

### Funcionalidad
- [ ] El cambio hace lo que promete
- [ ] Probaste el flujo end-to-end
- [ ] Casos edge están cubiertos

### Código
- [ ] Convenciones seguidas
- [ ] No hay código muerto
- [ ] Estructura es clara y mantenible

### Compatibilidad
- [ ] No rompe features existentes
- [ ] Backward compatible (si aplica)
- [ ] Cross-repo sincronizado (si aplica)

### Documentación
- [ ] Cambios en `docs/` actualizados
- [ ] Patrones nuevos documentados
- [ ] Commits tienen mensajes claros

### Risk Assessment
- **Riesgo**: Bajo / Medio / Alto
- **Justificación**: [Explica la evaluación]
- **Plan de rollback**: [Pasos si falla]
```

---

## Fase 5: Post-Validación (CEO)

### 5.1 - CEO revisa PR (Ver `docs/GUIA_REVISION_PR.md`)

El CEO:
1. Revisa cambios
2. Valida convenciones
3. Aprueba o solicita cambios

### 5.2 - Si CEO solicita cambios

El agente AVV:
1. Recibe feedback
2. Hace cambios en la misma rama
3. Commitea: `git commit -m "refactor: Cambios solicitados en PR #XXX"`
4. Push: `git push origin feature/nombre-descriptivo`
5. El PR se actualiza automáticamente

### 5.3 - Si CEO aprueba

El agente AVV:
1. Espera aprobación explícita
2. Hace merge a `develop`: `gh pr merge --squash` o `--rebase`
3. Elimina rama: `git branch -d feature/nombre-descriptivo`
4. Crea tag de versión (si aplica)

---

## Fase 6: Post-Merge

### 6.1 - Crear tag de versión (si es release)

```bash
git tag -a v1.2.3 -m "Release v1.2.3: Descripción"
git push origin v1.2.3
```

### 6.2 - Sincronizar cambios en repos hermanos (si aplica)

Si afectó `comunes/`:

```bash
# En AdminOrion/
git pull ../comunes main
git commit -m "chore: Sincronizar comunes con cambios de OrionCop"

# En OrionPlusCorreo/
git pull ../comunes main
git commit -m "chore: Sincronizar comunes con cambios de OrionCop"
```

### 6.3 - Actualizar documentación

- [ ] Actualizar `docs/PATRONES_DETECTADOS.md` si descubriste algo nuevo
- [ ] Actualizar `avv.agent.md` si cambió la arquitectura
- [ ] Crear entrada en changelog (si es release)

---

## Checklist de Calidad (Pre-PR)

Antes de crear el PR, el agente valida:

```
COMPILACIÓN
- [ ] Build local exitoso
- [ ] Cero errores de compilación
- [ ] Cero warnings críticos

CONVENCIONES
- [ ] Nombres de variables siguen camelCase
- [ ] Nombres de métodos siguen PascalCase
- [ ] Nombres de clases siguen PascalCase
- [ ] Prefijos aplicados correctamente
- [ ] Herencia correcta (ClsCBObjetoPan, ClsFormInterface)

IMPACTO
- [ ] Cambios en comunes/ documentados
- [ ] Impacto cross-repo identificado
- [ ] Documentación actualizada

CODIGO
- [ ] Lógica es clara y mantenible
- [ ] No hay código duplicado
- [ ] No hay código muerto
- [ ] Comentarios solo donde es necesario

COMMITS
- [ ] Mensajes descriptivos
- [ ] Commits lógicos (no amontonados)
- [ ] Co-author trailer incluido
```

---

## Errores Comunes a Evitar

| Error | Consecuencia | Solución |
|-------|-------------|----------|
| Mergear directamente a `main` | Requiere hotfix urgente | Siempre ir a `develop` primero |
| No documentar impacto cross-repo | Breaking changes inesperados | Revisar `comunes/` antes de comitear |
| Crear recursos XAML nuevos | Inconsistencia visual | Siempre usar `dicPanRecursos` |
| Cambiar convenciones de nombres | Confusión en código legacy | Seguir lo existente, no "arreglar" |
| No validar compilación | Merge fallido en CI | Compilar localmente antes de push |
| Commits amontonados | Historial ilegible | Hacer commits lógicos y pequeños |

---

## Contacto y Escalaciones

- 👨‍💼 **CEO (@juandevian)**: Aprobación de PR, decisiones arquitectónicas
- 🤖 **Agente AVV**: Crear PR, validar convenciones, alertar riesgos
- 👥 **Otros agentes**: Coordinar vía comentarios en PR

---

**Última actualización**: 2026-06-08  
**Versión**: 1.0  
**Estado**: Skill lista para usar