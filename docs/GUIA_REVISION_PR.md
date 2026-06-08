# 👨‍💼 GUÍA: Cómo Revisar un PR paso a paso

## Introducción

Esta guía te ayuda a revisar PRs de forma **rápida, consistente y efectiva** como CEO del proyecto.

No necesitas ser un experto técnico en todos los detalles; la guía está diseñada para que valides:
- ✅ Que el cambio funciona
- ✅ Que respeta convenciones
- ✅ Que no rompe nada existente
- ✅ Que está documentado

---

## Fase 1: Lectura Inicial (5 minutos)

### 1.1 - Lee la descripción del PR

Busca responder:

```
❓ ¿Qué cambia?
   → Lee la sección "Descripción"

❓ ¿Por qué cambia?
   → Lee "Razón" o "Motivación"

❓ ¿Afecta otros repos?
   → Busca en "Impacto Cross-Repo"
   
❓ ¿Hay riesgo?
   → Busca "Risk Assessment"
```

### 1.2 - Valida el tipo de cambio

En el PR debe haber checkmarks para:
- [ ] Bug fix
- [ ] Feature nueva
- [ ] Breaking change
- [ ] Refactoring
- [ ] Documentación

**Tu validación**:
- ¿El tipo de cambio es claro?
- ¿Coincide con el título del PR?

Si algo no está claro → **Pide clarificación** (comentario en PR)

---

## Fase 2: Validación Técnica (10 minutos)

### 2.1 - Revisar compilación

El agente debe haber comentado un bloque como:

```
✅ Status Compilación
Build local: ✓ Exitoso
Errores: 0
Warnings críticos: 0
```

**Tu validación**:
- ¿Dice "Exitoso"?
- ¿Errores = 0?
- ¿Warnings críticos = 0?

Si NO → **NO apruebes**, pide que recompile.

### 2.2 - Revisar lista de archivos

Busca el bloque "Archivos Modificados":

```
📁 Archivos Modificados
- OrionCopL/ClsClienteL.vb
- OrionCopIU/winCliente.xaml
```

**Tu validación**:
- ¿Los archivos tienen sentido para el cambio?
- ¿Hay archivos "extraños" o innecesarios?
- ¿Falta algún archivo que debería estar?

Ejemplo de "raro":
```
❌ Cambio: "Agregar reporte de ventas"
   Pero modificó: "OrionCopL/ClsClienteL.vb" (¿por qué?)
```

Si algo parece raro → **Pregunta** (comentario en PR)

---

## Fase 3: Revisar Cambios (15 minutos)

### 3.1 - Cambios en `docs/`

GitHub muestra cambios. Busca estos archivos:

```
docs/PLAN_ORIONCOP_V1.md
docs/ESTRATEGIA_ORIONCOP_V1.md
docs/PATRONES_DETECTADOS.md
avv.agent.md
```

**Tu validación**:
- Si cambió arquitectura, ¿actualizaron `docs/`?
- Si es nuevo patrón, ¿está en `PATRONES_DETECTADOS.md`?
- ¿Las actualizaciones son claras?

Si falta documentación → **Pide que agreguen**

### 3.2 - Cambios en código VB.NET

GitHub te muestra línea por línea. Busca:

#### Convenciones

```vb
❓ Variables usan camelCase con prefijo?
   ✅ lstrNombre, mbolEsProgramable, GCSTRCLAVE
   ❌ lStr_Nombre, MboolEsProgramable

❓ Métodos usan PascalCase con prefijo?
   ✅ FStrObtenerDatos(), SProcesar()
   ❌ FStrObtenerDatos(), s_Procesar()

❓ Clases heredan correctamente?
   ✅ Inherits ClsCBObjetoPan
   ✅ Inherits ClsFormInterface
   ❌ Sin herencia de clase base
```

Si hay desvío → **Comenta en la línea**: "Por convención, usar `lstrNombre` en lugar de `lStr_Nombre`"

#### Lógica

```vb
❓ ¿La lógica es clara?
   ✅ Código legible, bien estructurado
   ❌ Código oscuro, difícil de entender

❓ ¿Hay duplicación?
   ✅ Código reutiliza funciones existentes
   ❌ La misma lógica aparece en 2+ lugares

❓ ¿Hay validaciones?
   ✅ Entrada validada antes de usar
   ❌ Falta validar datos de entrada
```

Si hay problema → **Comenta en la línea** con sugerencia clara

### 3.3 - Cambios en XAML/UI

Si cambió interfaz:

```xaml
❓ ¿Usa dicPanRecursos?
   ✅ {StaticResource MiEstilo}
   ❌ Estilo creado inline

❓ ¿Respeta diseño existente?
   ✅ Colores y tipografía consistentes
   ❌ Estilos nuevos y distintos

❓ ¿Los bindings son correctos?
   ✅ {Binding Propiedad, Mode=TwoWay}
   ❌ Sin binding, datos hardcodeados
```

Si hay inconsistencia → **Comenta**: "Por favor, usa `dicPanRecursos` para los estilos"

---

## Fase 4: Validar Impacto Cross-Repo (5 minutos)

### 4.1 - Busca sección "Impacto Cross-Repo"

Debería decir algo como:

```
✅ Impacto Cross-Repo
- [ ] Afecta AdminOrion
- [ ] Afecta OrionPlusCorreo
- [x] Afecta comunes/ → Describe cambios: 
      "Agregué propiedad en ClsCBObjetoPan"
```

**Tu validación**:
- ¿Identificaron correctamente qué afecta?
- ¿Necesita sincronización en otros repos?

Si el impacto está subdocumentado → **Pide que expanda**: "¿Este cambio en `comunes/` afecta AdminOrion? Describe cómo sincronizar"

### 4.2 - Valida plan de rollback

Busca sección "Plan de Rollback":

```
Plan de Rollback:
1. git revert <commit-sha>
2. Recompilar
3. Deshacer cambios en base de datos (si aplica)
```

**Tu validación**:
- ¿Hay plan claro?
- ¿Es realista?

Si falta → **Pide que agreguen**: "¿Qué pasos si necesitamos revertir?"

---

## Fase 5: Checklist Final (Antes de Aprobar)

### 5.1 - Compilación

```
☐ Build local: ✅ Exitoso
☐ Errores: 0
☐ Warnings críticos: 0
```

### 5.2 - Convenciones

```
☐ Nombres de variables: camelCase + prefijo
☐ Nombres de métodos: PascalCase + prefijo
☐ Herencia correcta: ClsCBObjetoPan / ClsFormInterface
☐ XAML usa dicPanRecursos (no recursos inline)
☐ Comentarios solo cuando es necesario
```

### 5.3 - Funcionalidad

```
☐ Cambio hace lo que promete
☐ No rompe features existentes
☐ Casos edge cubiertos (si aplica)
```

### 5.4 - Documentación

```
☐ docs/ actualizado (si es arquitectura)
☐ Patrones nuevos documentados
☐ Commits tienen mensajes claros
☐ PR describe cambios claramente
```

### 5.5 - Impacto

```
☐ Cross-repo identificado correctamente
☐ Plan de rollback existe
☐ Risk assessment está en el PR
```

---

## Decisión Final

### Si TODO está ✅

Comenta en el PR:

```
Aprobado ✅

Compilación exitosa, convenciones seguidas, 
documentación completa. Listo para merge.
```

Luego:
1. Aprueba el PR (botón de GitHub)
2. Agente AVV hace merge automático

### Si ALGUNOS detalles falta

Comenta en el PR los puntos específicos:

```
Cambios solicitados:

1. Línea 42: Usar `lstrNombre` en lugar de `lStr_Nombre`
2. Sección XAML: Usar `dicPanRecursos` para el estilo
3. Documentación: Actualizar `docs/PLAN_ORIONCOP_V1.md` 
   describiendo el nuevo patrón

Por favor, hace estos cambios y vuelve a solicitar aprobación.
```

Luego:
1. Selecciona "Request changes" en GitHub
2. Agente AVV hace los ajustes
3. Vuelta al paso 1

### Si HAY PROBLEMAS CRÍTICOS

Comenta en el PR:

```
Problema crítico 🔴

Este cambio rompe [X] porque [Y].
Necesitamos discutir enfoque antes de continuar.

Cc @agente-avv para que veamos alternativa.
```

Luego:
1. Selecciona "Reject" o "Request changes"
2. Agenda conversación con el equipo

---

## Atajos Rápidos

### PR Pequeño (Refactoring, docs)
- ⏱️ Tiempo: 5 minutos
- Pasos: 1.1, 2.1, 5.2, 5.4

### PR Mediano (Feature nueva, UI)
- ⏱️ Tiempo: 15 minutos
- Pasos: 1, 2, 3.2, 3.3, 5

### PR Grande (Cambio arquitectónico)
- ⏱️ Tiempo: 30+ minutos
- Pasos: Todos, más discusión con el equipo

---

## Preguntas Frecuentes

### P: ¿Necesito entender todos los detalles del código?

**R**: No. Valida:
- ✅ Estructura y convenciones
- ✅ Compilación exitosa
- ✅ Documentación completa
- ✅ No hay cambios "raros"

Para detalles técnicos complejos, pregunta al agente AVV.

### P: ¿Qué si no sé qué es un "breaking change"?

**R**: Un breaking change es cuando el cambio requiere que otros proyectos lo adapten.

Ejemplo:
```
❌ Breaking: Eliminar parámetro de método usado en AdminOrion
✅ No-breaking: Agregar parámetro opcional
```

Pide que documenten si hay breaking changes.

### P: ¿Puedo rechazar un PR por "no me gusta el código"?

**R**: Rechaza por:
- ❌ Convenciones no seguidas
- ❌ Compilación fallida
- ❌ Lógica rota o confusa
- ❌ Documentación incompleta

No rechaces por:
- ✅ "Yo lo haría distinto"
- ✅ "No me gusta este patrón"

(Si es un patrón problemático, crea issue separado)

### P: ¿Cuánto tiempo debería tomar revisar un PR?

**R**: 
- Pequeño: 5 minutos
- Mediano: 15 minutos
- Grande: 30 minutos
- Muy complejo: +1 hora + reunión

---

## Contacto y Escalaciones

Si algo no está claro o hay conflicto:

1. **Pregunta al agente AVV** (en el PR):
   "¿Esto impacta AdminOrion? Necesito clarificar..."

2. **Escala a reunión** si es cambio arquitectónico grande

3. **Documenta decisión** en el PR para referencia futura

---

**Última actualización**: 2026-06-08  
**Versión**: 1.0  
**Creada para**: CEO @juandevian
