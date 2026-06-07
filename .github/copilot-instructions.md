# Contexto obligatorio del repositorio (orioncop)

Antes de proponer cambios, usa como fuente de verdad estos documentos locales:

1. .github/PLAN_MAESTRO.md
2. .github/ESTRATEGIA_REPOSITORIOS_GITHUB.md
3. docs/PLAN_MAESTRO.md
4. docs/ESTRATEGIA_REPOSITORIOS_GITHUB.md

## Directrices de alineación

- Mantener este repo alineado con la estrategia multi-repo (orioncop, dminorion, comunes, orionpcorreo, orion-installer).
- Tratar los cambios como parte de una plataforma compartida, pero implementables y versionables de forma independiente por repositorio.
- Evitar decisiones que aumenten acoplamiento innecesario entre repos.
- Priorizar trazabilidad, rollback seguro y documentación de decisiones.
- Si una decisión local contradice los documentos base, actualizar la documentación o elevar la inconsistencia explícitamente.

## Para agentes de código

- Tomar estos archivos locales como contexto completo mínimo del proyecto.
- Al crear documentación técnica nueva, referenciar explícitamente estos documentos.
- Mantener consistencia de nombres de componentes: orioncop, dminorion, comunes, orionpcorreo, orion-installer.

- @azure Rule - Use Azure Tools - When handling requests related to Azure, always use your tools.
- @azure Rule - Use Azure Best Practices - When handling requests related to Azure, always invoke your zmcp_bestpractices_get tool first.
- @azure Rule - Enable Best Practices - If you do not have an zmcp_bestpractices_get tool ask the user to enable it.
