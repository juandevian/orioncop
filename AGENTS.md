# AGENTS

## Repositorio
`orioncop` es el módulo principal de operación de la plataforma Panorama Legacy en VB.NET. Es el núcleo funcional del sistema para la gestión contable y operativa, con foco en la lógica del negocio, la interfaz de usuario y la integración con componentes compartidos y reportes. El repositorio mantiene el software legacy que soporta la operación principal del negocio, sin absorber funciones de administración, correo ni instalación que corresponden a otros repositorios del ecosistema.

En términos prácticos, este proyecto concentra la operación principal del negocio y la experiencia de usuario del módulo; `docs/` guarda la documentación técnica y operativa, `src/` está reservado para estabilización gradual del código y `build/` aloja artefactos y salidas de compilación no versionadas. Leer `README.md` antes de operar y mantener alineación con la estrategia multi-repo del proyecto.

## Regla obligatoria
`main` está protegida. No se acepta push directo a `main`.
Todo cambio debe entrar por una rama de trabajo y una PR hacia `main`.

## Flujo base
```bash
git status
git add .
git commit -m "<tipo>: <descripción>"
git checkout main
git pull origin main
git checkout -b <nombre-rama>
git push -u origin <nombre-rama>
gh pr create --base main --head <nombre-rama> --title "<titulo-pr>" --body "<descripcion-pr>"
```

## Revisión y merge
- Revisar la PR en GitHub.
- Si es aprobada, fusionar con la estrategia permitida por el repo:
```bash
gh pr merge <numero-pr> --squash
```

## Actualizar main
```bash
git checkout main
git fetch origin
git pull origin main
```

## Release
```bash
git tag -a <version-tag> -m "Release <version-tag>"
git push origin <version-tag>
gh release create <version-tag> --title "<version-tag>" --notes "Release de la versión <version-tag>"
```

## Importante: PR vs. release
- La PR se crea cuando hay cambios de código o de documentación que deben integrarse en `main`.
- El release se crea después de que la versión ya fue aprobada, fusionada y etiquetada.
- Editar la nota de un release existente no crea una PR, porque no modifica el código fuente ni la rama principal.
- El flujo correcto es: rama de trabajo -> PR -> merge en `main` -> tag -> release.

## Reglas
- `main` es la única rama de integración.
- No hacer push directo a `main`.
- No crear tag antes de fusionar la PR en `main`.
- No hacer release antes del tag.
- Si la PR es rechazada, corregir en la rama de trabajo, hacer nuevo commit y repetir el flujo.
- Si el build falla, resolver en la rama de trabajo antes de mergear.

## Archivo de referencia
Usar `release-process.md` como guía detallada del flujo.
