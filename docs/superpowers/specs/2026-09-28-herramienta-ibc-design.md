# Herramienta IBC (Interés Bancario Corriente) en Orión Plus

## Context
Hoy la tasa de mora se digita a mano en `winTasasMora` y se guarda en `OriTasasMora`; no hay fuente oficial ni tope legal. Se quiere una herramienta que consulte los certificados IBC de la SFC desde una API propia ya funcionando, los guarde localmente y permita parametrizar el interés de mora por copropiedad (Fijo o Variable), con un **tope duro de 1.5 × IBC** (Ley 675 de 2001, art. 30; validar la cita con abogado). Sin textos de exoneración: el sistema simplemente no permite superar el tope.

Clasificación brainstorming: **arquitectónico**. Este plan es el diseño; tras aprobarlo se escribe el spec en `docs/superpowers/specs/` y luego se invoca `writing-plans`.

## Decisiones tomadas por el usuario
- **Fijo:** consulta y guarda el IBC; la tasa cobrada nunca supera 1.5 × IBC (ver "Tope duro").
- **Variable:** tasa = IBC × factor; alimenta `OriTasasMora`. El **factor lo parametriza el usuario** (p. ej. 1.3) con un **máximo de 1.5**: cualquier valor entre 0 y 1.5 es válido, y uno mayor no se permite.
- Modos 0% y "No cobrar" ya existen por servicio (`FblnCausaMora`, etc.): **fuera de alcance**.
- Almacenamiento de certificados: **tabla nueva en MySQL**.
- API key: **una sola key embebida** en la app (no versionada en Git).
- **Sincronización automática solo al iniciar el Cierre de mes**, nunca al abrir la app (evita lentitud y consumo). La sincronización **manual** se puede hacer en cualquier momento.
- **Sin textos de exoneración ni consentimiento.**
- Convenciones de código: seguir las existentes; si hay varias, usar la más extendida (prefijos `Cls`/`F`/`S`/`Fdbl`/`Fbln`, `Friend`, regiones `#Region`, `ObjXxxDbl.ObjValorPro`, `GobjParametros`, `ClsPanorama.FdtbDataTable`, etc.). Ignorar archivos duplicados `* - New.vb`, `*_New.vb`, `*.bak`.

## Contrato de la API (leído de `C:\Users\juanv\dev\work\apis\api-ibc-certificados-sif`)
- Base prod: `https://api-ibc-certificados-sif.onrender.com`; header `X-Api-Key`. Local: `http://localhost:5080` (si no responde, probar `5100` antes de darlo por error).
- `GET /api/certificados/ultimo`, `/anio/{anio}`, `?n=N`, `/todos`. 404 sin cuerpo si no hay datos; 401 sin key; 503/500 con `{"error","status"}`.
- JSON camelCase: `idfile` (texto, NO Integer/Long), `fechaCertificado`, `fechaDesde`, `fechaHasta` (ISO sin zona), `ibc` (decimal, **% E.A. asumido**; el valor viene como porcentaje, p. ej. 19.38, y hay que dividir entre 100 para guardar como fracción como hace `OriTasasMora`).
- Render free: arranque en frío de 30-50 s y 503 por inactividad de la BD: timeout ~60 s y reintentos con backoff.
- "Periodo actual" no lo resuelve la API: el cliente elige el certificado con `fechaDesde <= fecha <= fechaHasta`.

## Cómo trata Orión la tasa hoy (verificado en código)
- `clsItemFactura.vb:1063` `FdecIntereseMora`: `DeudaCapital × (tasa / díasAño) × díasMora`, con `díasAño` = 365 (366 si año bisiesto). **No es 360.** La tasa es nominal anual simple, prorrateada por días.
- `clsCentroUtilidadOriCop.vb:997`: equivalente mensual ya definido como `TasaMora / 12`.
- Consecuencia: como se cobra tasa simple (sin capitalizar), cobrar `tasa = 1.5 × IBC` nunca supera el tope legal E.A. Para mostrar el mensual se usa la convención existente `anual / 12` (no la fórmula compuesta). Se cambia a 360 solo si el usuario lo decide (fuera de alcance).

## Diseño

### Componentes (unidades aisladas)
1. **`ClsIbcApiCliente`** (OriIntCon, junto a `clsInterfazMisFacturas`): `HttpClient` + Newtonsoft; consulta certificados; timeout, reintentos, mapeo de 401/404/503. Sin dependencia de BD ni UI.
2. **`ClsIbcCertificado`** (OrionCopL): entidad sobre `ClsCBObjetoPan`, tabla `OriIbcCertificados` (`Idfile` VARCHAR(20) PK, `FechaCertificado`, `FechaDesde`, `FechaHasta`, `Ibc` DECIMAL(6,2), `FechaDescarga`). Tabla global.
3. **`ClsIbcCalculo`** (OrionCopL, lógica pura, sin UI/BD, testeable):
   - `Tope(ibc) = 1.5 × ibc`.
   - `TasaFijaEfectiva(deseada, ibc) = Min(deseada, Tope(ibc))`.
   - `TasaVariable(ibc, factor) = ibc × factor`, con `factor` parametrizado por el usuario y validado `0 < factor <= 1.5` (constante `FactorMaximo = 1.5`).
   - `MensualParaMostrar(tasaAnual) = tasaAnual / 12`.
4. **Parametrización por centro de utilidad** (campos nuevos en la tabla de parámetros, expuestos vía `GobjParametros`): `ModoInteres` (Fijo/Variable), `TasaFijaDeseada`, `FactorVariable` (definido por el usuario, máx. 1.5). La **tasa efectiva** no se guarda como parámetro: se calcula y se escribe en `OriTasasMora`.
5. **`winParametrizacionInteres`** (WPF, OrionCopIU): elige modo, muestra IBC vigente, tasa anual y mensual equivalente, tope y aviso. Si el usuario digita un valor mayor al tope (tasa fija o factor), se muestra "no es posible superar 1.5 veces el IBC" y se pone automáticamente el máximo.
6. **`winConsultaIbc`** (menú Herramientas): botón consultar (manual, cualquier momento), lista de certificados locales, estado de la última consulta.
7. **Sincronización en Cierre de mes**: dentro de `ClsOrionCop.FblnCausoMoraGeneral` (`clsOrionCop.vb:4650`), justo después de `FdtmFechaCausaMoraGeneral()` y **antes** de `FdblTasaMoraFecha(...)` y de abrir la transacción, se llama `FblnSincronicaIbc(fechaCausacion, mensaje)`:
   - Consulta la API (o usa el certificado local si ya cubre la fecha de causación).
   - Fijo: efectiva = `Min(TasaFijaDeseada, 1.5 × IBC)`. Variable: `IBC × factor`. Si difiere de la vigente en `OriTasasMora`, cierra el periodo y agrega uno nuevo (`FechaDesde/FechaHasta` según vigencia).
   - Si no hay certificado que cubra la fecha (API caída y sin dato local vigente), **aborta la causación de intereses con mensaje claro**: no se leen tasas hasta que `OriTasasMora` esté actualizada. El usuario puede reintentar o sincronizar manualmente.
   - La misma sincronización manual reutiliza este método.

### Tope duro (Fijo): cómo se guarda "24 % que baja y vuelve a subir"
- `TasaFijaDeseada` (parámetro) es lo que el usuario definió (p. ej. 24 %) y **solo cambia si el usuario la edita**; el tope nunca la sobrescribe.
- La tasa efectiva se recalcula en cada sincronización: `Min(deseada, 1.5 × IBC vigente)`. Si el IBC baja y el tope queda en 21 %, `OriTasasMora` recibe 21 % y `TasaFijaDeseada` sigue en 24 %. Cuando el IBC sube y el tope vuelve a ≥ 24 %, se escribe de nuevo 24 %.
- Si el usuario digita a mano un valor por encima del tope actual, se le informa que no es posible y se guarda como deseada el máximo permitido en ese momento (comportamiento pedido: "automáticamente se pone la máxima").
- **Consecuencia a confirmar:** para poder reducir/restablecer, el modo Fijo también escribe en `OriTasasMora` cuando la efectiva cambia (no solo consulta).

### Puntos de integración (archivos a tocar)
- `..\Comunes\PanDat\XmlBd\OrionCop_Net.xml`: tabla `OriIbcCertificados`, campos de parámetros, `Version` 267 → 268.
- `OrionCopIU\MWOrionCop.xaml.vb`: declarar `MnuConsultaIbc` (~85-93), agregarlo a `HmnuHerramientas` (~300-335), `Case` en el despacho (~1715-1740) con `New WinConsultaIbc With {.WinPadre = Me}`.
- `OrionCopIU\winParametrizacion.xaml.vb` (~291): entrada para abrir la parametrización de interés, según el patrón de "Abrir Tasas de Mora".
- `OrionCopL\clsOrionCop.vb` (~4650): hook de sincronización descrito arriba.
- Reutilizar: `ClsCentroUtilidadOriCop.FdtbTasasMora`/`FdblTasaMoraFecha`, patrón `HttpClient` de `clsInterfazMisFacturas`, `MsgBox`/`SLevanteEveNoti`.
- El cálculo de mora en `clsItemFactura` **no se modifica**.

### API key
No se versiona en Git: se inyecta al compilar desde un archivo local ignorado (o el instalador) y se ofusca en el ejecutable. Riesgo aceptado: API de solo lectura con datos públicos; rotación vía `api_clients` con nueva versión de Orión.

### Riesgos / puntos abiertos
- Confirmar con el dueño de la API que `ibc` es % E.A.
- El repo no tiene proyectos de prueba. Un "proyecto de pruebas" es un proyecto VB.NET adicional dentro de la solución (MSTest o NUnit) que solo contiene código de prueba: llama a `ClsIbcCalculo` y `ClsIbcCliente` con datos de ejemplo y verifica resultados sin abrir la app ni usar la BD. Se ejecuta con Visual Studio o `vstest`. Propuesta: MSTest, por venir con Visual Studio.
- La llamada HTTP del cierre de mes puede tardar hasta ~60 s en arranque en frío; mostrar avance/mensaje al usuario.

## Verificación
1. Pruebas unitarias (aprobadas) de `ClsIbcCalculo`: IBC 19.38 → tope 29.07; fijo 24 con tope 29.07 → 24; IBC baja y tope 21 → efectiva 21 con deseada 24; IBC sube de nuevo → 24; factor 1.3 → IBC × 1.3 (válido); factor 1.5 válido; factor 1.6 rechazado y ajustado a 1.5; mensual = anual/12.
2. Cliente API contra `http://localhost:5080` (o `5100` si 5080 falla) y contra producción con la key: 200, 401, 404, 503 simulado y arranque en frío.
3. Compilar con MSBuild `OrionCop.Net.sln` y probar en desarrollo: menú Herramientas, actualización del esquema a versión 268 en una BD de prueba.
4. Manual: Fijo con valor > tope → mensaje y máximo; Variable → fila nueva en `OriTasasMora`; cierre de mes con API caída y sin dato local → causación abortada con mensaje; con dato local vigente → continúa.
5. Flujo de repo según `AGENTS.md`: rama desde `main`, PR, sin push directo a `main`.
