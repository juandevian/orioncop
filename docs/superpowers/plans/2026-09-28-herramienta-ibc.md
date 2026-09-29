# Herramienta IBC (Interés Bancario Corriente) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Consultar los certificados IBC de la SFC desde la API propia, guardarlos en MySQL, parametrizar el interés de mora por copropiedad (Fijo/Variable) con tope duro de 1.5 × IBC y actualizar `OriTasasMora` al iniciar el Cierre de mes.

**Architecture:** La lógica de negocio (tope, tasa efectiva, selección de certificado vigente, orquestación de la sincronización) va en clases puras de `OrionCopL` que dependen de tres interfaces pequeñas (`IIbcProveedor`, `IIbcAlmacen`, `IIbcTasasMora`), así se prueba sin BD ni red. Las implementaciones reales son: `ClsIbcApiCliente` (HTTP, en `OriIntCon`), `ClsIbcAlmacenBd` y `ClsIbcTasasMoraBd` (framework `ClsCBObjetoPan` de Panorama). La UI (dos ventanas WPF + menú) y el hook en el Cierre de mes se conectan al final.

**Tech Stack:** VB.NET, .NET Framework 4.7.2, WPF, MySQL (MySql.Data 8.0.17), Newtonsoft.Json 13.0.4, `HttpClient`, MSTest 2 (proyecto nuevo, fuera de la solución), MSBuild de Visual Studio 2022.

**Spec:** `docs/superpowers/specs/2026-09-28-herramienta-ibc-design.md` (léelo antes de empezar).

## Global Constraints

- Rama de trabajo: `feature/herramienta-ibc` (ya creada desde `main`). **Prohibido push directo a `main`**; el cierre es PR (ver Task 10).
- Convenciones de código existentes: clases `Friend Class ClsXxx`; funciones `F…`, subs `S…`; parámetros `a…` (`adblIbc`), locales `l…` (`ldblTasa`), campos privados `M…`; prefijos de tipo `str/dbl/dec/dtm/bln/ent/obj/lst/col/dtb/drw`; enumeraciones `EnuXxx`; regiones `#Region`; sin `Option`/atributos nuevos innecesarios. Si hay dos estilos, usar el más extendido en el archivo vecino.
- Ignorar y **no tocar** archivos duplicados `* - New.vb`, `*_New.vb`, `*.bak`.
- Proyectos viejos (`.vbproj` con `<Compile Include>` explícito): **todo archivo nuevo debe agregarse al `.vbproj`** (`.xaml` también como `<Page>`). Guardar archivos `.vb` con la misma codificación/finales de línea de los vecinos.
- Tasas: **en código y BD son fracciones** (0.1938 = 19.38 %). La API entrega `ibc` en porcentaje (19.38): dividir entre 100 y redondear a 6 decimales.
- `TasaMora` en `OriTasasMora` es **anual simple** (fracción). Cálculo vigente: `Deuda × (tasa / 365|366) × días` (`clsItemFactura.vb:1063`); mensual mostrado = `anual / 12`. **`clsItemFactura` no se modifica.**
- `ClsTasaMoraDbl` (`clsTasaMora.vb:309`) convierte lo asignado como si fuera tasa **mensual vencida** a anual simple: para guardar anual `A` se asigna `A/12` (fracción) a `ObjTasaMoraDbl.ObjValorPro`. La conversión usa `Val("…mv")` y `ToString` de la cultura: con coma decimal devuelve 0 en silencio, por eso el Task 6 relee y verifica.
- `ClsFechaDesdeTasaMoraDtm`: `FechaDesde` ≥ última `FechaDesde` + 1 día y ≤ `Today`. Al crear una fila el objeto cierra sola la `FechaHasta` de la anterior.
- Tope legal: **`FactorMaximo = 1.5`** (Ley 675 de 2001, art. 30; validar cita con abogado). El factor Variable lo define el usuario en `(0, 1.5]`.
- Modo `None` (centro no parametrizado) = comportamiento actual: ni se consulta la API ni se bloquea el cierre.
- API: header `X-Api-Key`; prod `https://api-ibc-certificados-sif.onrender.com`; local `http://localhost:5080` (probar `5100` si falla). Arranque en frío 30-50 s: timeout 60 s, 3 reintentos con backoff ante 503/timeout.
- `idfile` de la API es **texto** (nunca Integer/Long). Fechas ISO sin zona; sin conversión de zona.
- La API key **no se versiona**: vive en `OriIntCon\IbcApiKey.vb` (ignorado por Git), ofuscada; el repo trae `IbcApiKey.sample.vb`.
- Desviación deliberada del spec: `OriIbcCertificados` lleva `IdCarpeta`/`IdCentroUtil` (como toda tabla del framework `ClsCBObjetoPan`) en vez de ser global; cada centro guarda su copia (≈12 filas/año). Avisar al usuario en el PR.
- Compilar con MSBuild de VS 2022: `C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe`, config `Release|x64` (como `scripts/build/build.ps1`).

## Review Focus

1. Centro **sin parametrizar** (modo `None`): el Cierre de mes debe seguir igual que hoy, sin llamar a la API ni abortar (test en Task 7).
2. **API caída con dato local vigente**: el cierre continúa con el dato local; **sin dato local vigente**: aborta con mensaje y no causa intereses (Task 7).
3. **Cultura con coma decimal**: asignar la tasa a `ClsTasaMoraDbl` no debe dejar 0 en silencio; se relee y se lanza error si difiere (Task 6).
4. **Certificado con `fechaDesde` futura** (publicado a fin de mes) no debe elegirse como vigente hoy (Task 2).
5. **`FechaDesde` que queda > hoy** (última fila ya es de hoy): no escribir, avisar (Task 7).
6. **`idfile` numérico grande** (`9920261260`) no debe romperse al parsear ni guardar (Tasks 3 y 4).
7. **Tasa fija deseada 24 % con tope que baja a 21 % y sube de nuevo**: la deseada no se sobrescribe y la efectiva vuelve a 24 % (Task 1 y 7).
8. **IBC = 0 o negativo / respuesta 404 / lista vacía**: no debe producir tasa 0 ni crash (Tasks 2, 3 y 7).

---

## File Structure

| Archivo | Acción | Responsabilidad |
|---|---|---|
| `OrionCopL\clsIbcCalculo.vb` | Crear | Reglas puras: tope, tasa fija efectiva, variable, mensual, selección de certificado vigente, DTO `StcIbcCertificado`, enum `EnuModoInteres` |
| `OrionCopL\clsIbcInterfaces.vb` | Crear | `IIbcProveedor`, `IIbcAlmacen`, `IIbcTasasMora`, `ErrorIbcApiException` |
| `OrionCopL\clsIbcSincroniza.vb` | Crear | Orquestador de la sincronización (usa las interfaces) |
| `OrionCopL\clsIbcCertificado.vb` | Crear | Entidad `ClsIbcCertificado` (tabla `OriIbcCertificados`) + `ClsIbcAlmacenBd` |
| `OrionCopL\clsIbcTasasMoraBd.vb` | Crear | `ClsIbcTasasMoraBd` sobre `ClsTasaMora`/`FdblTasaMoraFecha` |
| `OrionCopL\clsCentroUtilidadOriCop.vb` | Modificar | 3 propiedades nuevas de parámetros |
| `OrionCopL\clsOrionCop.vb` | Modificar (~4650) | Hook en `FblnCausoMoraGeneral` + registro del proveedor |
| `OrionCopL\acOrionCopL.vb`, `OriIntCon\acOriIntCon.vb` | Modificar | `InternalsVisibleTo("OrionCopL.Tests")` |
| `OriIntCon\clsIbcApiCliente.vb`, `IbcApiKey.sample.vb` | Crear | Cliente HTTP y contenedor de la key |
| `..\Comunes\PanDat\XmlBd\OrionCop_Net.xml` | Modificar | Tabla, columnas, `Version` 267 → 268 |
| `OrionCopIU\winParametrizacionInteres.xaml(.vb)`, `winConsultaIbc.xaml(.vb)` | Crear | Ventanas |
| `OrionCopIU\MWOrionCop.xaml.vb`, `winParametrizacion.xaml.vb` | Modificar | Menú y acceso |
| `tests\OrionCopL.Tests\*` | Crear | Proyecto MSTest (fuera de la solución) |
| `.gitignore` | Modificar | Ignorar `OriIntCon/IbcApiKey.vb` |

---

### Task 1: Proyecto de pruebas y reglas de cálculo (`ClsIbcCalculo`)

**Files:**
- Create: `tests\OrionCopL.Tests\OrionCopL.Tests.vbproj`
- Create: `tests\OrionCopL.Tests\IbcCalculoTests.vb`
- Create: `OrionCopL\clsIbcCalculo.vb`
- Modify: `OrionCopL\OrionCopL.vbproj` (agregar `<Compile Include="clsIbcCalculo.vb" />` junto a `clsTasaMora.vb`, ~línea 211)
- Modify: `OrionCopL\acOrionCopL.vb` (agregar `<Assembly: InternalsVisibleTo("OrionCopL.Tests")>` tras la línea 8)

**Interfaces:**
- Produces:
  - `Friend Enum EnuModoInteres As Byte : None = 0 : EnuFijo : EnuVariable`
  - `ClsIbcCalculo.CDBLFACTORMAXIMO As Double` (= 1.5)
  - `Shared FdblIbcComoFraccion(adblPorcentaje As Double) As Double`
  - `Shared FdblTopeMaximo(adblIbc As Double) As Double`
  - `Shared FdblTasaFijaEfectiva(adblDeseada As Double, adblIbc As Double) As Double`
  - `Shared FblnFactorValido(adblFactor As Double) As Boolean`
  - `Shared FdblTasaVariable(adblIbc As Double, adblFactor As Double) As Double`
  - `Shared FdblMensualParaMostrar(adblTasaAnual As Double) As Double`

- [ ] **Step 1: Crear el proyecto de pruebas (SDK-style, fuera de la solución)**

`tests\OrionCopL.Tests\OrionCopL.Tests.vbproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net472</TargetFramework>
    <AssemblyName>OrionCopL.Tests</AssemblyName>
    <RootNamespace>OrionCopL.Tests</RootNamespace>
    <Platforms>x64</Platforms>
    <PlatformTarget>x64</PlatformTarget>
    <OptionStrict>Off</OptionStrict>
    <IsPackable>false</IsPackable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="MSTest.TestFramework" Version="3.6.4" />
    <PackageReference Include="MSTest.TestAdapter" Version="3.6.4" />
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.11.1" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\..\OrionCopL\OrionCopL.vbproj" />
    <ProjectReference Include="..\..\OriIntCon\OriIntCon.vbproj" />
  </ItemGroup>
</Project>
```

- [ ] **Step 2: Escribir las pruebas que fallan**

`tests\OrionCopL.Tests\IbcCalculoTests.vb`:

```vb
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class IbcCalculoTests
    Private Const CDBLTOL As Double = 0.0000001

    <TestMethod>
    Public Sub Ibc_porcentaje_se_convierte_a_fraccion()
        Assert.AreEqual(0.1938, ClsIbcCalculo.FdblIbcComoFraccion(19.38), CDBLTOL)
    End Sub

    <TestMethod>
    Public Sub Tope_es_uno_coma_cinco_veces_el_ibc()
        Assert.AreEqual(0.2907, ClsIbcCalculo.FdblTopeMaximo(0.1938), CDBLTOL)
    End Sub

    <TestMethod>
    Public Sub Fija_bajo_el_tope_se_mantiene()
        Assert.AreEqual(0.24, ClsIbcCalculo.FdblTasaFijaEfectiva(0.24, 0.1938), CDBLTOL)
    End Sub

    <TestMethod>
    Public Sub Fija_sobre_el_tope_baja_al_tope_y_vuelve_a_la_deseada_cuando_el_tope_sube()
        ' deseada 24 %: IBC 14 % => tope 21 %; luego IBC 19.38 % => tope 29.07 %
        Assert.AreEqual(0.21, ClsIbcCalculo.FdblTasaFijaEfectiva(0.24, 0.14), CDBLTOL)
        Assert.AreEqual(0.24, ClsIbcCalculo.FdblTasaFijaEfectiva(0.24, 0.1938), CDBLTOL)
    End Sub

    <TestMethod>
    Public Sub Factor_valido_esta_en_cero_exclusivo_a_uno_coma_cinco()
        Assert.IsTrue(ClsIbcCalculo.FblnFactorValido(1.3))
        Assert.IsTrue(ClsIbcCalculo.FblnFactorValido(1.5))
        Assert.IsFalse(ClsIbcCalculo.FblnFactorValido(1.6))
        Assert.IsFalse(ClsIbcCalculo.FblnFactorValido(0))
        Assert.IsFalse(ClsIbcCalculo.FblnFactorValido(-1))
    End Sub

    <TestMethod>
    Public Sub Variable_es_ibc_por_factor()
        Assert.AreEqual(0.25194, ClsIbcCalculo.FdblTasaVariable(0.1938, 1.3), CDBLTOL)
    End Sub

    <TestMethod>
    Public Sub Mensual_para_mostrar_es_anual_entre_doce()
        Assert.AreEqual(0.024225, ClsIbcCalculo.FdblMensualParaMostrar(0.2907), CDBLTOL)
    End Sub
End Class
```

- [ ] **Step 3: Ejecutar y verificar que falla (no compila)**

Run (PowerShell): `& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" tests\OrionCopL.Tests\OrionCopL.Tests.vbproj -restore /p:Configuration=Release /p:Platform=x64 /v:minimal`
Expected: error `BC30451: 'ClsIbcCalculo' no está declarado`. (Si falla antes por restauración NuGet/proyectos, arreglar eso primero; no avanzar sin llegar a este error.)

- [ ] **Step 4: Implementar `ClsIbcCalculo` y el enum**

`OrionCopL\clsIbcCalculo.vb`:

```vb
#Region "Enumeradores"
Friend Enum EnuModoInteres As Byte
    None = 0
    EnuFijo
    EnuVariable
End Enum
#End Region

Friend Class ClsIbcCalculo
#Region "Definiciones"
    ' Ley 675 de 2001, art. 30: el interés de mora no puede exceder 1.5 veces el IBC certificado
    Friend Const CDBLFACTORMAXIMO As Double = 1.5
#End Region
#Region "Reglas de cálculo"
    ''' <summary>
    ''' Convierte el IBC de la API (porcentaje, ej. 19.38) a fracción (0.1938).
    ''' </summary>
    Friend Shared Function FdblIbcComoFraccion(adblPorcentaje As Double) As Double
        Return Math.Round(adblPorcentaje / 100, 6)
    End Function

    Friend Shared Function FdblTopeMaximo(adblIbc As Double) As Double
        Return Math.Round(adblIbc * CDBLFACTORMAXIMO, 6)
    End Function

    ''' <summary>
    ''' Tasa fija que realmente se cobra: la deseada, limitada por el tope vigente.
    ''' </summary>
    Friend Shared Function FdblTasaFijaEfectiva(adblDeseada As Double, adblIbc As Double) As Double
        Return Math.Min(adblDeseada, FdblTopeMaximo(adblIbc))
    End Function

    Friend Shared Function FblnFactorValido(adblFactor As Double) As Boolean
        Return adblFactor > 0 AndAlso adblFactor <= CDBLFACTORMAXIMO
    End Function

    Friend Shared Function FdblTasaVariable(adblIbc As Double, adblFactor As Double) As Double
        Return Math.Round(adblIbc * adblFactor, 6)
    End Function

    ''' <summary>
    ''' Equivalente mensual que se muestra al usuario (misma convención de FdtbTasasMora: anual / 12).
    ''' </summary>
    Friend Shared Function FdblMensualParaMostrar(adblTasaAnual As Double) As Double
        Return Math.Round(adblTasaAnual / 12, 6)
    End Function
#End Region
End Class
```

Agregar `<Compile Include="clsIbcCalculo.vb" />` en `OrionCopL.vbproj` y la línea `InternalsVisibleTo("OrionCopL.Tests")` en `acOrionCopL.vb`.

- [ ] **Step 5: Ejecutar y verificar que pasa**

Run: `& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" tests\OrionCopL.Tests\OrionCopL.Tests.vbproj -restore /p:Configuration=Release /p:Platform=x64 /v:minimal` y luego `& "C:\Program Files\Microsoft Visual Studio\2022\Common7\IDE\CommonExtensions\Microsoft\TestWindow\vstest.console.exe" tests\OrionCopL.Tests\bin\x64\Release\net472\OrionCopL.Tests.dll /Platform:x64`
Expected: `Passed: 7`. (Si la ruta de salida difiere, usar la que reporte MSBuild.)

- [ ] **Step 6: Commit**

```bash
git add tests/OrionCopL.Tests OrionCopL/clsIbcCalculo.vb OrionCopL/OrionCopL.vbproj OrionCopL/acOrionCopL.vb
git commit -m "feat: reglas de calculo IBC y proyecto de pruebas MSTest"
```

---

### Task 2: DTO de certificado y selección del certificado vigente

**Files:**
- Modify: `OrionCopL\clsIbcCalculo.vb`
- Test: `tests\OrionCopL.Tests\IbcCertificadoVigenteTests.vb`

**Interfaces:**
- Consumes: `ClsIbcCalculo` (Task 1).
- Produces:
  - `Friend Structure StcIbcCertificado` con propiedades `StrIdfile As String`, `DtmFechaCertificado As Date`, `DtmFechaDesde As Date`, `DtmFechaHasta As Date`, `DblIbc As Double` (**fracción**).
  - `ClsIbcCalculo.FblnCertificadoVigente(alstCertificados As IEnumerable(Of StcIbcCertificado), adtmFecha As Date, ByRef astcVigente As StcIbcCertificado) As Boolean` — True si encontró uno con `DtmFechaDesde <= fecha <= DtmFechaHasta` e `DblIbc > 0`; si hay varios, el de mayor `DtmFechaCertificado` y luego `StrIdfile` descendente.

- [ ] **Step 1: Escribir las pruebas que fallan**

`tests\OrionCopL.Tests\IbcCertificadoVigenteTests.vb`:

```vb
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class IbcCertificadoVigenteTests
    Private Shared Function FstcCert(astrId As String, adtmDesde As Date, adtmHasta As Date,
            adblIbc As Double, Optional adtmEmision As Date = Nothing) As StcIbcCertificado
        Dim lstc As New StcIbcCertificado
        lstc.StrIdfile = astrId
        lstc.DtmFechaDesde = adtmDesde
        lstc.DtmFechaHasta = adtmHasta
        lstc.DblIbc = adblIbc
        lstc.DtmFechaCertificado = If(adtmEmision = Nothing, adtmDesde.AddDays(-5), adtmEmision)
        Return lstc
    End Function

    <TestMethod>
    Public Sub Elige_el_que_cubre_la_fecha()
        Dim llst = New List(Of StcIbcCertificado) From {
            FstcCert("A", #2026-08-01#, #2026-08-31#, 0.19),
            FstcCert("B", #2026-09-01#, #2026-09-30#, 0.2)}
        Dim lstc As StcIbcCertificado
        Assert.IsTrue(ClsIbcCalculo.FblnCertificadoVigente(llst, #2026-09-15#, lstc))
        Assert.AreEqual("B", lstc.StrIdfile)
    End Sub

    <TestMethod>
    Public Sub No_elige_certificado_con_fecha_desde_futura()
        Dim llst = New List(Of StcIbcCertificado) From {
            FstcCert("FUT", #2026-10-01#, #2026-10-31#, 0.21)}
        Dim lstc As StcIbcCertificado
        Assert.IsFalse(ClsIbcCalculo.FblnCertificadoVigente(llst, #2026-09-28#, lstc))
    End Sub

    <TestMethod>
    Public Sub Lista_vacia_no_tiene_vigente()
        Dim lstc As StcIbcCertificado
        Assert.IsFalse(ClsIbcCalculo.FblnCertificadoVigente(New List(Of StcIbcCertificado), #2026-09-28#, lstc))
    End Sub

    <TestMethod>
    Public Sub Ibc_cero_o_negativo_se_ignora()
        Dim llst = New List(Of StcIbcCertificado) From {
            FstcCert("CERO", #2026-09-01#, #2026-09-30#, 0),
            FstcCert("NEG", #2026-09-01#, #2026-09-30#, -0.1)}
        Dim lstc As StcIbcCertificado
        Assert.IsFalse(ClsIbcCalculo.FblnCertificadoVigente(llst, #2026-09-15#, lstc))
    End Sub

    <TestMethod>
    Public Sub Con_varios_gana_la_emision_mas_reciente()
        Dim llst = New List(Of StcIbcCertificado) From {
            FstcCert("VIEJO", #2026-09-01#, #2026-09-30#, 0.19, #2026-08-20#),
            FstcCert("NUEVO", #2026-09-01#, #2026-09-30#, 0.2, #2026-08-28#)}
        Dim lstc As StcIbcCertificado
        Assert.IsTrue(ClsIbcCalculo.FblnCertificadoVigente(llst, #2026-09-15#, lstc))
        Assert.AreEqual("NUEVO", lstc.StrIdfile)
    End Sub

    <TestMethod>
    Public Sub Idfile_numerico_grande_se_conserva_como_texto()
        Dim llst = New List(Of StcIbcCertificado) From {
            FstcCert("9920261260", #2026-09-01#, #2026-09-30#, 0.2)}
        Dim lstc As StcIbcCertificado
        Assert.IsTrue(ClsIbcCalculo.FblnCertificadoVigente(llst, #2026-09-15#, lstc))
        Assert.AreEqual("9920261260", lstc.StrIdfile)
    End Sub
End Class
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: el mismo comando de MSBuild del Task 1, Step 5.
Expected: error de compilación `'StcIbcCertificado' no está declarado`.

- [ ] **Step 3: Implementar**

Agregar en `clsIbcCalculo.vb`, antes de `Friend Class ClsIbcCalculo`:

```vb
#Region "Estructuras"
Friend Structure StcIbcCertificado
    Friend Property StrIdfile As String
    Friend Property DtmFechaCertificado As Date
    Friend Property DtmFechaDesde As Date
    Friend Property DtmFechaHasta As Date
    ''' <summary>IBC como fracción (0.1938).</summary>
    Friend Property DblIbc As Double
End Structure
#End Region
```

Y dentro de `#Region "Reglas de cálculo"`:

```vb
    ''' <summary>
    ''' Busca el certificado cuya vigencia cubre la fecha. Ignora IBC no positivos. Si hay varios,
    ''' toma el de emisión más reciente (desempate por Idfile descendente).
    ''' </summary>
    Friend Shared Function FblnCertificadoVigente(alstCertificados As IEnumerable(Of StcIbcCertificado),
            adtmFecha As Date, ByRef astcVigente As StcIbcCertificado) As Boolean
        Dim lblnHay = False
        astcVigente = Nothing
        If Not IsNothing(alstCertificados) Then
            For Each lstcCert As StcIbcCertificado In alstCertificados
                If lstcCert.DblIbc > 0 AndAlso lstcCert.DtmFechaDesde <= adtmFecha AndAlso
                        adtmFecha <= lstcCert.DtmFechaHasta Then
                    If Not lblnHay OrElse lstcCert.DtmFechaCertificado > astcVigente.DtmFechaCertificado OrElse
                            (lstcCert.DtmFechaCertificado = astcVigente.DtmFechaCertificado AndAlso
                            String.CompareOrdinal(lstcCert.StrIdfile, astcVigente.StrIdfile) > 0) Then
                        astcVigente = lstcCert
                        lblnHay = True
                    End If
                End If
            Next
        End If
        Return lblnHay
    End Function
```

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: build + vstest del Task 1, Step 5.
Expected: `Passed: 13`.

- [ ] **Step 5: Commit**

```bash
git add OrionCopL/clsIbcCalculo.vb tests/OrionCopL.Tests/IbcCertificadoVigenteTests.vb
git commit -m "feat: seleccion del certificado IBC vigente por fecha"
```

---

### Task 3: Interfaces y cliente HTTP de la API (`ClsIbcApiCliente`)

**Files:**
- Create: `OrionCopL\clsIbcInterfaces.vb`
- Create: `OriIntCon\clsIbcApiCliente.vb`, `OriIntCon\IbcApiKey.sample.vb`
- Modify: `OrionCopL\OrionCopL.vbproj`, `OriIntCon\OriIntCon.vbproj` (Compile Include; ver Step 5), `OriIntCon\acOriIntCon.vb` (`InternalsVisibleTo("OrionCopL.Tests")`), `.gitignore`
- Test: `tests\OrionCopL.Tests\IbcApiClienteTests.vb`

**Interfaces:**
- Consumes: `StcIbcCertificado` (Task 2), `ClsIbcCalculo.FdblIbcComoFraccion`.
- Produces:
  - `Friend Interface IIbcProveedor : Function FlstConsulteCertificados(aentCantidad As Integer) As List(Of StcIbcCertificado) : End Interface`
  - `Friend Interface IIbcAlmacen : Function FlstCertificados() As List(Of StcIbcCertificado) : Sub SGuarde(alstCertificados As IEnumerable(Of StcIbcCertificado)) : End Interface`
  - `Friend Interface IIbcTasasMora : Function FdblTasaVigente(adtmFecha As Date) As Double : Function FdtmFechaDesdeUltima() As Date : Sub SRegistreTasa(adtmFechaDesde As Date, adblTasaAnual As Double) : End Interface`
  - `Friend Class ErrorIbcApiException : Inherits Exception` con `Friend ReadOnly Property EntCodigoHttp As Integer`.
  - `Friend Class ClsIbcApiCliente : Implements IIbcProveedor` con `New(astrUrlBase As String, astrApiKey As String, Optional ahndManejador As HttpMessageHandler = Nothing, Optional aentReintentos As Integer = 3, Optional aentEsperaMs As Integer = 2000)`.
  - `Friend Module MdefIbcApiKey` con `Friend Function FstrApiKey() As String` (en el sample devuelve `String.Empty`).

- [ ] **Step 1: Escribir las pruebas que fallan**

`tests\OrionCopL.Tests\IbcApiClienteTests.vb`:

```vb
Imports System.Net
Imports System.Net.Http
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class IbcApiClienteTests
    Private Class ManejadorFalso
        Inherits HttpMessageHandler
        Friend Respuestas As New Queue(Of HttpResponseMessage)
        Friend Llamadas As Integer = 0
        Friend UltimaKey As String = String.Empty
        Friend UltimaUrl As String = String.Empty
        Protected Overrides Function SendAsync(request As HttpRequestMessage,
                cancellationToken As CancellationToken) As Task(Of HttpResponseMessage)
            Llamadas += 1
            UltimaUrl = request.RequestUri.ToString
            If request.Headers.Contains("X-Api-Key") Then
                UltimaKey = request.Headers.GetValues("X-Api-Key").First
            End If
            Return Task.FromResult(Respuestas.Dequeue())
        End Function
    End Class

    Private Shared Function FrespOk(astrJson As String) As HttpResponseMessage
        Return New HttpResponseMessage(HttpStatusCode.OK) With {
            .Content = New StringContent(astrJson, Text.Encoding.UTF8, "application/json")}
    End Function

    Private Const CSTRJSON As String =
        "[{""idfile"":""9920261260"",""fechaCertificado"":""2026-08-28T00:00:00""," &
        """fechaDesde"":""2026-09-01T00:00:00"",""fechaHasta"":""2026-09-30T00:00:00"",""ibc"":19.38}]"

    Private Shared Function FobjCli(aobjManej As ManejadorFalso) As ClsIbcApiCliente
        Return New ClsIbcApiCliente("http://x", "KEY123", aobjManej, 3, 1)
    End Function

    <TestMethod>
    Public Sub Parsea_certificado_con_idfile_texto_e_ibc_como_fraccion()
        Dim lobjM As New ManejadorFalso
        lobjM.Respuestas.Enqueue(FrespOk(CSTRJSON))
        Dim llst = FobjCli(lobjM).FlstConsulteCertificados(1)
        Assert.AreEqual(1, llst.Count)
        Assert.AreEqual("9920261260", llst(0).StrIdfile)
        Assert.AreEqual(0.1938, llst(0).DblIbc, 0.0000001)
        Assert.AreEqual(New Date(2026, 9, 1), llst(0).DtmFechaDesde)
        Assert.AreEqual(New Date(2026, 9, 30), llst(0).DtmFechaHasta)
        Assert.AreEqual(New Date(2026, 8, 28), llst(0).DtmFechaCertificado)
    End Sub

    <TestMethod>
    Public Sub Envia_header_api_key_y_pide_n()
        Dim lobjM As New ManejadorFalso
        lobjM.Respuestas.Enqueue(FrespOk(CSTRJSON))
        FobjCli(lobjM).FlstConsulteCertificados(12)
        Assert.AreEqual("KEY123", lobjM.UltimaKey)
        Assert.IsTrue(lobjM.UltimaUrl.EndsWith("/api/certificados?n=12"))
    End Sub

    <TestMethod>
    Public Sub Reintenta_ante_503_y_luego_tiene_exito()
        Dim lobjM As New ManejadorFalso
        lobjM.Respuestas.Enqueue(New HttpResponseMessage(HttpStatusCode.ServiceUnavailable))
        lobjM.Respuestas.Enqueue(FrespOk(CSTRJSON))
        Dim llst = FobjCli(lobjM).FlstConsulteCertificados(1)
        Assert.AreEqual(2, lobjM.Llamadas)
        Assert.AreEqual(1, llst.Count)
    End Sub

    <TestMethod>
    Public Sub Lanza_error_con_401_sin_reintentar()
        Dim lobjM As New ManejadorFalso
        lobjM.Respuestas.Enqueue(New HttpResponseMessage(HttpStatusCode.Unauthorized))
        Try
            FobjCli(lobjM).FlstConsulteCertificados(1)
            Assert.Fail("Debió lanzar ErrorIbcApiException")
        Catch ex As ErrorIbcApiException
            Assert.AreEqual(401, ex.EntCodigoHttp)
            Assert.AreEqual(1, lobjM.Llamadas)
        End Try
    End Sub

    <TestMethod>
    Public Sub Lista_vacia_devuelve_lista_vacia()
        Dim lobjM As New ManejadorFalso
        lobjM.Respuestas.Enqueue(FrespOk("[]"))
        Assert.AreEqual(0, FobjCli(lobjM).FlstConsulteCertificados(1).Count)
    End Sub

    <TestMethod>
    Public Sub Sin_api_key_configurada_lanza_error_claro()
        Dim lobjCli As New ClsIbcApiCliente("http://x", String.Empty, New ManejadorFalso)
        Try
            lobjCli.FlstConsulteCertificados(1)
            Assert.Fail("Debió lanzar ErrorIbcApiException")
        Catch ex As ErrorIbcApiException
            StringAssert.Contains(ex.Message, "API key")
        End Try
    End Sub
End Class
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: build del Task 1, Step 5. Expected: `'ClsIbcApiCliente' no está declarado`.

- [ ] **Step 3: Implementar interfaces y excepción**

`OrionCopL\clsIbcInterfaces.vb`:

```vb
Friend Interface IIbcProveedor
    ''' <summary>Consulta los últimos N certificados en la API (IBC como fracción).</summary>
    Function FlstConsulteCertificados(aentCantidad As Integer) As List(Of StcIbcCertificado)
End Interface

Friend Interface IIbcAlmacen
    Function FlstCertificados() As List(Of StcIbcCertificado)
    Sub SGuarde(alstCertificados As IEnumerable(Of StcIbcCertificado))
End Interface

Friend Interface IIbcTasasMora
    ''' <summary>Tasa anual vigente a la fecha en OriTasasMora (0 si no hay).</summary>
    Function FdblTasaVigente(adtmFecha As Date) As Double
    ''' <summary>FechaDesde de la última fila de OriTasasMora (GCDTMFECHANULA si no hay).</summary>
    Function FdtmFechaDesdeUltima() As Date
    ''' <summary>Agrega una fila de tasa anual (fracción) desde la fecha indicada.</summary>
    Sub SRegistreTasa(adtmFechaDesde As Date, adblTasaAnual As Double)
End Interface

Friend Class ErrorIbcApiException
    Inherits Exception
    Private ReadOnly MentCodigoHttp As Integer
    Friend Sub New(astrMensaje As String, aentCodigoHttp As Integer)
        MyBase.New(astrMensaje)
        MentCodigoHttp = aentCodigoHttp
    End Sub
    Friend ReadOnly Property EntCodigoHttp As Integer
        Get
            Return MentCodigoHttp
        End Get
    End Property
End Class
```

- [ ] **Step 4: Implementar el cliente y el contenedor de la key**

`OriIntCon\clsIbcApiCliente.vb`:

```vb
Imports System.Net
Imports System.Net.Http
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

Friend Class ClsIbcApiCliente
    Implements IIbcProveedor
#Region "Definiciones"
    Private Const MCENTTIMEOUTSEG As Integer = 60
    Private ReadOnly MstrUrlBase As String
    Private ReadOnly MstrApiKey As String
    Private ReadOnly MentReintentos As Integer
    Private ReadOnly MentEsperaMs As Integer
    Private ReadOnly MhttpCliente As HttpClient
#End Region
#Region "Constructores"
    Friend Sub New(astrUrlBase As String, astrApiKey As String,
            Optional ahndManejador As HttpMessageHandler = Nothing,
            Optional aentReintentos As Integer = 3, Optional aentEsperaMs As Integer = 2000)
        MstrUrlBase = astrUrlBase.TrimEnd("/"c)
        MstrApiKey = astrApiKey
        MentReintentos = aentReintentos
        MentEsperaMs = aentEsperaMs
        If IsNothing(ahndManejador) Then
            MhttpCliente = New HttpClient
        Else
            MhttpCliente = New HttpClient(ahndManejador)
        End If
        MhttpCliente.Timeout = TimeSpan.FromSeconds(MCENTTIMEOUTSEG)
    End Sub
#End Region
#Region "Procedimientos y funciones"
    Friend Function FlstConsulteCertificados(aentCantidad As Integer) As List(Of StcIbcCertificado) _
            Implements IIbcProveedor.FlstConsulteCertificados
        If String.IsNullOrWhiteSpace(MstrApiKey) Then
            Throw New ErrorIbcApiException("La API key del IBC no está configurada en esta versión.", 0)
        End If
        Dim lstrJson = Task.Run(Function() FstrObtengaJsonAsync(
                MstrUrlBase & "/api/certificados?n=" & aentCantidad.ToString(
                Globalization.CultureInfo.InvariantCulture))).GetAwaiter().GetResult()
        Return FlstParsee(lstrJson)
    End Function

    Private Async Function FstrObtengaJsonAsync(astrUrl As String) As Task(Of String)
        Dim lentCodigo = 0
        For i As Integer = 0 To MentReintentos
            Try
                Using lreqSolicitud As New HttpRequestMessage(HttpMethod.Get, astrUrl)
                    lreqSolicitud.Headers.Add("X-Api-Key", MstrApiKey)
                    Using lrspRespuesta = Await MhttpCliente.SendAsync(lreqSolicitud).ConfigureAwait(False)
                        lentCodigo = CInt(lrspRespuesta.StatusCode)
                        If lrspRespuesta.IsSuccessStatusCode Then
                            Return Await lrspRespuesta.Content.ReadAsStringAsync().ConfigureAwait(False)
                        End If
                        If lentCodigo = 401 Then
                            Throw New ErrorIbcApiException("La API rechazó la API key (401).", 401)
                        End If
                        If lentCodigo = 404 Then
                            Return "[]"
                        End If
                    End Using
                End Using
            Catch ex As TaskCanceledException
                lentCodigo = 0     ' timeout: se reintenta
            Catch ex As HttpRequestException
                lentCodigo = 0     ' red caída: se reintenta
            End Try
            If i < MentReintentos Then
                Await Task.Delay(MentEsperaMs * (i + 1)).ConfigureAwait(False)
            End If
        Next
        Throw New ErrorIbcApiException("La API del IBC no respondió (código " & lentCodigo.ToString &
                ") tras " & (MentReintentos + 1).ToString & " intentos.", lentCodigo)
    End Function

    Private Shared Function FlstParsee(astrJson As String) As List(Of StcIbcCertificado)
        Dim llstCertificados As New List(Of StcIbcCertificado)
        Dim ljarArreglo As JArray = JArray.Parse(astrJson)
        For Each ljobItem As JObject In ljarArreglo
            Dim lstcCert As New StcIbcCertificado
            lstcCert.StrIdfile = ljobItem.Value(Of String)("idfile")
            lstcCert.DtmFechaCertificado = ljobItem.Value(Of Date)("fechaCertificado").Date
            lstcCert.DtmFechaDesde = ljobItem.Value(Of Date)("fechaDesde").Date
            lstcCert.DtmFechaHasta = ljobItem.Value(Of Date)("fechaHasta").Date
            lstcCert.DblIbc = ClsIbcCalculo.FdblIbcComoFraccion(ljobItem.Value(Of Double)("ibc"))
            llstCertificados.Add(lstcCert)
        Next
        Return llstCertificados
    End Function
#End Region
End Class
```

Notas: `Value(Of Date)` de Newtonsoft respeta el `DateParseHandling` por defecto; el `.Date` descarta la hora. Si un valor `ibc` viene como string, `Value(Of Double)` lo convierte con cultura invariante.

`OriIntCon\IbcApiKey.sample.vb` (versionado):

```vb
' Copia este archivo a IbcApiKey.vb (ignorado por Git) y reemplaza MSTRCLAVEOFUSCADA
' con el resultado de: powershell -File scripts\build\Ofusca-IbcKey.ps1 -Key "<tu key>"
Friend Module MdefIbcApiKey
    Private Const MSTRCLAVEOFUSCADA As String = ""
    Friend Function FstrApiKey() As String
        If MSTRCLAVEOFUSCADA.Length = 0 Then Return String.Empty
        Dim lbytDatos = Convert.FromBase64String(MSTRCLAVEOFUSCADA)
        Dim lbytMascara = System.Text.Encoding.UTF8.GetBytes("OrionPlus-IBC")
        For i As Integer = 0 To lbytDatos.Length - 1
            lbytDatos(i) = lbytDatos(i) Xor lbytMascara(i Mod lbytMascara.Length)
        Next
        Return System.Text.Encoding.UTF8.GetString(lbytDatos)
    End Function
End Module
```

Crear `scripts\build\Ofusca-IbcKey.ps1`:

```powershell
param([Parameter(Mandatory)][string]$Key)
$mask = [Text.Encoding]::UTF8.GetBytes("OrionPlus-IBC")
$data = [Text.Encoding]::UTF8.GetBytes($Key)
for ($i = 0; $i -lt $data.Length; $i++) { $data[$i] = $data[$i] -bxor $mask[$i % $mask.Length] }
[Convert]::ToBase64String($data)
```

Agregar a `.gitignore`: `OriIntCon/IbcApiKey.vb`.

- [ ] **Step 5: Incluir archivos en los `.vbproj`**

En `OrionCopL.vbproj`: `<Compile Include="clsIbcInterfaces.vb" />`.
En `OriIntCon.vbproj`:

```xml
<Compile Include="clsIbcApiCliente.vb" />
<Compile Include="IbcApiKey.vb" Condition="Exists('IbcApiKey.vb')" />
<Compile Include="IbcApiKey.sample.vb" Condition="!Exists('IbcApiKey.vb')" />
```

Agregar `Newtonsoft.Json` ya referenciado en `OriIntCon` (verificar que `Imports Newtonsoft.Json.Linq` compila). Agregar `<Assembly: InternalsVisibleTo("OrionCopL.Tests")>` en `OriIntCon\acOriIntCon.vb` (abrir el archivo y seguir el formato de sus otras líneas `InternalsVisibleTo`; si no tiene ninguna, agregar `Imports System.Runtime.CompilerServices` arriba).

- [ ] **Step 6: Ejecutar y verificar que pasa**

Run: build + vstest del Task 1, Step 5.
Expected: `Passed: 19`.

- [ ] **Step 7: Prueba de contrato contra la API local (manual, no bloqueante)**

Con la API local levantada (`http://localhost:5080`; si no responde, `5100`) y una key de QA (`orion-QA`) **escrita por el usuario, no por el agente**: ejecutar una vez `FlstConsulteCertificados(1)` desde una prueba temporal o el Immediate window y confirmar que devuelve 1 certificado con `DblIbc` como fracción. No commitear la key.

- [ ] **Step 8: Commit**

```bash
git add OrionCopL/clsIbcInterfaces.vb OrionCopL/OrionCopL.vbproj OriIntCon/clsIbcApiCliente.vb OriIntCon/IbcApiKey.sample.vb OriIntCon/OriIntCon.vbproj OriIntCon/acOriIntCon.vb scripts/build/Ofusca-IbcKey.ps1 .gitignore tests/OrionCopL.Tests/IbcApiClienteTests.vb
git commit -m "feat: cliente HTTP de la API IBC con reintentos y key ofuscada"
```

---

### Task 4: Esquema de BD (versión 268) y almacén de certificados

**Files:**
- Modify: `C:\FuentesPanorama.Net\Trunk\Comunes\PanDat\XmlBd\OrionCop_Net.xml` (repo `comunes`: rama y PR propios)
- Create: `OrionCopL\clsIbcCertificado.vb`
- Modify: `OrionCopL\OrionCopL.vbproj`
- Modify: `OrionCopL\clsCentroUtilidadOriCop.vb` (solo en Task 5)

**Interfaces:**
- Consumes: `IIbcAlmacen`, `StcIbcCertificado` (Tasks 2-3).
- Produces: tabla `OriIbcCertificados` (`IdCarpeta` SHORT, `IdCentroUtil` SHORT, `Idfile` STRING(20), `FechaCertificado` DATE, `FechaDesde` DATE, `FechaHasta` DATE, `Ibc` DOUBLE, `FechaDescarga` DATE); PK (`IdCarpeta`,`IdCentroUtil`,`Idfile`); `ClsIbcCertificado` (entidad) y `ClsIbcAlmacenBd : IIbcAlmacen`.

- [ ] **Step 1: Ubicar y copiar el formato de una columna de texto y de una tabla con llave de texto**

Run: `Grep` en `OrionCop_Net.xml` por `TipoDato="STRING"` para copiar el atributo de longitud exacto (p. ej. `Longitud="30"`), y por `Version="267"` para el encabezado `<BD ...>`.
Expected: formato exacto de columna de texto y ubicación del atributo `Version`. Usar **ese** formato en el Step 2.

- [ ] **Step 2: Agregar la tabla y subir la versión**

En `OrionCop_Net.xml`, junto a `OriTasasMora` (~línea 1750) agregar:

```xml
      <Tabla Nombre="OriIbcCertificados" Vinculada="N">
        <Columnas>
          <Columna Nombre="FechaCertificado" Requerido="S" TipoDato="DATE" />
          <Columna Nombre="FechaDescarga" Requerido="S" TipoDato="DATE" />
          <Columna Nombre="FechaDesde" Requerido="S" TipoDato="DATE" />
          <Columna Nombre="FechaHasta" Requerido="S" TipoDato="DATE" />
          <Columna Nombre="IdCarpeta" Requerido="S" TipoDato="SHORT" />
          <Columna Nombre="IdCentroUtil" Requerido="S" TipoDato="SHORT" />
          <Columna Nombre="Idfile" Requerido="S" TipoDato="STRING" Longitud="20" />
          <Columna Nombre="Ibc" Requerido="S" TipoDato="DOUBLE" />
        </Columnas>
        <Indices>
          <Indice Nombre="PK_OriIbcCertificados" Principal="S" Unico="S">
            <ColumnasIndice>
              <ColumnaIndice Nombre="IdCarpeta" Orden="ASC" />
              <ColumnaIndice Nombre="IdCentroUtil" Orden="ASC" />
              <ColumnaIndice Nombre="Idfile" Orden="ASC" />
            </ColumnasIndice>
          </Indice>
          <Indice Nombre="Vigencia" Principal="N" Unico="N">
            <ColumnasIndice>
              <ColumnaIndice Nombre="IdCarpeta" Orden="ASC" />
              <ColumnaIndice Nombre="IdCentroUtil" Orden="ASC" />
              <ColumnaIndice Nombre="FechaDesde" Orden="ASC" />
            </ColumnasIndice>
          </Indice>
        </Indices>
      </Tabla>
```

Cambiar `Version="267"` a `Version="268"` en el `<BD Nombre="OrionCop_Net" ...>`. (Si en el Step 1 el formato de texto difiere de `Longitud="20"`, usar el encontrado.)

- [ ] **Step 3: Entidad y almacén**

`OrionCopL\clsIbcCertificado.vb` — copiar la estructura de `clsTasaMora.vb` (constructor `enuNavegable`/`enuUnico`, constructor de colección, propiedades identificadoras `HstrNombreTabla`, `HenuIdClase`, `HstrNombreClase`, `ColPropiedades`) y adaptar:

- Constante `MCSTRNOMBRETABLA = "OriIbcCertificados"`.
- `HenuIdClase`: **agregar** un valor nuevo `enuIbcCertificado` al enum `EnuIdClasesPanDef` (localizar con `Grep "enuInteresMora"` en `Comunes\PanL` u `OriWin`; usar el siguiente valor libre) — leer el uso de `HenuIdClase` en `ClsCBObjetoPan` antes de editar para confirmar que no exige registro adicional (permisos/log).
- Propiedades: `ObjIdCarpetaIbcShr` (`ClsIdCarpetaShr`), `ObjIdCentroUtilIbcShr` (`ClsIdCentroUtilShr`), `ObjIdfileIbcStr` (nueva `ClsIdfileIbcStr`: copia de `ClsCodigoCuentaCrStr` en `clsServicio.vb:1243`, `HstrNombre="Idfile"`, `HshrLongitud=20`, `HenuTipoValor=EnuTipoValor.enuString`, `HstrNombreCampoBd="Idfile"`, `HblnEsRequerido=True`, `HblnEsLlave=True`, `HbytPosicionLlave=2`, `SValide`: `ClsPanorama.FblnEsValidoString(HobjValorNew, 1, ShrLongitud, BlnEsRequerido)`), y clases de propiedad `ClsFechaCertificadoIbcDtm`, `ClsFechaDescargaIbcDtm`, `ClsFechaDesdeIbcDtm`, `ClsFechaHastaIbcDtm` (copiar `ClsFechaDesdeTasaMoraDtm` **sin** las reglas de encadenamiento: `SValide` → `ClsPanorama.FblnEsValidoFecha(HobjValorNew, DateSerial(2000,1,1), DateSerial(2100,12,31), HblnEsRequerido)`) y `ClsIbcDbl` (copia de `ClsTasaMoraDbl` **sin** `EvnPreSetValor`; `SValide`: `ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 1, HblnEsRequerido, EnuTipoValor.enuDouble)`).
- `SActualice` en creación: fijar `ObjIdCarpetaIbcShr`/`ObjIdCentroUtilIbcShr` con `GshrIdCarpeta`/`GshrIdCentroUtil` y llamar `MyBase.SActualice` (patrón de `ClsTasaMora.SActualice`, sin `SNumereObj`).

`ClsIbcAlmacenBd` (mismo archivo):

```vb
Friend Class ClsIbcAlmacenBd
    Implements IIbcAlmacen

    Friend Function FlstCertificados() As List(Of StcIbcCertificado) Implements IIbcAlmacen.FlstCertificados
        Dim llstResultado As New List(Of StcIbcCertificado)
        Dim lstrCampos As String() = {ClsIdfileIbcStr.SstrNombreCampoBd, ClsFechaCertificadoIbcDtm.SstrNombreCampoBd,
                ClsFechaDesdeIbcDtm.SstrNombreCampoBd, ClsFechaHastaIbcDtm.SstrNombreCampoBd,
                ClsIbcDbl.SstrNombreCampoBd}
        Dim lstrIndice(,) As String = {{ClsFechaDesdeIbcDtm.SstrNombreCampoBd, "ASC"}}
        Dim ldtbCert = ClsPanorama.FdtbDataTable(ClsIbcCertificado.SstrNombreTabla, lstrCampos, lstrIndice,
                ClsOrionCop.StrFiltroUbicacion)
        For Each ldrwCert As DataRow In ldtbCert.Rows
            Dim lstcCert As New StcIbcCertificado
            lstcCert.StrIdfile = ClsPanorama.FobjValorCampo(ldrwCert(ClsIdfileIbcStr.SstrNombreCampoBd),
                    EnuTipoValor.enuString).ToString
            lstcCert.DtmFechaCertificado = ClsPanorama.FobjValorCampo(
                    ldrwCert(ClsFechaCertificadoIbcDtm.SstrNombreCampoBd), EnuTipoValor.enuDate)
            lstcCert.DtmFechaDesde = ClsPanorama.FobjValorCampo(
                    ldrwCert(ClsFechaDesdeIbcDtm.SstrNombreCampoBd), EnuTipoValor.enuDate)
            lstcCert.DtmFechaHasta = ClsPanorama.FobjValorCampo(
                    ldrwCert(ClsFechaHastaIbcDtm.SstrNombreCampoBd), EnuTipoValor.enuDate)
            lstcCert.DblIbc = ClsPanorama.FobjValorCampo(ldrwCert(ClsIbcDbl.SstrNombreCampoBd),
                    EnuTipoValor.enuDouble)
            llstResultado.Add(lstcCert)
        Next
        Return llstResultado
    End Function

    ''' <summary>Guarda solo los certificados que aún no existen localmente (por Idfile).</summary>
    Friend Sub SGuarde(alstCertificados As IEnumerable(Of StcIbcCertificado)) Implements IIbcAlmacen.SGuarde
        Dim lcolExistentes As New HashSet(Of String)(FlstCertificados().Select(Function(x) x.StrIdfile))
        Dim lobjCert As New ClsIbcCertificado(EnuModoInstanciaObjDef.enuNavegable)
        For Each lstcCert As StcIbcCertificado In alstCertificados
            If Not lcolExistentes.Contains(lstcCert.StrIdfile) Then
                lobjCert.SCree()
                lobjCert.ObjIdfileIbcStr.ObjValorPro = lstcCert.StrIdfile
                lobjCert.ObjFechaCertificadoIbcDtm.ObjValorPro = lstcCert.DtmFechaCertificado
                lobjCert.ObjFechaDesdeIbcDtm.ObjValorPro = lstcCert.DtmFechaDesde
                lobjCert.ObjFechaHastaIbcDtm.ObjValorPro = lstcCert.DtmFechaHasta
                lobjCert.ObjIbcDbl.ObjValorPro = lstcCert.DblIbc
                lobjCert.ObjFechaDescargaIbcDtm.ObjValorPro = Date.Today
                lobjCert.SActualice(True)
                lobjCert.SNormaliceEstado(True)
                lcolExistentes.Add(lstcCert.StrIdfile)
            End If
        Next
    End Sub
End Class
```

Agregar `<Compile Include="clsIbcCertificado.vb" />` en `OrionCopL.vbproj`.

- [ ] **Step 4: Verificar el patrón de creación por código**

Leer en `C:\FuentesPanorama.Net\Trunk\Comunes\PanL\clsCBObjetoPan.vb` la secuencia de creación (`SCree`, asignar propiedades, `SActualice(True)`, `SNormaliceEstado(True)`), que es la que usa `winTasasMora.xaml.vb:123-190` a través de la ventana base. Si el nombre o los parámetros difieren de los del Step 3, corregir `SGuarde` (y el Task 6) **antes** de compilar.

- [ ] **Step 5: Compilar la solución**

Run: `& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" OrionCop.Net.sln -restore /p:Configuration=Release /p:Platform=x64 /v:minimal`
Expected: `Build succeeded` sin errores nuevos.

- [ ] **Step 6: Verificación de esquema (manual, BD de desarrollo)**

Abrir Orión Plus (build de desarrollo) contra una BD de prueba con versión 267. Expected: el actualizador crea `OriIbcCertificados` y la versión queda en 268; `DESCRIBE OriIbcCertificados;` muestra las 8 columnas y la PK. Guardar un certificado de prueba desde el Immediate window/prueba temporal con `New ClsIbcAlmacenBd().SGuarde(...)` y verificar la fila con `SELECT * FROM OriIbcCertificados;`, incluido un `Idfile` de 10 dígitos (`9920261260`).

- [ ] **Step 7: Commits (dos repos)**

```bash
# repo orioncop
git add OrionCopL/clsIbcCertificado.vb OrionCopL/OrionCopL.vbproj
git commit -m "feat: entidad y almacen de certificados IBC"
```
Para `Comunes` (`C:\FuentesPanorama.Net\Trunk\Comunes`): hacer commit del XML en **su propia rama** y PR según su `AGENTS.md`; el PR de `orioncop` debe declarar la dependencia (impacto cross-repo en la plantilla, sección 5).

---

### Task 5: Campos de parametrización en el centro de utilidad

**Files:**
- Modify: `OrionCopL\clsCentroUtilidadOriCop.vb` (propiedades ~84, colección ~139, clases de propiedad al final del archivo, junto a `ClsFechaUltCausacionGralDtm` ~3193)
- Modify: `..\Comunes\PanDat\XmlBd\OrionCop_Net.xml` (tabla `OriCentrosUtilidadOriCop`, línea 124+)
- Test: `tests\OrionCopL.Tests\IbcParametrosTests.vb`

**Interfaces:**
- Consumes: `EnuModoInteres`, `ClsIbcCalculo` (Tasks 1-2).
- Produces en `ClsCentroUtilOriCop` (accesible como `GobjParametros`):
  - `ObjModoInteresByt As ClsModoInteresByt` (valor `EnuModoInteres`, defecto `None`)
  - `ObjTasaFijaDeseadaDbl As ClsTasaFijaDeseadaDbl` (fracción 0..1; nunca se sobrescribe por el tope)
  - `ObjFactorVariableDbl As ClsFactorVariableDbl` (0 < f ≤ 1.5)
  - `ClsIbcCalculo.FdblTasaFijaPermitida(adblDigitada, adblIbc, ByRef ablnSeAjusto) As Double` y `FdblFactorPermitido(adblDigitado, ByRef ablnSeAjusto) As Double`.

- [ ] **Step 1: Escribir las pruebas que fallan (ajuste al máximo al digitar)**

`tests\OrionCopL.Tests\IbcParametrosTests.vb`:

```vb
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class IbcParametrosTests
    <TestMethod>
    Public Sub Tasa_fija_sobre_el_tope_se_ajusta_al_maximo_y_avisa()
        Dim lblnAjusto As Boolean
        Dim ldbl = ClsIbcCalculo.FdblTasaFijaPermitida(0.35, 0.1938, lblnAjusto)
        Assert.AreEqual(0.2907, ldbl, 0.0000001)
        Assert.IsTrue(lblnAjusto)
    End Sub

    <TestMethod>
    Public Sub Tasa_fija_dentro_del_tope_no_se_toca()
        Dim lblnAjusto As Boolean
        Dim ldbl = ClsIbcCalculo.FdblTasaFijaPermitida(0.24, 0.1938, lblnAjusto)
        Assert.AreEqual(0.24, ldbl, 0.0000001)
        Assert.IsFalse(lblnAjusto)
    End Sub

    <TestMethod>
    Public Sub Factor_sobre_uno_coma_cinco_se_ajusta_a_uno_coma_cinco()
        Dim lblnAjusto As Boolean
        Assert.AreEqual(1.5, ClsIbcCalculo.FdblFactorPermitido(1.8, lblnAjusto), 0.0000001)
        Assert.IsTrue(lblnAjusto)
    End Sub

    <TestMethod>
    Public Sub Factor_definido_por_el_usuario_dentro_del_limite_se_respeta()
        Dim lblnAjusto As Boolean
        Assert.AreEqual(1.3, ClsIbcCalculo.FdblFactorPermitido(1.3, lblnAjusto), 0.0000001)
        Assert.IsFalse(lblnAjusto)
    End Sub

    <TestMethod>
    Public Sub Tasa_fija_negativa_o_cero_no_es_permitida_y_queda_en_cero()
        Dim lblnAjusto As Boolean
        Assert.AreEqual(0, ClsIbcCalculo.FdblTasaFijaPermitida(-0.1, 0.1938, lblnAjusto), 0.0000001)
    End Sub
End Class
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: build del Task 1, Step 5. Expected: `'FdblTasaFijaPermitida' no es un miembro de 'ClsIbcCalculo'`.

- [ ] **Step 3: Implementar las reglas de digitación en `ClsIbcCalculo`**

```vb
    ''' <summary>
    ''' Tasa fija que se acepta al digitar: nunca negativa ni mayor al tope. Si se ajusta, avisa.
    ''' </summary>
    Friend Shared Function FdblTasaFijaPermitida(adblDigitada As Double, adblIbc As Double,
            ByRef ablnSeAjusto As Boolean) As Double
        Dim ldblTope = FdblTopeMaximo(adblIbc)
        Dim ldblPermitida = Math.Max(0, Math.Min(adblDigitada, ldblTope))
        ablnSeAjusto = ldblPermitida <> adblDigitada
        Return ldblPermitida
    End Function

    Friend Shared Function FdblFactorPermitido(adblDigitado As Double, ByRef ablnSeAjusto As Boolean) As Double
        Dim ldblPermitido = Math.Min(adblDigitado, CDBLFACTORMAXIMO)
        ablnSeAjusto = ldblPermitido <> adblDigitado
        Return ldblPermitido
    End Function
```

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: build + vstest del Task 1, Step 5. Expected: `Passed: 24`.

- [ ] **Step 5: Columnas nuevas y propiedades del centro de utilidad**

En `OrionCop_Net.xml`, tabla `OriCentrosUtilidadOriCop`:

- Dentro de `<Comandos>` (línea 125+), agregar (mismo formato de la línea 136):

```xml
          <Comando Tipo="Propio" Condicion="VBD&lt;268" Secuencia="Inicio" Accion="ACM"
              Parametros="ModoInteres,TasaFijaDeseada,FactorVariable" />
```
  (Confirmar leyendo `mActualizaBD.vb` cómo `ACM` recibe **varias** columnas; en la línea 136-137 recibe una. Si solo admite una, agregar tres `Comando` con una columna cada uno.)
- Dentro de `<Columnas>`, en orden alfabético con las demás:

```xml
          <Columna Nombre="FactorVariable" Requerido="S" TipoDato="DOUBLE" ValorDefecto="0" />
          <Columna Nombre="ModoInteres" Requerido="S" TipoDato="BYTE" ValorDefecto="0" />
          <Columna Nombre="TasaFijaDeseada" Requerido="S" TipoDato="DOUBLE" ValorDefecto="0" />
```

En `clsCentroUtilidadOriCop.vb`: agregar `Friend ReadOnly Property ObjModoInteresByt As New ClsModoInteresByt(Me)` (y las otras dos) junto a la línea 84, `HcolPropiedades.Add(...)` junto a la 139, y al final las tres clases de propiedad copiando `ClsFechaUltCausacionGralDtm` (líneas 3193-3230) con:

```vb
Friend Class ClsModoInteresByt
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "ModoInteres"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Modo de interés de mora"
        HenuTipoValor = EnuTipoValor.enuByte
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoEnumByte(HobjValorNew, EnuModoInteres.None,
                EnuModoInteres.EnuVariable, HblnEsRequerido)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then Return "" Else Return HobjValorPro.ToString
    End Function
End Class

Friend Class ClsTasaFijaDeseadaDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "TasaFijaDeseada"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Tasa fija deseada"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, 1, HblnEsRequerido,
                EnuTipoValor.enuDouble)
        If HblnEsValido Then HobjValorNew = Math.Round(HobjValorNew, 6)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then Return "" Else Return HobjValorPro.ToString
    End Function
End Class

Friend Class ClsFactorVariableDbl
    Inherits ClsCBPropiedad
    Private Const MCSTRNOMBRECAMPOBD As String = "FactorVariable"
    Public Sub New(aobjPadre As ClsCBObjetoPan)
        MyBase.New(aobjPadre)
        HstrNombre = "Factor variable"
        HenuTipoValor = EnuTipoValor.enuDouble
        HstrNombreCampoBd = MCSTRNOMBRECAMPOBD
        HblnRegistrarLogCambio = True
    End Sub
    Public Overrides Sub SValide()
        HblnEsValido = ClsPanorama.FblnEsValidoNumero(HobjValorNew, 0, ClsIbcCalculo.CDBLFACTORMAXIMO,
                HblnEsRequerido, EnuTipoValor.enuDouble)
        If HblnEsValido Then HobjValorNew = Math.Round(HobjValorNew, 4)
    End Sub
    Friend Shared ReadOnly Property SstrNombreCampoBd As String
        Get
            Return MCSTRNOMBRECAMPOBD
        End Get
    End Property
    Public Overrides Function ToString() As String
        If IsNothing(HobjValorPro) Then Return "" Else Return HobjValorPro.ToString
    End Function
End Class
```

(`EnuTipoValor.enuByte` y `ClsPanorama.FblnEsValidoEnumByte` se usan ya en `acOrionCopL.vb` (`ClsVerEFacEnt`) y `clsServicio`; si `enuByte` no existe con ese nombre, usar el que emplee `ObjModoCausaInteresesByt` en `clsServicio.vb`.)

- [ ] **Step 6: Compilar y verificar que el centro de utilidad abre**

Run: compilación de la solución del Task 4, Step 5. Luego abrir Orión Plus en desarrollo, entrar a un centro de utilidad existente y verificar que carga sin error y que los tres campos valen `0` (modo `None`), sin cambiar el comportamiento actual.
Expected: `Build succeeded`; el centro abre normal; `SELECT ModoInteres, TasaFijaDeseada, FactorVariable FROM OriCentrosUtilidadOriCop;` devuelve `0,0,0`.

- [ ] **Step 7: Commit**

```bash
git add OrionCopL/clsCentroUtilidadOriCop.vb OrionCopL/clsIbcCalculo.vb tests/OrionCopL.Tests/IbcParametrosTests.vb
git commit -m "feat: campos de parametrizacion de interes por centro de utilidad"
```
(El XML de `Comunes` va en su rama/PR, junto con el del Task 4.)

---

### Task 6: Escritura de la tasa en `OriTasasMora` (`ClsIbcTasasMoraBd`)

**Files:**
- Create: `OrionCopL\clsIbcTasasMoraBd.vb`
- Modify: `OrionCopL\OrionCopL.vbproj`
- Test: `tests\OrionCopL.Tests\IbcTasaMoraCulturaTests.vb`

**Interfaces:**
- Consumes: `IIbcTasasMora` (Task 3), `ClsTasaMora`, `GobjParametros.FdblTasaMoraFecha`, `GobjParametros.FdtbTasasMora`.
- Produces: `ClsIbcTasasMoraBd : IIbcTasasMora`; helper puro `ClsIbcCalculo.FdblTasaMensualParaAsignar(adblTasaAnual) As Double` (= `Math.Round(anual/12, 8)`) y `ClsIbcCalculo.FblnTasaCoincide(adblEsperada, adblLeida) As Boolean` (tolerancia 0.000002).

- [ ] **Step 1: Escribir las pruebas que fallan**

`tests\OrionCopL.Tests\IbcTasaMoraCulturaTests.vb`:

```vb
Imports System.Globalization
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class IbcTasaMoraCulturaTests
    <TestMethod>
    Public Sub Mensual_a_asignar_es_anual_entre_doce()
        Assert.AreEqual(0.024225, ClsIbcCalculo.FdblTasaMensualParaAsignar(0.2907), 0.000000001)
    End Sub

    <TestMethod>
    Public Sub Coincide_dentro_de_tolerancia_y_no_coincide_con_cero()
        Assert.IsTrue(ClsIbcCalculo.FblnTasaCoincide(0.2907, 0.290700001))
        Assert.IsFalse(ClsIbcCalculo.FblnTasaCoincide(0.2907, 0))
    End Sub

    <TestMethod>
    Public Sub Texto_de_tasa_para_el_framework_usa_punto_decimal_sin_importar_la_cultura()
        Dim lcultOriginal = Threading.Thread.CurrentThread.CurrentCulture
        Try
            Threading.Thread.CurrentThread.CurrentCulture = New CultureInfo("es-CO")
            ' Val() y el parser del framework esperan punto: el helper debe formatear con cultura invariante
            Assert.AreEqual("0.024225", ClsIbcCalculo.FstrTasaMensualInvariante(0.2907))
        Finally
            Threading.Thread.CurrentThread.CurrentCulture = lcultOriginal
        End Try
    End Sub
End Class
```

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: build del Task 1, Step 5. Expected: `'FdblTasaMensualParaAsignar' no es un miembro de 'ClsIbcCalculo'`.

- [ ] **Step 3: Implementar los helpers puros**

En `ClsIbcCalculo`:

```vb
    Friend Shared Function FdblTasaMensualParaAsignar(adblTasaAnual As Double) As Double
        Return Math.Round(adblTasaAnual / 12, 8)
    End Function

    Friend Shared Function FstrTasaMensualInvariante(adblTasaAnual As Double) As String
        Return FdblTasaMensualParaAsignar(adblTasaAnual).ToString("0.########",
                Globalization.CultureInfo.InvariantCulture)
    End Function

    Friend Shared Function FblnTasaCoincide(adblEsperada As Double, adblLeida As Double) As Boolean
        Return Math.Abs(adblEsperada - adblLeida) <= 0.000002
    End Function
```

- [ ] **Step 4: Implementar `ClsIbcTasasMoraBd`**

`OrionCopL\clsIbcTasasMoraBd.vb`:

```vb
Friend Class ClsIbcTasasMoraBd
    Implements IIbcTasasMora

    Friend Function FdblTasaVigente(adtmFecha As Date) As Double Implements IIbcTasasMora.FdblTasaVigente
        Return GobjParametros.FdblTasaMoraFecha(adtmFecha)
    End Function

    Friend Function FdtmFechaDesdeUltima() As Date Implements IIbcTasasMora.FdtmFechaDesdeUltima
        Dim lobjTasa As New ClsTasaMora(EnuModoInstanciaObjDef.enuNavegable)
        Return lobjTasa.FdtmFechaDesdeUltima
    End Function

    ''' <summary>
    ''' Agrega una fila a OriTasasMora. ClsTasaMoraDbl interpreta lo asignado como tasa MENSUAL
    ''' vencida y la convierte a anual simple (x12), por eso se asigna anual/12 y luego se relee.
    ''' </summary>
    Friend Sub SRegistreTasa(adtmFechaDesde As Date, adblTasaAnual As Double) Implements IIbcTasasMora.SRegistreTasa
        Dim lobjTasa As New ClsTasaMora(EnuModoInstanciaObjDef.enuNavegable)
        Dim lcultOriginal = Threading.Thread.CurrentThread.CurrentCulture
        Try
            ' El framework usa Val()/ToString() con la cultura del hilo: forzar punto decimal
            Threading.Thread.CurrentThread.CurrentCulture = Globalization.CultureInfo.InvariantCulture
            lobjTasa.SCree()
            lobjTasa.ObjFechaDesdeTasaMoraDtm.ObjValorPro = adtmFechaDesde
            lobjTasa.ObjFechaHastaTasaMoraDtm.ObjValorPro = Date.Today
            lobjTasa.ObjTasaMoraDbl.ObjValorPro = ClsIbcCalculo.FdblTasaMensualParaAsignar(adblTasaAnual)
            lobjTasa.SActualice(True)
            lobjTasa.SNormaliceEstado(True)
        Finally
            Threading.Thread.CurrentThread.CurrentCulture = lcultOriginal
        End Try
        ' Verificación: lo guardado debe coincidir con lo pedido (evita un 0 silencioso)
        Dim ldblLeida = GobjParametros.FdblTasaMoraFecha(adtmFechaDesde.AddDays(1))
        If Not ClsIbcCalculo.FblnTasaCoincide(adblTasaAnual, ldblLeida) Then
            Throw New ErrorInesperadoPanLException("La tasa de mora guardada (" & ldblLeida.ToString &
                    ") no coincide con la calculada (" & adblTasaAnual.ToString & ").")
        End If
    End Sub
End Class
```

Agregar `<Compile Include="clsIbcTasasMoraBd.vb" />`. Nota: `FdblTasaMoraFecha(f)` consulta `f - 1`; por eso se lee con `FechaDesde + 1`.

- [ ] **Step 5: Ejecutar pruebas puras**

Run: build + vstest del Task 1, Step 5. Expected: `Passed: 27`.

- [ ] **Step 6: Verificación manual con BD de desarrollo (cultura es-CO)**

Con Windows en cultura `es-CO` (coma decimal) y una BD de prueba con al menos una tasa en `OriTasasMora`: ejecutar `New ClsIbcTasasMoraBd().SRegistreTasa(<hoy>, 0.2907)` desde una prueba temporal/Immediate window. Expected: nueva fila con `TasaMora = 0.2907` (±0.000001), la fila anterior con `FechaHasta = hoy - 1`; **sin** excepción. Repetir con cultura `en-US`. Si aparece la excepción de verificación, el forzado de cultura no fue suficiente: investigar `ClsTasaMoraDbl.ClsTasaMoraDbl_evnPreSetValor` antes de seguir.

- [ ] **Step 7: Commit**

```bash
git add OrionCopL/clsIbcTasasMoraBd.vb OrionCopL/OrionCopL.vbproj OrionCopL/clsIbcCalculo.vb tests/OrionCopL.Tests/IbcTasaMoraCulturaTests.vb
git commit -m "feat: escritura de la tasa IBC en OriTasasMora con verificacion"
```

---

### Task 7: Orquestador de sincronización y hook del Cierre de mes

**Files:**
- Create: `OrionCopL\clsIbcSincroniza.vb`
- Modify: `OrionCopL\OrionCopL.vbproj`, `OrionCopL\clsOrionCop.vb` (~4650), `OrionCopIU\MWOrionCop.xaml.vb` (registro del proveedor)
- Test: `tests\OrionCopL.Tests\IbcSincronizaTests.vb`

**Interfaces:**
- Consumes: `IIbcProveedor`, `IIbcAlmacen`, `IIbcTasasMora`, `ErrorIbcApiException`, `ClsIbcCalculo`, `EnuModoInteres`.
- Produces:
  - `ClsIbcSincroniza.New(aobjProveedor As IIbcProveedor, aobjAlmacen As IIbcAlmacen, aobjTasas As IIbcTasasMora)`
  - `Friend Function FblnSincronice(adtmFecha As Date, aenuModo As EnuModoInteres, adblTasaFijaDeseada As Double, adblFactor As Double, ablnForzarConsulta As Boolean, adtmHoy As Date, ByRef astrMens As String) As Boolean` — True si al terminar `OriTasasMora` refleja la tasa que corresponde a la fecha (o si el modo es `None`); False si no se pudo (mensaje en `astrMens`).
  - `ClsOrionCop.SobjProveedorIbc As IIbcProveedor` (compartido; lo asigna la UI al arrancar, sin red) y `Friend Function FblnSincronizaIbcCierre(adtmFecha As Date, ByRef astrMens As String) As Boolean`.

- [ ] **Step 1: Escribir las pruebas que fallan (con dobles de prueba)**

`tests\OrionCopL.Tests\IbcSincronizaTests.vb`:

```vb
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class IbcSincronizaTests
    Private Class ProveedorFalso
        Implements IIbcProveedor
        Friend Lista As New List(Of StcIbcCertificado)
        Friend Falla As Boolean = False
        Friend Llamadas As Integer = 0
        Public Function FlstConsulteCertificados(aentCantidad As Integer) As List(Of StcIbcCertificado) _
                Implements IIbcProveedor.FlstConsulteCertificados
            Llamadas += 1
            If Falla Then Throw New ErrorIbcApiException("API caída", 503)
            Return Lista
        End Function
    End Class

    Private Class AlmacenFalso
        Implements IIbcAlmacen
        Friend Lista As New List(Of StcIbcCertificado)
        Public Function FlstCertificados() As List(Of StcIbcCertificado) Implements IIbcAlmacen.FlstCertificados
            Return Lista
        End Function
        Public Sub SGuarde(alstCertificados As IEnumerable(Of StcIbcCertificado)) Implements IIbcAlmacen.SGuarde
            For Each lstc In alstCertificados
                If Not Lista.Any(Function(x) x.StrIdfile = lstc.StrIdfile) Then Lista.Add(lstc)
            Next
        End Sub
    End Class

    Private Class TasasFalsas
        Implements IIbcTasasMora
        Friend Vigente As Double = 0
        Friend UltimaDesde As Date = #2026-01-01#
        Friend Registros As New List(Of Tuple(Of Date, Double))
        Public Function FdblTasaVigente(adtmFecha As Date) As Double Implements IIbcTasasMora.FdblTasaVigente
            Return Vigente
        End Function
        Public Function FdtmFechaDesdeUltima() As Date Implements IIbcTasasMora.FdtmFechaDesdeUltima
            Return UltimaDesde
        End Function
        Public Sub SRegistreTasa(adtmFechaDesde As Date, adblTasaAnual As Double) Implements IIbcTasasMora.SRegistreTasa
            Registros.Add(Tuple.Create(adtmFechaDesde, adblTasaAnual))
            Vigente = adblTasaAnual
            UltimaDesde = adtmFechaDesde
        End Sub
    End Class

    Private Shared Function FstcCert(astrId As String, adtmDesde As Date, adtmHasta As Date,
            adblIbc As Double) As StcIbcCertificado
        Dim lstc As New StcIbcCertificado
        lstc.StrIdfile = astrId : lstc.DtmFechaDesde = adtmDesde : lstc.DtmFechaHasta = adtmHasta
        lstc.DblIbc = adblIbc : lstc.DtmFechaCertificado = adtmDesde.AddDays(-3)
        Return lstc
    End Function

    Private Shared Function FobjSinc(aobjP As ProveedorFalso, aobjA As AlmacenFalso, aobjT As TasasFalsas) As ClsIbcSincroniza
        Return New ClsIbcSincroniza(aobjP, aobjA, aobjT)
    End Function

    ' Fecha de "hoy" para las reglas de FechaDesde: se inyecta como parámetro de la función
    Private ReadOnly MdtmHoy As Date = #2026-09-28#

    <TestMethod>
    Public Sub Modo_none_no_consulta_ni_escribe_y_devuelve_true()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        Dim lstrMens As String = String.Empty
        Assert.IsTrue(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.None, 0, 0,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(0, lobjP.Llamadas)
        Assert.AreEqual(0, lobjT.Registros.Count)
    End Sub

    <TestMethod>
    Public Sub Variable_con_dato_local_vigente_no_llama_a_la_api_y_escribe_ibc_por_factor()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        Dim lstrMens As String = String.Empty
        Assert.IsTrue(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.3,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(0, lobjP.Llamadas)
        Assert.AreEqual(1, lobjT.Registros.Count)
        Assert.AreEqual(0.25194, lobjT.Registros(0).Item2, 0.0000001)
    End Sub

    <TestMethod>
    Public Sub Fijo_baja_al_tope_y_luego_vuelve_a_la_deseada()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.14))     ' tope 21 %
        lobjA.Lista.Add(FstcCert("B", #2026-10-01#, #2026-10-31#, 0.1938))   ' tope 29.07 %
        Dim lobjS = FobjSinc(lobjP, lobjA, lobjT)
        Dim lstrMens As String = String.Empty
        Assert.IsTrue(lobjS.FblnSincronice(#2026-09-28#, EnuModoInteres.EnuFijo, 0.24, 0, False, #2026-09-28#, lstrMens))
        Assert.AreEqual(0.21, lobjT.Vigente, 0.0000001)
        Assert.IsTrue(lobjS.FblnSincronice(#2026-10-15#, EnuModoInteres.EnuFijo, 0.24, 0, False, #2026-10-15#, lstrMens))
        Assert.AreEqual(0.24, lobjT.Vigente, 0.0000001)
    End Sub

    <TestMethod>
    Public Sub No_escribe_si_la_tasa_vigente_ya_es_la_que_corresponde()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        lobjT.Vigente = 0.24
        Dim lstrMens As String = String.Empty
        Assert.IsTrue(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuFijo, 0.24, 0,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(0, lobjT.Registros.Count)
    End Sub

    <TestMethod>
    Public Sub Api_caida_y_sin_dato_local_devuelve_false_con_mensaje_y_no_escribe()
        Dim lobjP As New ProveedorFalso With {.Falla = True}
        Dim lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.3,
                False, MdtmHoy, lstrMens))
        StringAssert.Contains(lstrMens, "API caída")
        Assert.AreEqual(0, lobjT.Registros.Count)
    End Sub

    <TestMethod>
    Public Sub Api_caida_pero_con_dato_local_vigente_continua()
        Dim lobjP As New ProveedorFalso With {.Falla = True}
        Dim lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        Dim lstrMens As String = String.Empty
        Assert.IsTrue(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.3,
                True, MdtmHoy, lstrMens))    ' forzar consulta falla, pero hay dato local vigente
        Assert.AreEqual(1, lobjT.Registros.Count)
    End Sub

    <TestMethod>
    Public Sub Sin_certificado_que_cubra_la_fecha_tras_consultar_devuelve_false()
        Dim lobjP As New ProveedorFalso
        lobjP.Lista.Add(FstcCert("FUT", #2026-10-01#, #2026-10-31#, 0.2))
        Dim lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.3,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(1, lobjA.Lista.Count)   ' el certificado consultado se guardó igual
    End Sub

    <TestMethod>
    Public Sub No_escribe_si_la_fecha_desde_calculada_supera_hoy()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        lobjT.UltimaDesde = MdtmHoy        ' última fila es de hoy => FechaDesde nueva = mañana > hoy
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.3,
                False, MdtmHoy, lstrMens))
        Assert.AreEqual(0, lobjT.Registros.Count)
        Assert.IsFalse(String.IsNullOrEmpty(lstrMens))
    End Sub

    <TestMethod>
    Public Sub Factor_invalido_devuelve_false()
        Dim lobjP As New ProveedorFalso, lobjA As New AlmacenFalso, lobjT As New TasasFalsas
        lobjA.Lista.Add(FstcCert("A", #2026-09-01#, #2026-09-30#, 0.1938))
        Dim lstrMens As String = String.Empty
        Assert.IsFalse(FobjSinc(lobjP, lobjA, lobjT).FblnSincronice(MdtmHoy, EnuModoInteres.EnuVariable, 0, 1.8,
                False, MdtmHoy, lstrMens))
    End Sub
End Class
```

La firma final (con `adtmHoy` inyectable para probar el tope de `FechaDesde`) es:
`FblnSincronice(adtmFecha, aenuModo, adblTasaFijaDeseada, adblFactor, ablnForzarConsulta, adtmHoy, ByRef astrMens)`. Actualizar también el bloque **Interfaces** de este task con esta firma.

- [ ] **Step 2: Ejecutar y verificar que falla**

Run: build del Task 1, Step 5. Expected: `'ClsIbcSincroniza' no está declarado`.

- [ ] **Step 3: Implementar `ClsIbcSincroniza`**

`OrionCopL\clsIbcSincroniza.vb`:

```vb
Friend Class ClsIbcSincroniza
#Region "Definiciones"
    Private Const MCENTCERTIFICADOSACONSULTAR As Integer = 12
    Private ReadOnly MobjProveedor As IIbcProveedor
    Private ReadOnly MobjAlmacen As IIbcAlmacen
    Private ReadOnly MobjTasas As IIbcTasasMora
#End Region
#Region "Constructores"
    Friend Sub New(aobjProveedor As IIbcProveedor, aobjAlmacen As IIbcAlmacen, aobjTasas As IIbcTasasMora)
        MobjProveedor = aobjProveedor
        MobjAlmacen = aobjAlmacen
        MobjTasas = aobjTasas
    End Sub
#End Region
#Region "Procedimientos y funciones"
    ''' <summary>
    ''' Deja OriTasasMora con la tasa que corresponde a la fecha según el modo del centro de utilidad.
    ''' Devuelve True si quedó al día (o si el modo es None); False con mensaje si no se pudo.
    ''' </summary>
    Friend Function FblnSincronice(adtmFecha As Date, aenuModo As EnuModoInteres, adblTasaFijaDeseada As Double,
            adblFactor As Double, ablnForzarConsulta As Boolean, adtmHoy As Date, ByRef astrMens As String) As Boolean
        astrMens = String.Empty
        If aenuModo = EnuModoInteres.None Then Return True
        If aenuModo = EnuModoInteres.EnuVariable AndAlso Not ClsIbcCalculo.FblnFactorValido(adblFactor) Then
            astrMens = "El factor de interés debe ser mayor que 0 y máximo " &
                    ClsIbcCalculo.CDBLFACTORMAXIMO.ToString & "."
            Return False
        End If
        Dim lstcCert As StcIbcCertificado = Nothing
        Dim lblnTieneCert = ClsIbcCalculo.FblnCertificadoVigente(MobjAlmacen.FlstCertificados(), adtmFecha, lstcCert)
        If ablnForzarConsulta OrElse Not lblnTieneCert Then
            Try
                Dim llstNuevos = MobjProveedor.FlstConsulteCertificados(MCENTCERTIFICADOSACONSULTAR)
                MobjAlmacen.SGuarde(llstNuevos)
                lblnTieneCert = ClsIbcCalculo.FblnCertificadoVigente(MobjAlmacen.FlstCertificados(),
                        adtmFecha, lstcCert)
            Catch ex As ErrorIbcApiException
                If Not lblnTieneCert Then
                    astrMens = "No se pudo consultar el IBC y no hay certificado local vigente: " & ex.Message
                    Return False
                End If
            End Try
        End If
        If Not lblnTieneCert Then
            astrMens = "No existe un certificado de IBC vigente para el " & Format(adtmFecha, "yyyy-MM-dd") & "."
            Return False
        End If
        Dim ldblTasa As Double
        If aenuModo = EnuModoInteres.EnuFijo Then
            ldblTasa = ClsIbcCalculo.FdblTasaFijaEfectiva(adblTasaFijaDeseada, lstcCert.DblIbc)
        Else
            ldblTasa = ClsIbcCalculo.FdblTasaVariable(lstcCert.DblIbc, adblFactor)
        End If
        If ldblTasa <= 0 Then
            astrMens = "La tasa de interés calculada es cero; revise la parametrización."
            Return False
        End If
        If ClsIbcCalculo.FblnTasaCoincide(ldblTasa, MobjTasas.FdblTasaVigente(adtmFecha)) Then Return True
        Dim ldtmDesdeUltima = MobjTasas.FdtmFechaDesdeUltima()
        Dim ldtmDesde = If(lstcCert.DtmFechaDesde > ldtmDesdeUltima, lstcCert.DtmFechaDesde,
                ldtmDesdeUltima.AddDays(1))
        If ldtmDesde > adtmHoy Then
            astrMens = "No se puede registrar una tasa nueva con fecha " & Format(ldtmDesde, "yyyy-MM-dd") &
                    " porque es posterior a hoy. Registre la tasa manualmente en Tasas de Mora."
            Return False
        End If
        MobjTasas.SRegistreTasa(ldtmDesde, ldblTasa)
        Return True
    End Function
#End Region
End Class
```

Agregar `<Compile Include="clsIbcSincroniza.vb" />` al `.vbproj`.

- [ ] **Step 4: Ejecutar y verificar que pasa**

Run: build + vstest del Task 1, Step 5. Expected: `Passed: 36`. Si `Fijo_baja_al_tope_y_luego_vuelve_a_la_deseada` falla en la segunda llamada, revisar que `TasasFalsas.UltimaDesde` (que `SRegistreTasa` avanza) no supere `adtmHoy` (2026-10-15).

- [ ] **Step 5: Registro del proveedor y hook en el Cierre de mes**

En `clsOrionCop.vb`, junto a las demás propiedades compartidas de `ClsOrionCop`:

```vb
    ''' <summary>Proveedor de certificados IBC; lo asigna la UI al iniciar (solo asignación, sin red).</summary>
    Friend Shared Property SobjProveedorIbc As IIbcProveedor
```

y, en la región "Calculo y causación de intereses de Mora a todas las deudas" (antes de `FblnCausoMoraGeneral`):

```vb
    ''' <summary>
    ''' Deja OriTasasMora al día según el modo de interés del centro. Se ejecuta al iniciar el
    ''' Cierre de mes, antes de leer tasas para causar mora.
    ''' </summary>
    Friend Function FblnSincronizaIbcCierre(adtmFecha As Date, ByRef astrMens As String) As Boolean
        If IsNothing(SobjProveedorIbc) Then
            astrMens = String.Empty
            Return GobjParametros.ObjModoInteresByt.ObjValorPro = EnuModoInteres.None
        End If
        Dim lobjSinc As New ClsIbcSincroniza(SobjProveedorIbc, New ClsIbcAlmacenBd, New ClsIbcTasasMoraBd)
        Return lobjSinc.FblnSincronice(adtmFecha, CType(GobjParametros.ObjModoInteresByt.ObjValorPro, EnuModoInteres),
                GobjParametros.ObjTasaFijaDeseadaDbl.ObjValorPro, GobjParametros.ObjFactorVariableDbl.ObjValorPro,
                False, Date.Today, astrMens)
    End Function
```

Al **inicio** de `FblnCausoMoraGeneral` (línea 4650), **antes** de `GobjPanDat.SControleProcesoObj(True)`:

```vb
        If Not FblnSincronizaIbcCierre(FdtmFechaCausaMoraGeneral(), astrMens) Then
            Return False
        End If
```

Antes de editar: leer `FdtmFechaCausaMoraGeneral` (línea 4779) y confirmar que **no tiene efectos secundarios** (solo calcula la fecha). Si los tuviera, calcular la fecha una sola vez y reutilizarla dentro del `Try`. Leer también `winCausaMora.xaml.vb:275` para confirmar que, cuando `FblnCausoMoraGeneral` devuelve `False` con `astrMens` no vacío, se muestra el mensaje al usuario; si no lo muestra, agregarlo con `MsgBox(lstrMens, MsgBoxStyle.Exclamation, "Interés bancario corriente")`.

En `MWOrionCop.xaml.vb`, en el punto donde se inicializa la ventana principal (tras crear `MobjOrionCop`, junto a otras asignaciones de arranque):

```vb
        ClsOrionCop.SobjProveedorIbc = New ClsIbcApiCliente(CSTRURLAPIIBC, MdefIbcApiKey.FstrApiKey())
```
con `Private Const CSTRURLAPIIBC As String = "https://api-ibc-certificados-sif.onrender.com"`. Es solo una asignación: no hace llamadas de red al abrir la app.

- [ ] **Step 6: Compilar y probar el cierre de mes (manual)**

Run: compilación de solución (Task 4, Step 5) y, en desarrollo:
1. Centro en modo `None` → Cierre de mes idéntico al actual (no llama a la API: verificar en el log/tráfico).
2. Modo Variable con certificado local vigente → `OriTasasMora` recibe fila nueva `IBC × factor`, causación continúa.
3. Modo Variable, sin certificado local y API inaccesible (desconectar red o key inválida) → mensaje claro y **no** se causan intereses.
Expected: los tres comportamientos descritos.

- [ ] **Step 7: Commit**

```bash
git add OrionCopL/clsIbcSincroniza.vb OrionCopL/OrionCopL.vbproj OrionCopL/clsOrionCop.vb OrionCopIU/MWOrionCop.xaml.vb OrionCopIU/winCausaMora.xaml.vb tests/OrionCopL.Tests/IbcSincronizaTests.vb
git commit -m "feat: sincronizacion del IBC al iniciar el Cierre de mes"
```

---

### Task 8: Ventana de parametrización de interés (`winParametrizacionInteres`)

**Files:**
- Create: `OrionCopIU\winParametrizacionInteres.xaml`, `OrionCopIU\winParametrizacionInteres.xaml.vb`
- Modify: `OrionCopIU\OrionCopIU.vbproj` (`<Page Include>` y `<Compile Include ... DependentUpon>`, patrón de `winTasasMora` líneas 351 y 644), `OrionCopIU\winParametrizacion.xaml.vb` (~291, entrada de menú)

**Interfaces:**
- Consumes: `GobjParametros.ObjModoInteresByt/ObjTasaFijaDeseadaDbl/ObjFactorVariableDbl`, `ClsIbcCalculo` (Tasks 1, 2, 5), `ClsIbcAlmacenBd`, `ClsIbcSincroniza` (Tasks 4, 7).
- Produces: ventana modal `WinParametrizacionInteres` (abre con `New WinParametrizacionInteres With {.WinPadre = Me}.ShowDialog()`).

Comportamiento exacto (criterios de aceptación):
1. Radio `Fijo` / `Variable` (y estado inicial según `ObjModoInteresByt`; `None` muestra ambos sin marcar y un texto "Sin parametrizar: la tasa se gestiona manualmente en Tasas de Mora").
2. Muestra el IBC vigente del día (`FblnCertificadoVigente` sobre `ClsIbcAlmacenBd().FlstCertificados()`), el **tope** (`FdblTopeMaximo`) y, para el valor definido, la tasa **anual** y la **mensual** (`FdblMensualParaMostrar`), ambas con `Format(x, "#0.00%")`.
3. Fijo: `txtTasaFija` (anual, %). Al salir del control: `FdblTasaFijaPermitida`; si `ablnSeAjusto` → `MsgBox("No es posible cobrar un interés superior a 1.5 veces el IBC (" & Format(tope,"#0.00%") & "). Se ajustó al máximo permitido.", MsgBoxStyle.Exclamation, "Interés máximo")` y el texto queda en el máximo. Sin certificado local: el tope no se puede calcular → se permite guardar la deseada y se muestra "Sin IBC local: se validará al sincronizar" (la sincronización igual aplica `Min`).
4. Variable: `txtFactor` (p. ej. 1.3). Al salir: `FdblFactorPermitido`; si se ajustó → mismo tipo de mensaje ("El factor no puede superar 1.5.").
5. Botón **Guardar**: asigna `ObjModoInteresByt`, `ObjTasaFijaDeseadaDbl` (solo en Fijo; **nunca** se reemplaza por la efectiva), `ObjFactorVariableDbl` (solo en Variable) sobre `GobjParametros`, `SActualice(True)` y cierra. Botón **Sincronizar ahora**: `ClsIbcSincroniza.FblnSincronice(Date.Today, modo, deseada, factor, True, Date.Today, lstrMens)` con cursor de espera; muestra `lstrMens` o "Tasa actualizada" y refresca.
6. Botón **Cancelar** cierra sin guardar.

- [ ] **Step 1: Leer la ventana plantilla**

Abrir `OrionCopIU\winCopiaSeg.xaml(.vb)` (ventana de Herramientas) y `winParametrizacion.xaml.vb` (~291) para copiar: clase base (`ClsFormInterface` de `OriWin`), `WinPadre`, estilos (`RecToolBarPan`, `RecNotifica`, `RecImgCabeza`), cómo se edita `GobjParametros` (secuencia `SModifique` → asignar → `SActualice(True)` → `SNormaliceEstado(True)`) y qué helpers de mensaje usan. Seguir esos patrones en los Steps 2-3; no inventar controles de otro estilo.

- [ ] **Step 2: Crear XAML y code-behind**

Crear `winParametrizacionInteres.xaml` con: `RadioButton` `optFijo`/`optVariable`; `TextBox` `txtTasaFija`, `txtFactor`; `TextBlock` `lblIbcVigente`, `lblTope`, `lblAnual`, `lblMensual`, `lblAviso`; botones `btnGuardar`, `btnSincronizar`, `btnCancelar`. Code-behind con los handlers de los criterios 1-6, usando exclusivamente `ClsIbcCalculo` para las reglas (sin lógica de tope en el code-behind). Formato de porcentajes con `Format(...)` (cultura del usuario) y lectura con `Val`/`Double.TryParse(..., NumberStyles.Float, CurrentCulture, ...)`; el usuario digita porcentaje ("24" o "24,5") → dividir entre 100.

- [ ] **Step 3: Registrar en el `.vbproj` y en el menú de parametrización**

Agregar `<Compile Include="winParametrizacionInteres.xaml.vb"><DependentUpon>winParametrizacionInteres.xaml</DependentUpon></Compile>` y `<Page Include="winParametrizacionInteres.xaml"><SubType>Designer</SubType><Generator>MSBuild:Compile</Generator></Page>` (copiar los atributos exactos de `winTasasMora` en las líneas 351 y 644). En `winParametrizacion.xaml.vb` (~291), junto a "Abrir Tasas de Mora", agregar la entrada "Interés de mora (IBC)…" que abre la ventana.

- [ ] **Step 4: Compilar**

Run: compilación de solución (Task 4, Step 5). Expected: `Build succeeded`.

- [ ] **Step 5: Verificación manual (desarrollo)**

Con un certificado local vigente `IBC = 19.38 %`:
1. Fijo, digitar `35` → mensaje "No es posible…", el campo queda en `29.07`; Guardar → `TasaFijaDeseada = 0.2907` en BD.
2. Fijo, digitar `24` → sin mensaje; mensual mostrado `2.00%`; Guardar → `TasaFijaDeseada = 0.24`.
3. Variable, factor `1.8` → mensaje, queda `1.5`; factor `1.3` → tasa `25.19 %` (mensual `2.10 %`).
4. **Sincronizar ahora** en Variable → aparece fila nueva en `OriTasasMora` con `TasaMora ≈ 0.25194`.
5. Reabrir la ventana → muestra lo guardado.
Expected: todo según lo anterior.

- [ ] **Step 6: Commit**

```bash
git add OrionCopIU/winParametrizacionInteres.xaml OrionCopIU/winParametrizacionInteres.xaml.vb OrionCopIU/OrionCopIU.vbproj OrionCopIU/winParametrizacion.xaml.vb
git commit -m "feat: ventana de parametrizacion de interes con tope 1.5 x IBC"
```

---

### Task 9: Ventana de consulta IBC y opción en el menú Herramientas

**Files:**
- Create: `OrionCopIU\winConsultaIbc.xaml`, `OrionCopIU\winConsultaIbc.xaml.vb`
- Modify: `OrionCopIU\OrionCopIU.vbproj`, `OrionCopIU\MWOrionCop.xaml.vb` (declaración ~85-93, alta ~300-335, despacho ~1715-1740)

**Interfaces:**
- Consumes: `ClsOrionCop.SobjProveedorIbc`, `ClsIbcAlmacenBd`, `ClsIbcSincroniza`, `ClsIbcTasasMoraBd` (Tasks 3-7).
- Produces: `WinConsultaIbc` (modal) y `MnuConsultaIbc` en `HmnuHerramientas`.

Comportamiento (criterios de aceptación):
1. `DataGrid` con los certificados locales (`FlstCertificados()` ordenados por `DtmFechaDesde` desc.): Idfile, Emisión, Desde, Hasta, IBC (`#0.00%`), Tope (`FdblTopeMaximo`, `#0.00%`).
2. Botón **Consultar ahora**: cursor de espera; `MobjProveedor.FlstConsulteCertificados(12)` → `SGuarde` → refresca la lista y muestra "N certificados nuevos". Errores: `ErrorIbcApiException` → `MsgBox` con el mensaje (401: "la API key no es válida"; 0/503: "el servicio no responde, intente de nuevo; la primera consulta puede tardar hasta 1 minuto").
3. Estado en pantalla: fecha/hora de la última consulta exitosa de la sesión y el certificado vigente hoy.
4. Botón **Aplicar a Tasas de Mora**: ejecuta `FblnSincronice(Date.Today, …, True, …)` con los parámetros del centro y muestra el resultado.

- [ ] **Step 1: Crear ventana** siguiendo la plantilla del Task 8, Step 1 (`winCopiaSeg`), con los criterios 1-4.

- [ ] **Step 2: Menú.** En `MWOrionCop.xaml.vb`:
  - Declarar `Private MnuConsultaIbc As MenuItem` junto a `MnuConsultaSql` (~85-93) (usar el mismo tipo que las vecinas).
  - Crear con el helper que usan las vecinas: `MnuConsultaIbc = FmnuiMenuItem("MnuConsultaIbc", "Consulta _IBC (interés bancario corriente)", "RecMnuItemSec")` (patrón ~565-570).
  - `HmnuHerramientas.Items.Add(MnuConsultaIbc)` tras `MnuConsultaSql` (~300-335).
  - En el `Select Case` de despacho (~1715-1740): `Case "MnuConsultaIbc"` → `Dim lwinIbc As New WinConsultaIbc With {.WinPadre = Me} : lwinIbc.ShowDialog()`.
  - Si el menú se controla por permisos, revisar `EnuIdAccionWin`/`EnuIdVentanaDef` y replicar el tratamiento de `MnuConsultaSql`.

- [ ] **Step 3: Registrar archivos en `OrionCopIU.vbproj`** (mismo patrón que Task 8, Step 3).

- [ ] **Step 4: Compilar.** Run: compilación de solución. Expected: `Build succeeded`.

- [ ] **Step 5: Verificación manual (desarrollo, con key real ingresada por el usuario en `IbcApiKey.vb` local)**

1. Herramientas → "Consulta IBC" abre la ventana con la lista local (vacía la primera vez).
2. **Consultar ahora** (con la API dormida) → tarda hasta ~1 min y luego lista los certificados; el vigente hoy aparece resaltado en el estado.
3. Con red desconectada → mensaje claro, la app no se cierra ni se congela más allá del timeout.
4. **Aplicar a Tasas de Mora** en centro Variable → fila nueva en `OriTasasMora`; en centro `None` → mensaje "modo sin parametrizar".
Expected: todo según lo descrito; la app abre igual de rápido que antes (no hay llamadas de red al arrancar).

- [ ] **Step 6: Commit**

```bash
git add OrionCopIU/winConsultaIbc.xaml OrionCopIU/winConsultaIbc.xaml.vb OrionCopIU/OrionCopIU.vbproj OrionCopIU/MWOrionCop.xaml.vb
git commit -m "feat: ventana y menu de consulta de certificados IBC en Herramientas"
```

---

### Task 10: Verificación integral, documentación y PR

**Files:**
- Modify: `docs/CHANGELOG.md` (entrada de la funcionalidad), `docs/REGLAS_NEGOCIO.md` (regla del tope 1.5 × IBC y sincronización en Cierre de mes) — seguir el formato existente de cada archivo.
- Modify: `.github/pull_request_template.md` no se toca; se usa para el cuerpo de la PR.

- [ ] **Step 1: Ejecutar todas las pruebas**

Run: build + vstest del Task 1, Step 5.
Expected: `Passed: 36` (o más), `Failed: 0`.

- [ ] **Step 2: Compilar la solución completa como en release**

Run: `powershell -NoProfile -File scripts\build\build.ps1` (o el `MSBuild` del Task 4, Step 5).
Expected: build OK y **sin** el proyecto de pruebas en la solución (no afecta el release). Confirmar además que sin `OriIntCon\IbcApiKey.vb` el proyecto compila usando `IbcApiKey.sample.vb`.

- [ ] **Step 3: Recorrer la lista de verificación del spec en desarrollo**

1. Fijo con valor > tope → mensaje y máximo; Variable → fila nueva en `OriTasasMora`.
2. Cierre de mes con API caída y sin dato local → causación abortada con mensaje; con dato local vigente → continúa.
3. Modo `None` → Cierre de mes y app **idénticos** a hoy.
4. Cultura `es-CO` y `en-US` al escribir la tasa.
5. `git status` limpio salvo lo esperado; la key **no** aparece en `git diff` ni en el historial (`git log -p -S"<primeros caracteres de la key>"` no devuelve nada).

- [ ] **Step 4: Documentar** (CHANGELOG y REGLAS_NEGOCIO) y commitear: `git commit -m "docs: changelog y reglas de negocio de la herramienta IBC"`.

- [ ] **Step 5: PR según `AGENTS.md`**

```bash
git status --short --branch
git status
git push -u origin feature/herramienta-ibc
gh pr create --base main --head feature/herramienta-ibc --title "feat: herramienta IBC (interes bancario corriente) y tope 1.5 x IBC" --body-file <cuerpo con la plantilla .github/pull_request_template.md>
```
El cuerpo debe declarar: impacto cross-repo (`comunes`: XML versión 268 en su propio PR, **debe fusionarse primero**), la desviación de la tabla por centro, el plan de rollback (`git revert` + versión de esquema 268 queda con columnas extra inofensivas) y que la key se inyecta localmente. No hacer push a `main`; no crear tag ni release hasta fusionar.

---

## Self-Review

**Spec coverage:** consulta y almacenamiento local (T3, T4) · parametrización Fijo/Variable con factor de usuario ≤ 1.5 (T5, T8) · tope duro con auto-ajuste y deseada conservada (T1, T5, T7, T8) · alimenta `OriTasasMora` (T6, T7) · sincronización solo al iniciar Cierre de mes y manual libre (T7, T9) · bloqueo de causación sin tasa actualizada (T7) · menú Herramientas (T9) · API key embebida no versionada (T3) · mensual como `anual/12` (T1, T8) · pruebas unitarias (T1-T7) · convenciones y flujo de repo (Global Constraints, T10). Sin textos de exoneración (no hay tarea, por diseño).

**Placeholder scan:** las tareas de ventanas WPF (T8, T9) delegan estructura visual a la plantilla `winCopiaSeg` pero fijan controles, reglas y criterios de aceptación; los Steps de investigación (T4-1, T4-4, T5-5 nota `ACM`, T7-5) son lecturas puntuales de código con la decisión ya definida, no huecos de diseño.

**Type consistency:** `StcIbcCertificado` (StrIdfile, DtmFechaCertificado, DtmFechaDesde, DtmFechaHasta, DblIbc) idéntico en T2, T3, T4, T7. `FblnCertificadoVigente` (T2) es la que usan T7 y T8. `FblnSincronice(adtmFecha, aenuModo, adblTasaFijaDeseada, adblFactor, ablnForzarConsulta, adtmHoy, ByRef astrMens)` es la firma final en T7, T8, T9 y en `FblnSincronizaIbcCierre`. `IIbcTasasMora.SRegistreTasa(adtmFechaDesde, adblTasaAnual)` igual en T3, T6, T7.
