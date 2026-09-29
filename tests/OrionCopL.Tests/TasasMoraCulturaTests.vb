Imports System.Globalization
Imports OPT.OrionP.OrionCopL
Imports Microsoft.VisualStudio.TestTools.UnitTesting

''' <summary>
''' FdtbTasasMora reemplaza la FechaHasta de la última fila con hoy. Antes lo hacía con el texto
''' Format(Today, "dd/MM/yyyy") en una columna DateTime, que el DataTable interpreta con la cultura del hilo:
''' con cultura MM/dd fallaba los días 13-31 y cambiaba el día por el mes los días 1-12.
''' </summary>
<TestClass>
Public Class TasasMoraCulturaTests
    Private Shared Function FdtbTablaTasas() As DataTable
        Dim ldtbTabla As New DataTable("TasaMora")
        ldtbTabla.Columns.Add("Ordinal", GetType(Integer))
        ldtbTabla.Columns.Add("FechaDesde", GetType(Date))
        ldtbTabla.Columns.Add("FechaHasta", GetType(Date))
        ldtbTabla.Columns.Add("TasaMora", GetType(Double))
        ldtbTabla.Rows.Add(1, New Date(2026, 1, 1), New Date(2026, 5, 31), 0.28)
        ldtbTabla.Rows.Add(2, New Date(2026, 6, 1), New Date(2026, 6, 30), 0.29)
        Return ldtbTabla
    End Function

    Private Shared Sub SVerifique(astrCultura As String, adtmHoy As Date)
        Dim lcultOriginal = Threading.Thread.CurrentThread.CurrentCulture
        Try
            Threading.Thread.CurrentThread.CurrentCulture = If(astrCultura Is Nothing,
                    CultureInfo.InvariantCulture, New CultureInfo(astrCultura))
            Dim ldtbTabla = FdtbTablaTasas()
            ClsCentroUtilOriCop.SAsigneFechaHastaUltimaTasa(ldtbTabla, adtmHoy)
            Assert.AreEqual(adtmHoy, CDate(ldtbTabla.Rows(1)("FechaHasta")),
                    "Cultura " & If(astrCultura, "Invariant"))
            ' La penúltima fila conserva su valor de la BD
            Assert.AreEqual(New Date(2026, 5, 31), CDate(ldtbTabla.Rows(0)("FechaHasta")))
        Finally
            Threading.Thread.CurrentThread.CurrentCulture = lcultOriginal
        End Try
    End Sub

    <TestMethod>
    Public Sub EsCO_dia_mayor_que_doce()
        SVerifique("es-CO", New Date(2026, 9, 28))
    End Sub

    <TestMethod>
    Public Sub EsCO_dia_menor_o_igual_a_doce()
        SVerifique("es-CO", New Date(2026, 9, 5))
    End Sub

    <TestMethod>
    Public Sub EnUS_dia_mayor_que_doce()
        SVerifique("en-US", New Date(2026, 9, 28))
    End Sub

    <TestMethod>
    Public Sub EnUS_dia_menor_o_igual_a_doce()
        SVerifique("en-US", New Date(2026, 9, 5))
    End Sub

    <TestMethod>
    Public Sub Invariante_dia_mayor_que_doce()
        SVerifique(Nothing, New Date(2026, 9, 28))
    End Sub

    <TestMethod>
    Public Sub Invariante_dia_menor_o_igual_a_doce()
        SVerifique(Nothing, New Date(2026, 9, 5))
    End Sub

    <TestMethod>
    Public Sub Hoy_con_hora_se_guarda_solo_la_fecha()
        Dim ldtbTabla = FdtbTablaTasas()
        ClsCentroUtilOriCop.SAsigneFechaHastaUltimaTasa(ldtbTabla, New Date(2026, 9, 28, 15, 30, 0))
        Assert.AreEqual(New Date(2026, 9, 28), CDate(ldtbTabla.Rows(1)("FechaHasta")))
    End Sub

    <TestMethod>
    Public Sub Tabla_vacia_no_falla()
        Dim ldtbTabla = FdtbTablaTasas()
        ldtbTabla.Rows.Clear()
        ClsCentroUtilOriCop.SAsigneFechaHastaUltimaTasa(ldtbTabla, New Date(2026, 9, 28))
        Assert.AreEqual(0, ldtbTabla.Rows.Count)
    End Sub
End Class
