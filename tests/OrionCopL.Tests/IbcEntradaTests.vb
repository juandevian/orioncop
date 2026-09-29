Imports System.Globalization
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports OPT.OrionP.OrionCopL

<TestClass>
Public Class IbcEntradaTests
    Private Const CDBLTOL As Double = 0.0000001

    <TestMethod>
    Public Sub Porcentaje_entero_se_convierte_a_fraccion()
        Dim ldbl As Double
        Assert.IsTrue(ClsIbcCalculo.FblnTryParsePorcentaje("24", ldbl))
        Assert.AreEqual(0.24, ldbl, CDBLTOL)
    End Sub

    <TestMethod>
    Public Sub Porcentaje_con_coma_o_punto_decimal_da_el_mismo_valor()
        Dim ldblComa As Double, ldblPunto As Double
        Assert.IsTrue(ClsIbcCalculo.FblnTryParsePorcentaje("24,5", ldblComa))
        Assert.IsTrue(ClsIbcCalculo.FblnTryParsePorcentaje("24.5", ldblPunto))
        Assert.AreEqual(0.245, ldblComa, CDBLTOL)
        Assert.AreEqual(0.245, ldblPunto, CDBLTOL)
    End Sub

    <TestMethod>
    Public Sub Porcentaje_no_depende_de_la_cultura_del_hilo()
        Dim lcultOriginal = Threading.Thread.CurrentThread.CurrentCulture
        Try
            Threading.Thread.CurrentThread.CurrentCulture = New CultureInfo("es-CO")
            Dim ldbl As Double
            Assert.IsTrue(ClsIbcCalculo.FblnTryParsePorcentaje("24.5", ldbl))
            Assert.AreEqual(0.245, ldbl, CDBLTOL)     ' en es-CO "." es separador de miles: no debe leerse como 245
        Finally
            Threading.Thread.CurrentThread.CurrentCulture = lcultOriginal
        End Try
    End Sub

    <TestMethod>
    Public Sub Porcentaje_acepta_signo_y_espacios()
        Dim ldbl As Double
        Assert.IsTrue(ClsIbcCalculo.FblnTryParsePorcentaje("  24 % ", ldbl))
        Assert.AreEqual(0.24, ldbl, CDBLTOL)
    End Sub

    <TestMethod>
    Public Sub Porcentaje_invalido_devuelve_false_y_cero()
        Dim ldbl As Double = 5
        Assert.IsFalse(ClsIbcCalculo.FblnTryParsePorcentaje("abc", ldbl))
        Assert.AreEqual(0, ldbl, CDBLTOL)
        Assert.IsFalse(ClsIbcCalculo.FblnTryParsePorcentaje("", ldbl))
        Assert.IsFalse(ClsIbcCalculo.FblnTryParsePorcentaje("   ", ldbl))
        Assert.IsFalse(ClsIbcCalculo.FblnTryParsePorcentaje(Nothing, ldbl))
        Assert.IsFalse(ClsIbcCalculo.FblnTryParsePorcentaje("Infinity", ldbl))
        Assert.IsFalse(ClsIbcCalculo.FblnTryParsePorcentaje("NaN", ldbl))
    End Sub

    <TestMethod>
    Public Sub Factor_se_lee_sin_dividir_entre_cien()
        Dim ldbl As Double
        Assert.IsTrue(ClsIbcCalculo.FblnTryParseFactor("1,3", ldbl))
        Assert.AreEqual(1.3, ldbl, CDBLTOL)
        Assert.IsTrue(ClsIbcCalculo.FblnTryParseFactor("1.5", ldbl))
        Assert.AreEqual(1.5, ldbl, CDBLTOL)
        Assert.IsFalse(ClsIbcCalculo.FblnTryParseFactor("uno", ldbl))
    End Sub
End Class
