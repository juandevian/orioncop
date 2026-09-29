Imports System.Globalization
Imports OPT.OrionP.OrionCopL
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
