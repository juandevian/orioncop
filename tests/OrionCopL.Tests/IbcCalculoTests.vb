Imports OPT.OrionP.OrionCopL
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
