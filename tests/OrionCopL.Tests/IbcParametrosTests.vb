Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports OPT.OrionP.OrionCopL

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
