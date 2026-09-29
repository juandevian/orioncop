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

    <TestMethod>
    Public Sub Tasa_fija_ya_guardada_sobre_el_tope_se_conserva_al_guardar()
        ' deseada 24 %, el IBC bajo y el tope es 21 %: abrir y guardar NO debe bajar la deseada
        Dim lblnAjusto As Boolean
        Assert.AreEqual(0.24, ClsIbcCalculo.FdblTasaFijaAGuardar(0.24, 0.24, 0.14, True, lblnAjusto), 0.0000001)
        Assert.IsFalse(lblnAjusto)
    End Sub

    <TestMethod>
    Public Sub Tasa_fija_nueva_sobre_el_tope_se_ajusta_y_avisa()
        Dim lblnAjusto As Boolean
        Assert.AreEqual(0.21, ClsIbcCalculo.FdblTasaFijaAGuardar(0.3, 0.24, 0.14, True, lblnAjusto), 0.0000001)
        Assert.IsTrue(lblnAjusto)
    End Sub

    <TestMethod>
    Public Sub Tasa_fija_nueva_bajo_el_tope_se_acepta_sin_ajuste()
        Dim lblnAjusto As Boolean
        Assert.AreEqual(0.2, ClsIbcCalculo.FdblTasaFijaAGuardar(0.2, 0.24, 0.14, True, lblnAjusto), 0.0000001)
        Assert.IsFalse(lblnAjusto)
    End Sub

    <TestMethod>
    Public Sub Primera_tasa_fija_sobre_el_tope_se_ajusta_al_maximo()
        Dim lblnAjusto As Boolean
        Assert.AreEqual(0.2907, ClsIbcCalculo.FdblTasaFijaAGuardar(0.35, 0, 0.1938, True, lblnAjusto), 0.0000001)
        Assert.IsTrue(lblnAjusto)
    End Sub

    <TestMethod>
    Public Sub Sin_ibc_local_la_tasa_digitada_se_acepta_sin_ajuste()
        Dim lblnAjusto As Boolean
        Assert.AreEqual(0.35, ClsIbcCalculo.FdblTasaFijaAGuardar(0.35, 0.24, 0, False, lblnAjusto), 0.0000001)
        Assert.IsFalse(lblnAjusto)
        Assert.AreEqual(0.24, ClsIbcCalculo.FdblTasaFijaAGuardar(0.24, 0.24, 0, False, lblnAjusto), 0.0000001)
    End Sub
End Class
