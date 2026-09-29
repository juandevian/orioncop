Imports OPT.OrionP.OrionCopL
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass>
Public Class IbcTasaMoraGuardadaTests
    <TestMethod>
    Public Sub Excepcion_de_tasa_guardada_lleva_ordinal_fecha_y_dice_que_si_se_guardo()
        Dim ldtmFecha As New Date(2026, 9, 28)
        Dim lobjInterna As New InvalidOperationException("detalle")
        Dim lobjEx As New ErrorTasaMoraGuardadaException(17, ldtmFecha, "la anterior no quedo cerrada", lobjInterna)
        Assert.AreEqual(17, lobjEx.EntOrdinal)
        Assert.AreEqual(ldtmFecha, lobjEx.DtmFechaDesde)
        Assert.AreSame(lobjInterna, lobjEx.InnerException)
        StringAssert.Contains(lobjEx.Message, "SÍ se guardó")
        StringAssert.Contains(lobjEx.Message, "Ordinal 17")
        StringAssert.Contains(lobjEx.Message, "28/09/2026")
        StringAssert.Contains(lobjEx.Message, "Tasas de Mora")
        StringAssert.Contains(lobjEx.Message, "la anterior no quedo cerrada")
    End Sub
End Class
