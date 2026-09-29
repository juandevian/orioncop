Imports OPT.OrionP.OrionCopL
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
