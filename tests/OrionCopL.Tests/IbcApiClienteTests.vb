Imports System.Net
Imports System.Net.Http
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports OPT.OrionP.OrionCopL
Imports OPT.OrionP.OriIntCon

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
