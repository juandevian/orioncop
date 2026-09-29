Imports System.Net
Imports System.Net.Http
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq
Imports OPT.OrionP.OrionCopL

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
        Dim lstrUrl = MstrUrlBase & "/api/certificados?n=" & aentCantidad.ToString(
                Globalization.CultureInfo.InvariantCulture)
        Dim lstrJson = Task.Run(Function() FstrObtengaJsonAsync(lstrUrl)).GetAwaiter().GetResult()
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
