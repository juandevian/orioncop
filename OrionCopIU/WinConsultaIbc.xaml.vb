Public Class WinConsultaIbc
    Inherits Window
#Region "Definiciones"
    Private Const MCSTRTITULO As String = "Consulta IBC"
    Private Const MCENTCERTIFICADOSACONSULTAR As Integer = 12
    Private MdtmUltimaConsulta As Date = Nothing
#End Region
#Region "Constructores"
    Public Sub New()
        InitializeComponent()
    End Sub
#End Region
#Region "Presentación"
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        SRefresque()
    End Sub

    ''' <summary>Muestra los certificados guardados en este equipo y cuál está vigente hoy.</summary>
    Private Sub SRefresque()
        Try
        Dim llstCertificados = New ClsIbcAlmacenBd().FlstCertificados()
        llstCertificados.Sort(Function(a, b) b.DtmFechaDesde.CompareTo(a.DtmFechaDesde))
        Dim ldtbCertificados As New DataTable
        For Each lstrColumna In {"Idfile", "Emision", "Desde", "Hasta", "Ibc", "Tope"}
            ldtbCertificados.Columns.Add(lstrColumna, GetType(String))
        Next
        For Each lstcCert As StcIbcCertificado In llstCertificados
            ldtbCertificados.Rows.Add(lstcCert.StrIdfile, Format(lstcCert.DtmFechaCertificado, "yyyy-MM-dd"),
                    Format(lstcCert.DtmFechaDesde, "yyyy-MM-dd"), Format(lstcCert.DtmFechaHasta, "yyyy-MM-dd"),
                    Format(lstcCert.DblIbc, "#0.00%"), Format(ClsIbcCalculo.FdblTopeMaximo(lstcCert.DblIbc), "#0.00%"))
        Next
        dgrCertificados.ItemsSource = ldtbCertificados.DefaultView
        Dim lstcVigente As StcIbcCertificado
        If ClsIbcCalculo.FblnCertificadoVigente(llstCertificados, Date.Today, lstcVigente) Then
            lblVigente.Text = "Vigente hoy: " & Format(lstcVigente.DblIbc, "#0.00%") & " E.A. (certificado " &
                    lstcVigente.StrIdfile & "). Tope legal: " &
                    Format(ClsIbcCalculo.FdblTopeMaximo(lstcVigente.DblIbc), "#0.00%") & "."
        Else
            lblVigente.Text = "No hay un certificado vigente hoy guardado en este equipo."
        End If
        lblEstado.Text = If(MdtmUltimaConsulta = Nothing, "Aún no se ha consultado en esta sesión.",
                "Última consulta exitosa: " & Format(MdtmUltimaConsulta, "yyyy-MM-dd HH:mm") & ".")
        Catch ex As Exception
            lblVigente.Text = "No se pudieron leer los certificados guardados: " & ex.Message
            lblEstado.Text = String.Empty
        End Try
    End Sub
#End Region
#Region "Eventos"
    Private Sub BttCerrar_Click(sender As Object, e As RoutedEventArgs) Handles bttCerrar.Click
        Close()
    End Sub

    ''' <summary>Consulta la API y guarda los certificados nuevos.</summary>
    Private Sub BttConsultar_Click(sender As Object, e As RoutedEventArgs) Handles bttConsultar.Click
        If IsNothing(ClsOrionCop.SobjProveedorIbc) Then
            MsgBox("El servicio de consulta del IBC no está disponible en esta versión.",
                    MsgBoxStyle.Exclamation, MCSTRTITULO)
            Return
        End If
        Dim lobjAlmacen As New ClsIbcAlmacenBd
        Dim lstrMens = String.Empty
        Dim lenuIcono = MsgBoxStyle.Information
        Mouse.OverrideCursor = Cursors.Wait
        Try
            Dim lentAntes = lobjAlmacen.FlstCertificados().Count
            lobjAlmacen.SGuarde(ClsOrionCop.SobjProveedorIbc.FlstConsulteCertificados(MCENTCERTIFICADOSACONSULTAR))
            Dim lentNuevos = lobjAlmacen.FlstCertificados().Count - lentAntes
            MdtmUltimaConsulta = Now
            lstrMens = If(lentNuevos = 0, "Consulta exitosa: no hay certificados nuevos.",
                    "Consulta exitosa: " & lentNuevos.ToString & " certificado(s) nuevo(s).")
        Catch ex As ErrorIbcApiException
            lenuIcono = MsgBoxStyle.Exclamation
            If ex.EntCodigoHttp = 401 Then
                lstrMens = "La API del IBC rechazó las credenciales de esta versión. Comuníquese con soporte."
            Else
                lstrMens = "El servicio del IBC no responde. Intente de nuevo en unos minutos (la primera consulta " &
                        "puede tardar hasta un minuto). Detalle: " & ex.Message
            End If
        Catch ex As Exception
            lenuIcono = MsgBoxStyle.Exclamation
            lstrMens = "No se pudo completar la consulta del IBC: " & ex.Message
        Finally
            Mouse.OverrideCursor = Nothing
        End Try
        SRefresque()
        MsgBox(lstrMens, lenuIcono, MCSTRTITULO)
    End Sub

    ''' <summary>Deja OriTasasMora al día según el modo de interés del centro de utilidad.</summary>
    Private Sub BttAplicar_Click(sender As Object, e As RoutedEventArgs) Handles bttAplicar.Click
        Dim lstrMens = String.Empty
        Dim lblnOk As Boolean
        Mouse.OverrideCursor = Cursors.Wait
        Try
            lblnOk = ClsOrionCop.FblnSincronizaIbcManual(lstrMens)
        Catch ex As Exception
            lblnOk = False
            lstrMens = ex.Message
        Finally
            Mouse.OverrideCursor = Nothing
        End Try
        SRefresque()
        If lblnOk Then
            MsgBox(lstrMens, MsgBoxStyle.Information, MCSTRTITULO)
        Else
            MsgBox(lstrMens, MsgBoxStyle.Exclamation, MCSTRTITULO)
        End If
    End Sub
#End Region
End Class
