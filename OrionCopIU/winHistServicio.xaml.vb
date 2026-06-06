Public Class WinHistServicio
#Region "Definiciones"
    Private ReadOnly MobjServicio As ClsServicio = Nothing
    Private MobjHistServicio As ClsHistServicio = Nothing
    Private McolHistServicios As Collection = Nothing
#End Region
#Region "Constructores"
    Friend Sub New(aobjServicio As ClsServicio)
        InitializeComponent()
        MobjServicio = aobjServicio
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        SLoad()
    End Sub

    Private Sub DgrHisServicios_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dgrHisServicios.SelectionChanged
        If TypeOf sender Is DataGrid Then
            Dim ldgrSender As DataGrid = sender
            If ldgrSender.Name = "dgrHisServicios" Then
                Dim ldrvFilaActual As DataRowView = ldgrSender.SelectedItem
                If Not IsNothing(ldrvFilaActual) Then
                    Dim lshrOrdinal As Short = ldrvFilaActual(ClsOrdinal_HistServicioShr.SstrNombreCampoBd)
                    SEstablezcaHistServicio(lshrOrdinal.ToString)
                End If
            End If
        End If
    End Sub

    Private Sub SEstablezcaHistServicio(astrOrdinalHisSer As String)
        MobjHistServicio = McolHistServicios(astrOrdinalHisSer)
        Dim ldtbHistModServicio As DataTable = MobjHistServicio.DtbHistModServicio
        cnvModulos.DataContext = ldtbHistModServicio
        Dim ldrwHisModulosServicio As DataRow() = ldtbHistModServicio.Select()
        Dim ldecValorTotal = 0D
        For Each ldrwHisModSer As DataRow In ldrwHisModulosServicio
            Dim ldecValor As Decimal = ClsPanorama.FobjValorCampo(
                    ldrwHisModSer(ClsValor_HistModServicioDec.SstrNombreCampoBd), EnuTipoValor.enuDecimal)
            ldecValorTotal += ldecValor
        Next
        txtVlrTotal.Content = Format(ldecValorTotal, "c")
    End Sub

    Private Sub Btt_Click(sender As Object, e As RoutedEventArgs) Handles bttCerrar.Click
        Me.Close()
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SLoad()
        SMuestreUbicacion()
        McolHistServicios = MobjServicio.ColHistServicios
        cnvServicios.DataContext = MobjServicio.DtbHistServicio
    End Sub

    Private Sub SMuestreUbicacion()
        txtCarpeta.Content = ClsOrionCop.StrNombreCarpetaActual
        txtCentroUtilidad.Content = ClsOrionCop.StrNombreCentroUtilActual
        txtIdServicio.Content = MobjServicio.ObjIdServicioShr.ToString
    End Sub
#End Region
End Class
