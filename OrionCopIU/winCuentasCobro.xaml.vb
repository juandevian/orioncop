Imports System.ComponentModel
Public Class WinCuentasCobro
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    '
#End Region
    ' Variables
    Private MobjObjetoWin As ClsEstadoCuenta = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomCtasCob
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuCtasCobro
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneControlRestringido(dgrFacturasEstado)
        SAdicioneControlRestringido(chkImprimirDetallado)
        SCargueForma(EnuElementosAdicionalesDef.enuImprimir, 0,
                Nothing, Nothing, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        txtNroCtaCobro.IsEnabled = True
        txtNroCtaCobro.Focus()
    End Sub
    Protected Overrides ReadOnly Property StrNombreVentana As String
        Get
            Return MstrNombreVentana
        End Get
    End Property
    Protected Overrides ReadOnly Property EnuIdVentana As EnuIdVentanaDef
        Get
            Return HenuIdVentana
        End Get
    End Property
    Protected Overrides Sub SInicialiceObjeto()
        If IsNothing(ObjObjetoWin) Then
            ObjObjetoWin = New ClsEstadoCuenta(EnuModoInstanciaObjDef.enuNavegable, False)
            ObjObjetoWin.SVayaAlUltimo()
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        HbttAceptar.TabIndex = 30
        HbttCancelar.TabIndex = 31
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
            SLevanteEveNoti("No hay Cuentas de Cliente para ser mostradas!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            txtIdCliente.IsEnabled = False
        End If
        With MobjObjetoWin
            txtNroCtaCobro.Text = .ObjIdEstadoCuentaEnt.ObjValorPro
            txtIdCliente.Content = Format(.ObjIdCliente_EstadoDbl.ObjValorPro, "#,###,###,###,###")
            If Not IsNothing(.ObjIdCliente_EstadoDbl.ObjCliente) Then
                txtNombreCliente.Content = .ObjIdCliente_EstadoDbl.ObjCliente.ObjNombreCompletoStr.ObjValorPro
            Else
                txtNombreCliente.Content = String.Empty
            End If
            If String.IsNullOrEmpty(.ObjIdPredioAgr_EstadoStr.ObjValorPro) Then
                If MobjObjetoWin.BlnExiste Then
                    txtPredioAgr.Content = GCSTRSINPA
                Else
                    txtPredioAgr.Content = String.Empty
                End If
            Else
                txtPredioAgr.Content = .ObjIdPredioAgr_EstadoStr.ObjValorPro
            End If
            txtFechaEstado.Content = Format(.ObjFechaEstadoDtm.ObjValorPro, GCSTRFMTFECHASIMPLE)
            txtDeudaCapital.Content = Format(.ObjDeudaCapitalDec.ObjValorPro, "c")
            txtDeudaMora.Content = Format(.ObjDeudaIntMoraDec.ObjValorPro, "c")
            txtTotalDeuda.Content = Format(.DecTotalDeuda, "c")
        End With
        SMuestreFacturasEstado()
        SValide()
        Title = My.Resources.FichaCtaCobro
        If Not IsNothing(MobjObjetoWin.ObjIdEstadoCuentaEnt.ObjValorPro) Then
            Title &= MobjObjetoWin.ObjIdEstadoCuentaEnt.ObjValorPro & " " & txtNombreCliente.Content
        End If
    End Sub
    Protected Overrides Sub SValide()
        SHabiliteBotonesTlb()
        SHabiliteImprimir()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        '
    End Sub
    Protected Overrides Sub SConfigureMenuesPropios()
        HmnuImprimir = FmnuiMenuItem("MnuImprimir", "Im_primir", "RecMnuItemSec")
        Dim lentPosicion = HmnuAcciones.Items.Count - 2
        Dim lsepSeparad As New Separator
        HmnuAcciones.Items.Insert(lentPosicion, HmnuImprimir)
        HmnuAcciones.Items.Insert(lentPosicion, lsepSeparad)
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SImprima()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Mouse.OverrideCursor = Cursors.Wait
            If MobjObjetoWin.BlnExiste Then
                SLevanteEveNoti("Imprimiendo", String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Dim lentIdCtaCobroPrimera = MobjObjetoWin.ObjIdEstadoCuentaEnt.ObjValorPro
                Dim lentIdCtaCobroUltima = MobjObjetoWin.ObjIdEstadoCuentaEnt.ObjValorPro
                Dim lobjParaCuentaCobro As New ClsParametrosReportesDocs("",
                        lentIdCtaCobroPrimera, lentIdCtaCobroUltima)
                Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                    .ObjParRepDocs = lobjParaCuentaCobro
                    }
                If chkImprimirDetallado.IsChecked Then
                    lobjRep.EnuReporte = EnuReporteDef.enuCtaCobroDet
                Else
                    lobjRep.EnuReporte = EnuReporteDef.enuCtaCobro
                End If
                lobjRep.SGenereReporte()
            End If
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentNullException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If Not lblnNoHayError Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            Else
                SLevanteEveNoti(String.Empty, String.Empty, 0, EnuSeveridadNot.EnuOk)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub
#End Region
#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
            If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable Then
                Dim lentIdCtaCobro As Integer = CType(StrResultadoBusqueda, Integer)
                MobjObjetoWin.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lentIdCtaCobro})
                SMuestreDatos()
            End If
        End If
    End Sub
    ''' <summary>
    ''' Invalida la funcion "fblnDefinioBusqueda" de la clase base.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        SDefineBusquedaPredioAgr()
        SDefineBusquedaNombreCompleto()
        Return True
    End Function
    Private Sub SDefineBusquedaPredioAgr()
        Dim lstrTablaPri = ClsEstadoCuenta.SstrNombreTabla
        Dim lstrTablaSec = ClsCliente.SstrNombreTabla
        Dim lstrCamposTabPri = {ClsIdPredioAgr_EstadoStr.SstrNombreCampoBd,
                                 ClsIdEstadoCuentaEnt.SstrNombreCampoBd,
                                 ClsFechaEstadoDtm.SstrNombreCampoBd}
        Dim lstrCamposTabSec = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamTabPriRel = {ClsIdCliente_EstadoDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsIdPredioAgr_EstadoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdEstadoCuentaEnt.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil
        lstrFiltro &= " AND " & ClsIdFactura_EstadoEnt.SstrNombreCampoBd & " = 0"
        lstrFiltro &= " AND (" & ClsDeudaCapitalDec.SstrNombreCampoBd & " > 0 " & " OR " &
                ClsDeudaIntMoraDec.SstrNombreCampoBd & " > 0)"
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, True)
    End Sub
    Private Sub SDefineBusquedaNombreCompleto()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsEstadoCuenta.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgr_EstadoStr.SstrNombreCampoBd,
                                            ClsFechaEstadoDtm.SstrNombreCampoBd,
                                            ClsIdEstadoCuentaEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_EstadoDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdEstadoCuentaEnt.SstrNombreCampoBd
        Dim lstrFiltro As String = "S." & StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND S." & StrCampoCentroUtil & " = " & GshrIdCentroUtil
        lstrFiltro &= " AND " & ClsIdFactura_EstadoEnt.SstrNombreCampoBd & " = 0"
        lstrFiltro &= " AND (" & ClsDeudaCapitalDec.SstrNombreCampoBd & " > 0 " & " OR " &
                ClsDeudaIntMoraDec.SstrNombreCampoBd & " > 0)"
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTablaPri, lstrTablaSec,
                        lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                        lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, False)
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SMuestreFacturasEstado()
        Dim ldtbFactEstado = MobjObjetoWin.DtbFacturasEstado
        dgrFacturasEstado.DataContext = ldtbFactEstado
        SOrdeneDataGrid(dgrFacturasEstado, dgrFacturasEstado.Columns(1),
                ClsIdFacturaVivaEnt.SstrNombreCampoBd,
                ListSortDirection.Ascending)

    End Sub
    Private Sub SAbraCuentaCobro()
        With MobjObjetoWin
            Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, txtNroCtaCobro.Text}
            MobjObjetoWin.SAbra(lobjValorLlave)
            SMuestreDatos()
        End With
    End Sub
    Private Sub SHabiliteImprimir()
        SHabiliteBotonTlb(True, HbttImprimir)
        SHabiliteMenuItem(True, HmnuImprimir)
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox Then
                SAbraCuentaCobro()
            End If
        End If
    End Sub
    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
                dgrFacturasEstado.MouseRightButtonUp
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            Dim lstrPrefijo As String
            Dim lentIdObjeto As Integer
            ldrvFilaActual = ldgrActual.SelectedItem
            If Not IsNothing(ldrvFilaActual) Then
                lstrPrefijo = ldrvFilaActual("PrefijoFactViva")
                lentIdObjeto = ldrvFilaActual("IdFacturaViva")
                SAbraFactura(lstrPrefijo, lentIdObjeto)
            End If
        End If
    End Sub
    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtNroCtaCobro.KeyDown
        If e.Key = Key.Return Then
            If txtNroCtaCobro.Focus Then
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                    SAbraCuentaCobro()
                End If
            End If
        End If
    End Sub
#End Region
End Class
