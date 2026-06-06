Imports System.ComponentModel
Public Class WinAnticipos
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuIdAnticipo
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsAnticipo = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomAnti
#End Region

#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuAnticipos
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneControlRestringido(dgrNovedaAnt)
        Dim lcolControlesLlave As New Collection From {
            txtIdAnticipo
        }
        SCargueForma(EnuElementosAdicionalesDef.None, 1,
                lcolControlesLlave, Nothing, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
    End Sub
    Protected Overrides ReadOnly Property StrNombreVentana As String
        Get
            Return MstrNombreVentana
        End Get
    End Property
    Protected Overrides ReadOnly Property Enuidventana As EnuIdVentanaDef
        Get
            Return HenuIdVentana
        End Get
    End Property
    Protected Overrides Sub SInicialiceObjeto()
        If IsNothing(ObjObjetoWin) Then
            ObjObjetoWin = ClsOrionCop.FobjAnticipo
            If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                ObjObjetoWin.SVayaAlUltimo()
            End If
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuIdAnticipo) = lblIdAnticipo
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
            SLevanteEveNoti("No hay Anticipos para ser mostradas!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            txtIdAnticipo.IsEnabled = False
        Else
            With MobjObjetoWin
                txtIdAnticipo.Text = .ObjIdAnticipoEnt.ObjValorPro
                txtIdCliente.Content = .ObjIdCliente_AntDbl.ObjValorPro
                txtNombreCliente.Content = .ObjIdCliente_AntDbl.StrNombreCliente
                txtIdPredioAgru.Content = If(String.IsNullOrEmpty(
                        .ObjIdPredioAgrupador_AntStr.ToString), GCSTRSINPA,
                        .ObjIdPredioAgrupador_AntStr.ToString())
                Dim lstrServicios = .ObjServicios_AntStr.ToString()
                Dim lstrNombreSer = String.Empty
                If lstrServicios.Contains("A") Then
                    lstrNombreSer = "Todos"
                ElseIf lstrServicios = "0" Then
                    lstrNombreSer = "Cuota de Administración"
                ElseIf lstrServicios.Contains(",") Then
                    lstrNombreSer = "Varios Servicios"
                ElseIf IsNumeric(lstrServicios) Then
                    lstrNombreSer = GobjParametros.FstrNombreServicio(lstrServicios)
                End If
                txtServicio.Content = lstrNombreSer
                txtFechaAnticipo.Content = Format(.ObjFechaAnticipoDtm.ObjValorPro, GCSTRFMTFECHASIMPLE)
                txtFechaA.Content = Format(.ObjFechaAnticipoDtm.ObjValorPro, GCSTRFMTFECHASIMPLE)
                Select Case .ObjIdTipoDocOrigen_AntByt.ObjValorPro
                    Case EnuTipoDocOri.EnuReciboCaja
                        txtOrigen.Content = My.Resources.RecCaja
                        txtIdOrigen.Content = .ObjPrefijoDocOrigen_AntStr.ToString & "-" &
                                .ObjIdDocOrigen_AntEnt.ToString
                        lblIdOrigen.Content = My.Resources.IdRecCaja
                    Case EnuTipoDocOri.EnuNotaAjuste
                        txtOrigen.Content = My.Resources.AjuCuota
                        txtIdOrigen.Content = .ObjIdPredio_AntStr.ObjValorPro
                        lblIdOrigen.Content = My.Resources.NombrePredio
                    Case Else
                        txtOrigen.Content = String.Empty
                        txtIdOrigen.Content = String.Empty
                        lblIdOrigen.Content = String.Empty
                End Select
                txtValorAnticipo.Content = Format(.ObjValor_AntDec.ObjValorPro, "c")
                txtValorA.Content = Format(.ObjValor_AntDec.ObjValorPro, "c")
                txtDebitos.Content = Format(.ObjDebitos_AntDec.ObjValorPro, "c")
                txtCreditos.Content = Format(.ObjCreditos_AntDec.ObjValorPro, "c")
                txtSaldo.Content = Format(.DecAnticipoPorAplicar, "c")
            End With
        End If
        SMuestreEstado()
        Title = My.Resources.FichaAnt
        If Not String.IsNullOrEmpty(txtIdAnticipo.Text) Then
            Title &= txtIdAnticipo.Text & My.Resources.De & txtNombreCliente.Content
        End If
        SValide()
        SEstablezcaDataContext()
        txtIdAnticipo.SelectAll()
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.CenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntradaDef.enuIdAnticipo) = .ObjIdAnticipoEnt.BlnEsValido
            End With
        End If
        '
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        SValide()
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        ' Adicionar Menues Reportes
        Dim lmnuReportes As MenuItem = FmnuiMenuItem("MnuReportes", "R_eportes", "RecMnuItemPriInf")
        MenuVen.Items.Insert(3, lmnuReportes)
        Dim lmnuItem = FmnuiMenuItem("MnuRepAnticiposPorAplicar", "Anticipos por Aplicar",
                "RecMnuItemSec")
        lmnuItem.ToolTip = "Genera un Reporte de los Anticipos por Aplicar."
        lmnuReportes.Items.Add(lmnuItem)
        lmnuItem = FmnuiMenuItem("MnuAnticiposPorPredioAgr", "Anticipos por Predio Agrupador",
                "RecMnuItemSec")
        lmnuItem.ToolTip = "Genera un Reporte de los Anticipos por Aplicar por predio " &
                "agrupador."
        lmnuReportes.Items.Add(lmnuItem)
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    '
#End Region

#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
            If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable Then
                MobjObjetoWin.SAbra({GshrIdCarpeta, GshrIdCentroUtil, StrResultadoBusqueda})
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
        SDefineNombreCliente()
        Return True
    End Function
    Private Sub SDefineNombreCliente()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsAnticipo.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgrupador_AntStr.SstrNombreCampoBd,
                                            ClsIdAnticipoEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_AntDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdAnticipoEnt.SstrNombreCampoBd
        Dim lstrFiltro = "S." & StrCampoCarpeta & " = " &
                GshrIdCarpeta.ToString & " AND S." & StrCampoCentroUtil &
                " = " & GshrIdCentroUtil.ToString
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, False)
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SMuestreEstado()
        If MobjObjetoWin.FblnAnticipoReversado Then
            txtEstado.Style = FindResource("RecDocRversado")
        Else
            If MobjObjetoWin.ObjAnuladoBln.ObjValorPro Then
                txtEstado.Style = FindResource("RecDocAnulado")
            Else
                txtEstado.Style = FindResource("RecDocNormal")
            End If
        End If
    End Sub
    Private Sub SEstablezcaDataContext()
        cnvDatosAnticipo.DataContext = MobjObjetoWin.DtbNovedadesAnt
        SOrdeneDataGrid(dgrNovedaAnt, dgrNovedaAnt.Columns(0), ClsIdNovedadAntShr.SstrNombreCampoBd,
                ListSortDirection.Ascending)
    End Sub
    Private Sub SAbraAnticipo()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                If txtIdAnticipo.Text <> MobjObjetoWin.ObjIdAnticipoEnt.ToString Then
                    Dim lobjVlrLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                            txtIdAnticipo.Text}
                    MobjObjetoWin.SAbra(lobjVlrLlave)
                End If
                SMuestreDatos()
                lblnNoHayError = True
            Catch ex As PanLException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As PanDatException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As Exception
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Finally
                If Not lblnNoHayError Then
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Mouse.OverrideCursor = Cursors.Wait
            Try
                GobjPanDat.SControleProcesoObj(True)
                If lelmElemento.Name = "MnuRepAnticiposPorAplicar" Then
                    Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO)
                    With lobjRep
                        .EnuReporte = EnuReporteDef.enuAnticiposPorAplicar
                        .SGenereReporte()
                    End With
                ElseIf lelmElemento.Name = "MnuAnticiposPorPredioAgr" Then
                    Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO)
                    With lobjRep
                        .EnuReporte = EnuReporteDef.enuAnticiposPorPredioAgru
                        .SGenereReporte()
                    End With
                End If
                lblnNoHayError = True
                Catch ex As PanDatException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As PanLException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As Exception
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Finally
                If lblnNoHayError Then
                    GobjPanDat.SControleProcesoObj(False)
                Else
                    GobjPanDat.SControleProcesoObj(False)
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
                Mouse.OverrideCursor = Cursors.Arrow
            End Try
        End If
    End Sub
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        End If
    End Sub
    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdAnticipo.KeyDown
        If e.Key = Key.Return OrElse e.Key = Key.Tab Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando AndAlso
                    MobjObjetoWin.ObjIdAnticipoEnt.ToString() <> txtIdAnticipo.Text Then
                SAbraAnticipo()
            End If
        End If
    End Sub
    Private Sub DgrNovedaAnt_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
            dgrNovedaAnt.MouseRightButtonUp
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If lelmElemento.Name = "dgrNovedaAnt" Then
            Dim ldrvFilaActual As DataRowView = dgrNovedaAnt.SelectedItem
            If Not IsNothing(ldrvFilaActual) Then
                Dim lenuTipoDocOrigen As EnuTipoDocOri = ldrvFilaActual(
                        ClsIdTipoDocOrigen_AntByt.SstrNombreCampoBd)
                Dim lstrPrefDocOri = ldrvFilaActual(ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd)
                Dim lentIdDocOri = ldrvFilaActual(ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd)
                Select Case lenuTipoDocOrigen
                    Case EnuTipoDocOri.EnuNotaAjuste
                        SAbraNotaAju(lstrPrefDocOri, lentIdDocOri)
                    Case EnuTipoDocOri.EnuReciboCaja
                        SAbraRecibo(lstrPrefDocOri, lentIdDocOri)
                    Case EnuTipoDocOri.EnuNotaCon
                        SAbraNotaCon(lstrPrefDocOri, lentIdDocOri)
                    Case EnuTipoDocOri.EnuNotaDevAnt
                        SAbraNotaDevAnt(lstrPrefDocOri, lentIdDocOri)
                    Case EnuTipoDocOri.EnuNotaRevCr
                        SAbraNotaRCr(lstrPrefDocOri, lentIdDocOri)
                End Select
            End If
        End If
    End Sub
    Private Sub Ctrl_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles _
            txtIdOrigen.MouseDoubleClick
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If lelmElemento.Name = "txtIdOrigen" Then
            Dim lstrPrefDocOri = ClsPanorama.FstrPrefijoDcto(txtIdOrigen.Content)
            Dim lentIdDocOri = ClsPanorama.FentIdDcto(txtIdOrigen.Content)
            If MobjObjetoWin.ObjIdTipoDocOrigen_AntByt.ObjValorPro =
                        EnuTipoDocOri.EnuReciboCaja Then
                SAbraRecibo(lstrPrefDocOri, lentIdDocOri)
            ElseIf MobjObjetoWin.ObjIdTipoDocOrigen_AntByt.ObjValorPro =
                        EnuTipoDocOri.EnuNotaAjuste Then
                SAbraNotaAju(lstrPrefDocOri, lentIdDocOri)
            End If
        End If
    End Sub
#End Region
End Class
