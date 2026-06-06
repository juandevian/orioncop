Imports System.Windows.Controls
Public Class WinNotasAjuste
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuIdNota
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsNotaAjusteCuotaAdmin = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomNotAju
    Private MblnPoblandoCombo As Boolean = False
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuNotaAjuste
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneControlRestringido(dgrNovedades)
        Dim lcolControlesLlave As New Collection From {
            cboPref,
            txtIdNota
        }
        SCargueForma(EnuElementosAdicionalesDef.enuImprimir, 1,
                lcolControlesLlave, Nothing, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        txtIdNota.Focus()
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
            Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuNotaAjuste)
            ObjObjetoWin = New ClsNotaAjusteCuotaAdmin(lstrPref)
            If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                ObjObjetoWin.SVayaAlUltimo()
            End If
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuIdNota) = lblIdNotaAjuste
        SPuebleCombos()
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                SLevanteEveNoti("No hay Notas para ser mostradas!", "", 0,
                        EnuSeveridadNot.EnuInformacion)
                cboPref.IsEnabled = False
                txtIdNota.IsEnabled = False
            End If
        End If
        With MobjObjetoWin
            cboPref.SelectedItem = .ObjPrefijo_NotaAjusteStr.ObjValorPro
            txtIdNota.Text = .ObjIdNotaAjusteEnt.ToString
            txtIdCliente.Content = .ObjIdCliente_NotaAjusteDbl.ObjValorPro
            txtNombreCliente.Content = .ObjIdCliente_NotaAjusteDbl.StrNombreCliente
            txtIdPredioAgru.Content = .ObjIdPredioAgrupador_NotaAjusteStr.ToString
            txtIdPredio.Content = .ObjIdPredio_NotaAjusteStr.ObjValorPro
            txtIdAnticipo.Content = .ObjIdAnticipo_NotaAjusteEnt.ObjValorPro
            txtFechaNota.Content = Format(.ObjFecha_NotaAjusteDtm.ObjValorPro, GCSTRFMTFECHASIMPLE)
            txtFechaNotaP.Content = Format(.ObjFecha_NotaAjusteDtm.ObjValorPro, GCSTRFMTFECHASIMPLE)
            txtValorNota.Content = Format(.ObjValor_NotaAjusteDec.ObjValorPro, "c")
            txtValorNotaP.Content = Format(.ObjValor_NotaAjusteDec.ObjValorPro, "c")
        End With
        Title = My.Resources.FichaNAju
        If Not String.IsNullOrEmpty(txtIdNota.Text) Then
            Title &= txtIdNota.Text & My.Resources.De & txtNombreCliente.Content
        End If
        SValide()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            txtUsuarioGenero.Content = MobjObjetoWin.ObjIdUsuario_NotaAjusteStr.ObjValorPro
            SEstablezcaDataContext()
            txtIdNota.SelectAll()
        End If
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.cenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntradaDef.enuIdNota) = .ObjIdNotaAjusteEnt.BlnEsValido
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
        HmnuImprimir = FmnuiMenuItem("MnuImprimir", "Im_primir", "RecMnuItemSec")
        Dim lentPosicion = HmnuAcciones.Items.Count - 2
        Dim lsepSeparad As New Separator
        HmnuAcciones.Items.Insert(lentPosicion, HmnuImprimir)
        HmnuAcciones.Items.Insert(lentPosicion, lsepSeparad)
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SRefresqueWin()
        SInicialiceObjeto()
        MyBase.SRefresqueWin()
        SMuestreDatos()
    End Sub
    Protected Overrides Sub SImprima()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Mouse.OverrideCursor = Cursors.Wait
            If MobjObjetoWin.BlnExiste Then
                SLevanteEveNoti("Imprimiendo", String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Dim lstrPrefNotaCon = String.Empty
                Dim lentIdNotaConPrimera = MobjObjetoWin.ObjIdNotaAjusteEnt.ObjValorPro
                Dim lentIdNotaConUltima = MobjObjetoWin.ObjIdNotaAjusteEnt.ObjValorPro
                Dim lobjParaFact As New ClsParametrosReportesDocs(lstrPrefNotaCon,
                        lentIdNotaConPrimera, lentIdNotaConUltima)
                Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                    .ObjParRepDocs = lobjParaFact,
                    .EnuReporte = EnuReporteDef.enuNotasAjuste
                    }
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
        If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable Then
            If BlnBusquedaOk AndAlso StrResutadosBusqueda.Length > 0 Then
                cboPref.SelectedItem = StrResutadosBusqueda(0)
                txtIdNota.Text = StrResutadosBusqueda(1)
                SAbraNotaAjuste()
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
        Dim lstrTablaSec As String = ClsNotaAjusteCuotaAdmin.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgrupador_NotaAjusteStr.SstrNombreCampoBd,
                                             ClsPrefijo_NotaAjusteStr.SstrNombreCampoBd,
                                            ClsIdNotaAjusteEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_NotaAjusteDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCamposRetornar As String() = {ClsPrefijo_NotaAjusteStr.SstrNombreCampoBd,
                                              ClsIdNotaAjusteEnt.SstrNombreCampoBd}
        Dim lstrFiltro As String = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCamposRetornar, lstrFiltro, False)
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SInicialiceNota()
        If MobjObjetoWin IsNot Nothing AndAlso MobjObjetoWin.BlnExiste Then
            If cboPref.SelectedItem Is Nothing Then
                cboPref.SelectedIndex = 0
            End If
            Dim lstrPref As String = cboPref.SelectedItem
            If MobjObjetoWin.ObjPrefijo_NotaAjusteStr.ToString <> lstrPref Then
                ObjObjetoWin = New ClsNotaAjusteCuotaAdmin(lstrPref)
                MobjObjetoWin = ObjObjetoWin
                If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                    MobjObjetoWin.SVayaAlUltimo()
                End If
            End If
        End If
    End Sub
    Private Sub SPuebleCombos()
        MblnPoblandoCombo = True
        Dim ldrwConst = ClsOrionCop.FdrwPrefDoc(EnuTipoDocOri.EnuNotaAjuste)
        SPuebleComboBox(ldrwConst, cboPref)
        MblnPoblandoCombo = False
    End Sub
    ''' <summary>
    ''' Establece el ToolTip de los Controles a un valor constante
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SEstablezcaDataContext()
        If Not IsNothing(MobjObjetoWin) Then
            tbcNotaAjuste.DataContext = MobjObjetoWin.DtbNovedadesNota
        End If
    End Sub
    Private Sub SAbraNotaAjuste()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                If txtIdNota.Text <> MobjObjetoWin.ObjIdNotaAjusteEnt.ToString Then
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, cboPref.SelectedItem,
                            txtIdNota.Text}
                    MobjObjetoWin.SAbra(lobjValorLlave)
                    SMuestreDatos()
                End If
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
                If lblnNoHayError Then
                    GobjPanDat.SControleProcesoObj(False)
                Else
                    GobjPanDat.SControleProcesoObj(False, True)
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
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
    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        If Not MblnPoblandoCombo Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is ComboBox AndAlso Not HblnSeEstaCerrando Then
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                    If lelmElemento.Name = "cboPref" Then
                        SInicialiceNota()
                        SMuestreDatos()
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdNota.KeyDown
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If e.Key = Key.Return OrElse e.Key = Key.Tab Then
                SAbraNotaAjuste()
            End If
        End If
    End Sub
    Private Sub OnRatonUp(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TabControl Then
            SEstablezcaDataContext()
        End If
    End Sub
#End Region
End Class
