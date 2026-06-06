Imports System.Windows.Controls
Imports System.ComponentModel
Public Class WinNotasAplicaAnt
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuPrefijo
        enuIdNota
    End Enum
#End Region
    ' Variables
    Private MnuEnviarPorCorreo As MenuItem = Nothing
    Private MobjObjetoWin As ClsNotaCon = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomNotAplAnt
    Private MnuEstadoProvEfac As MenuItem = Nothing
    Private MnuReprocesarEDoc As MenuItem = Nothing
    Private MnuEDoc As MenuItem = Nothing
    Private ReadOnly MwinMW As MWOrionCop = Nothing
    Private MblnPoblandoCombo As Boolean = False
#End Region

#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuNotaAplAnt
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneControlRestringido(dgrItemsNotaCon)
        SAdicioneControlRestringido(dgrNovedades)
        Dim lcolControlesLlave As New Collection From {
            cboPref,
            txtIdNota
        }
        SCargueForma(EnuElementosAdicionalesDef.enuImprimir, 2,
                lcolControlesLlave, Nothing, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SHabiliteMenuEFac()
        SHabiliteReprocesarEDoc()
        SCambieNombresDocCobro()
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
            Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuNotaAplicacionAnt)
            ObjObjetoWin = New ClsNotaCon(lstrPref)
            If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                ObjObjetoWin.SVayaAlUltimo()
            End If
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub

    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuPrefijo) = lblIdNotaCon
        StcValidaControl(EnuValidEntradaDef.enuIdNota) = lblIdNotaCon
        SPuebleCombos()
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
            SLevanteEveNoti("No hay Notas para ser mostradas!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            cboPref.IsEnabled = False
            txtIdNota.IsEnabled = False
        Else
            With MobjObjetoWin
                cboPref.SelectedItem = .ObjPrefijo_NotaConStr.ObjValorPro
                txtIdNota.Text = .ObjIdNotaConEnt.ToString()
                txtIdCliente.Content = .ObjIdCliente_NotaConDbl.ObjValorPro
                txtNombreCliente.Content = .ObjIdCliente_NotaConDbl.StrNombreCliente
                txtIdPredioAgru.Content = .ObjIdPredioAgrupador_NotaConStr.ToString
                txtIdAnticipo.Content = .ObjIdAnticipo_NotaConEnt.ObjValorPro
                txtFechaNotaConP.Content = Format(.ObjFecha_NotaConDtm.ObjValorPro, GCSTRFMTFECHASIMPLE)
                txtValorNotaP.Content = Format(.ObjValor_NotaConDec.ObjValorPro, "c")
            End With
        End If
        Title = My.Resources.FichaNApl
        If Not String.IsNullOrEmpty(txtIdNota.Text) Then
            Title &= txtIdNota.Text & My.Resources.De & txtNombreCliente.Content
        End If
        SValide()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            txtNroNotaRCr.Content = MobjObjetoWin.StrNumeroNotaRCr
            SMuestreEstado()
            txtUsuarioGenero.Content = MobjObjetoWin.ObjIdUsuario_NotaConStr.ObjValorPro
            SEstablezcaDataContext()
            SHabiliteMenuEFac()
            SHabiliteReprocesarEDoc()
            txtIdNota.Focus()
        End If
    End Sub

    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.cenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntradaDef.enuPrefijo) = .ObjPrefijo_NotaConStr.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuIdNota) = .ObjIdNotaConEnt.BlnEsValido
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
        MnuEnviarPorCorreo = FmnuiMenuItemPan("MnuEnviarPorCorreo", "_Enviar por eMail", 1,
                "")
        HmnuImprimir = FmnuiMenuItem("MnuImprimir", "Im_primir", "RecMnuItemSec")
        Dim lentPosicion = HmnuAcciones.Items.Count - 2
        Dim lsepSeparad As New Separator
        If ClsPanorama.FblnEmailsHabilitado Then
            HmnuAcciones.Items.Insert(lentPosicion, MnuEnviarPorCorreo)
        End If
        HmnuAcciones.Items.Insert(lentPosicion, HmnuImprimir)
        HmnuAcciones.Items.Insert(lentPosicion, lsepSeparad)
        If GobjParametros.BlnEFacAutorizado AndAlso MobjObjetoWin.BlnEsDocEle Then
            MnuEDoc = FmnuiMenuItemPan("MnuEDoc", "eDocumento", 0, "", True)
            HmnuMiMenu.Items.Insert(2, MnuEDoc)
            MnuEstadoProvEfac = FmnuiMenuItem("MnuEstadoProvEfac", "Estado en Pro_veedor eFactura",
                    "RecMnuItemSec")
            MnuReprocesarEDoc = FmnuiMenuItem("MnuReprocesarEDoc",
                    "_Reprocesar Factura Electrónica", "RecMnuItemSec")
            MnuEDoc.Items.Add(MnuEstadoProvEfac)
            MnuEDoc.Items.Add(MnuReprocesarEDoc)
        End If
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    Protected Overrides Sub SHabiliteMenues()
        MyBase.SHabiliteMenues()
    End Sub

    Protected Overrides Sub SImprima()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Dim lblnPuede = MobjObjetoWin.BlnEstaRegEFac OrElse Not MobjObjetoWin.BlnEsDocEle
            If lblnPuede Then
                Mouse.OverrideCursor = Cursors.Wait
                If MobjObjetoWin.BlnExiste Then
                    SLevanteEveNoti("Imprimiendo", String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    Dim lstrPrefNotaCon = MobjObjetoWin.ObjPrefijo_NotaConStr.ObjValorPro
                    Dim lentIdNotaConPrimera = MobjObjetoWin.ObjIdNotaConEnt.ObjValorPro
                    Dim lentIdNotaConUltima = MobjObjetoWin.ObjIdNotaConEnt.ObjValorPro
                    Dim lobjParaFact As New ClsParametrosReportesDocs(lstrPrefNotaCon,
                            lentIdNotaConPrimera, lentIdNotaConUltima)
                    Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                        .ObjParRepDocs = lobjParaFact,
                        .EnuReporte = EnuReporteDef.enuNotasCon
                        }
                    lobjRep.SGenereReporte()
                End If
            Else
                lstrMens = "La Nota no se puede imprimir porque aún no esta registrada en API de EFac!"
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
            If lblnNoHayError Then
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Else
                    SLevanteEveNoti(String.Empty, String.Empty, 0, EnuSeveridadNot.EnuOk)
                End If
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub

    Protected Overrides Sub SRefresqueWin()
        MyBase.SRefresqueWin()
        SCambieNombresDocCobro()
    End Sub
#End Region

#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable Then
            If BlnBusquedaOk AndAlso StrResutadosBusqueda.Length > 0 Then
                cboPref.SelectedItem = StrResutadosBusqueda(0)
                txtIdNota.Text = StrResutadosBusqueda(1)
                SAbraNotaCon()
            End If
        End If
    End Sub
    ''' <summary>
    ''' Invalida la funcion "fblnDefinioBusqueda" de la clase base.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 
    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        SDefineNombreCliente()
        Return True
    End Function

    Private Sub SDefineNombreCliente()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsNotaCon.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgrupador_NotaConStr.SstrNombreCampoBd,
                                            ClsPrefijo_NotaConStr.SstrNombreCampoBd,
                                            ClsIdNotaConEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_NotaConDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCamposRetornar As String() = {ClsPrefijo_NotaConStr.SstrNombreCampoBd,
                                             ClsIdNotaConEnt.SstrNombreCampoBd}
        Dim lstrFiltro As String = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCamposRetornar, lstrFiltro, False)
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SInicialiceNota()
        If MobjObjetoWin IsNot Nothing Then
            If cboPref.SelectedItem Is Nothing Then
                cboPref.SelectedIndex = 0
            End If
            Dim lstrPref As String = cboPref.SelectedItem
            If MobjObjetoWin.ObjPrefijo_NotaConStr.ToString <> lstrPref Then
                ObjObjetoWin = New ClsNotaCon(lstrPref)
                MobjObjetoWin = ObjObjetoWin
                If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                    MobjObjetoWin.SVayaAlUltimo()
                End If
            End If
        End If
    End Sub

    Private Sub SPuebleCombos()
        MblnPoblandoCombo = True
        Dim ldrwConst = ClsOrionCop.FdrwPrefDoc(EnuTipoDocOri.EnuNotaCon)
        SPuebleComboBox(ldrwConst, cboPref)
        MblnPoblandoCombo = False
    End Sub

    Private Sub SEstablezcaDataContext()
        If Not IsNothing(MobjObjetoWin) Then
            If tbcNotaCon.SelectedIndex = 0 Then
                tbcNotaCon.DataContext = MobjObjetoWin.DtbItemsNotaCon
                SOrdeneDataGrid(dgrItemsNotaCon, dgrItemsNotaCon.Columns(0),
                        ClsIdItemNotaConShr.SstrNombreCampoBd,
                        ListSortDirection.Ascending)
            ElseIf tbcNotaCon.SelectedIndex = 1 Then
                dgrNovedades.DataContext = MobjObjetoWin.DtbNovedadesNotaCon
                SOrdeneDataGrid(dgrNovedades, dgrNovedades.Columns(1),
                        ClsIdNovedadShr.SstrNombreCampoBd, ListSortDirection.Ascending)
            End If
        End If
    End Sub

    Private Sub SMuestreEstado()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If txtNroNotaRCr.Content = "0" Then
                If MobjObjetoWin.ObjAnuladoBln.ObjValorPro Then
                    txtEstado.Style = FindResource("RecDocAnulado")
                Else
                    txtEstado.Style = FindResource("RecDocNormal")
                End If
            Else
                txtEstado.Style = FindResource("RecDocRversado")
            End If
        End If
    End Sub

    Private Sub SAbraNotaCon()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                If txtIdNota.Text <> MobjObjetoWin.ObjIdNotaConEnt.ToString Then
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
            End Try
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End If
    End Sub

    Private Sub SEnvieEmailNotaCon(ByRef astrMens As String)
        If ClsPanorama.FblnEmailsHabilitado Then
            If MobjObjetoWin.ObjClienteNotaCon.ObjRecibeDocsPorEmailBln.ObjValorPro Then
                SEnvieCorreo(EnuTipoCorreoE.EnuNAA,
                        MobjObjetoWin.ObjClienteNotaCon.ObjIdClienteDbl.ObjValorPro,
                        MobjObjetoWin.ObjIdPredioAgrupador_NotaConStr.ObjValorPro,
                        MobjObjetoWin.StrNumeroNotaCon, astrMens)
            Else
                astrMens = "El Cliente de la Nota no tiene habilitado " &
                        "el envio de Documentos por Email!"
            End If
        Else
            astrMens = "Aún no tiene instalada la Aplicación para el envío de Documentos " &
                              "por Email!"
        End If
    End Sub

    Private Sub SCambieNombresDocCobro()
        If GobjParametros.BlnEFacAutorizado Then
            dgtNroFac.Header = "Número Factura"
            dgtItFac.Header = "Item Fac."
            dgrNroFacNov.Header = "Número Factura"
            dgtItFacNov.Header = "Item Fac."
        Else
            dgtNroFac.Header = "Número Cta. Cobro"
            dgtItFac.Header = "Item C.C."
            dgrNroFacNov.Header = "Número Cta. Cobro"
            dgtItFacNov.Header = "Item C.C."
        End If
    End Sub
#End Region

#Region "eFactura"
    Private Sub SHabiliteMenuEFac()
        If MnuEstadoProvEfac IsNot Nothing Then
            MnuEstadoProvEfac.Visibility = Visibility.Collapsed
            MnuReprocesarEDoc.Visibility = Visibility.Collapsed
        End If
        If GobjParametros.BlnEFacAutorizado AndAlso MobjObjetoWin.BlnExiste Then
            If Not HblnCargandoForma AndAlso MobjObjetoWin.BlnEsDocEle Then
                MnuEDoc.Visibility = Visibility.Visible
                Dim lenuEstadoEFac As EnuEstadoEDoc =
                        MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro
                Dim lstrMens = String.Empty, lblnProcesar = True
                MnuEstadoProvEfac.Visibility = Visibility.Visible
                Select Case lenuEstadoEFac
                    Case EnuEstadoEDoc.EnuErrorFtp
                        lstrMens = "El Documento parece tener Problemas en su Información. " &
                                "Ya esta corregido?"
                    Case EnuEstadoEDoc.EnuNoReg
                        MsgBox("El Documento debe ser procesado!", vbOKOnly,
                                "Actualizar Documento!")
                    Case EnuEstadoEDoc.EnuInvalida
                        lstrMens = "El Documento fue rechazado por la DIAN. " &
                                "Ya esta corregido?"
                    Case EnuEstadoEDoc.EnuEnProceso
                        MsgBox("El Documento debe ser actualizado!", vbOKOnly,
                                "Actualizar Documento!")
                        lblnProcesar = True
                    Case EnuEstadoEDoc.EnuRegi
                        If GobjParametros.ObjSubirFacBln.ObjValorPro Then
                            MsgBox("El Documento debe ser actualizado!", vbOKOnly,
                                    "Actualizar Documento!")
                        Else
                            lblnProcesar = False
                        End If
                        MnuReprocesarEDoc.Visibility = Visibility.Visible
                    Case EnuEstadoEDoc.EnuEnviada
                        lblnProcesar = False
                        MnuReprocesarEDoc.Visibility = Visibility.Visible
                    Case Else
                        Exit Select
                End Select
                If lblnProcesar Then
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        lblnProcesar = (MsgBox(lstrMens, vbYesNo, "Registrar Documento?") = vbYes)
                        If lblnProcesar Then
                            If lenuEstadoEFac = EnuEstadoEDoc.EnuInvalida OrElse
                                    lenuEstadoEFac = EnuEstadoEDoc.EnuErrorFtp Then
                                MobjObjetoWin.SHabiliteProcesarEFac()
                            End If
                        End If
                    End If
                End If
            ElseIf MnuEDoc IsNot Nothing Then
                MnuEDoc.Visibility = Visibility.Collapsed
            End If
        End If
    End Sub

    Private Sub SHabiliteReprocesarEDoc()
        If MnuReprocesarEDoc IsNot Nothing Then
            MnuReprocesarEDoc.Visibility = Visibility.Collapsed
        End If
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If Not HblnCargandoForma AndAlso MobjObjetoWin.BlnEsDocEle AndAlso
                    MobjObjetoWin.BlnEsAjusteCuotaAdmin Then
                MnuReprocesarEDoc.Visibility = Visibility.Visible
            End If
        End If
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                Dim lmnuOpcion As MenuItem = lelmElemento
                If lmnuOpcion.Name = "MnuEnviarPorCorreo" Then
                    SEnvieEmailNotaCon(lstrMens)
                ElseIf lmnuOpcion.Name = "MnuEstadoProvEfac" Then
                    Mouse.OverrideCursor = Cursors.Wait
                    SRefresqueWin()
                    Dim lstrIdNotaCon As String
                    If String.IsNullOrEmpty(MobjObjetoWin.ObjPrefijo_NotaConStr.ToString()) Then
                        lstrIdNotaCon = "NCON" & MobjObjetoWin.StrIdObjeto
                    Else
                        lstrIdNotaCon = MobjObjetoWin.StrIdObjeto
                    End If
                    SMuestreEstadoEFac(MobjObjetoWin.ObjCUDocStr.ObjValorPro,
                            EnuTipoDocOri.EnuNotaCr, lstrIdNotaCon, False,
                            MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro, lstrMens)
                    Mouse.OverrideCursor = Cursors.Arrow
                ElseIf lmnuOpcion.Name = "MnuReprocesarEDoc" Then
                    If MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                            EnuEstadoEDoc.EnuNoEDoc AndAlso
                            MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                            EnuEstadoEDoc.EnuNoReg AndAlso
                            MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                             EnuEstadoEDoc.EnuInvalida Then
                        If MobjObjetoWin.ObjFecha_NotaConDtm.ObjValorPro > Today.AddDays(-30) Then
                            If MsgBox("Realmente desea reprocesar esta Nota?", vbYesNo,
                                    "Reprocesar Nota?") = vbYes Then
                                MobjObjetoWin.SPrepareParaReprocesarEfac()
                                SProceseEFac(lstrMens)
                            End If
                        Else
                            lstrMens = "Solo es posible reprocesar un documento si su fecha está " &
                                    "dentro de los 30 dias anteriores al día de hoy!"
                        End If
                    End If
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
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                Else
                    GobjPanDat.SControleProcesoObj(False, True)
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
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
                SAbraNotaCon()
            End If
        End If
    End Sub

    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
                 dgrItemsNotaCon.MouseRightButtonUp, dgrNovedades.MouseRightButtonUp
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If lelmElemento.Name = "dgrItemsNotaCon" OrElse lelmElemento.Name = "dgrNovedades" Then
            Dim ldgrSeleccionada As DataGrid
            ldgrSeleccionada = sender
            Dim ldrvAnticipo As DataRowView = ldgrSeleccionada.SelectedItem
            If Not IsNothing(ldrvAnticipo) Then
                Dim lstrNroFact As String
                If ldgrSeleccionada.Name = "dgrItemsNotaCon" Then
                    lstrNroFact = ldrvAnticipo("NroFact")
                Else
                    lstrNroFact = ldrvAnticipo("NroFac")
                End If
                Dim lstrPrefFac As String = ClsPanorama.FstrPrefijoDcto(lstrNroFact)
                Dim lentIdFact As Integer = ClsPanorama.FentIdDcto(lstrNroFact)
                SAbraFactura(lstrPrefFac, lentIdFact)
            End If
        End If
    End Sub

    Private Sub Ctl_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles _
            txtIdAnticipo.MouseDoubleClick, txtNroNotaRCr.MouseDoubleClick
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If lelmElemento.Name = "txtIdAnticipo" Then
            SAbraAnticipo(CType(txtIdAnticipo.Content, Integer))
        ElseIf lelmElemento.Name = "txtNroNotaRCr" Then
            If Not String.IsNullOrEmpty(txtNroNotaRCr.Content) Then
                Dim lstrPrefNotaRRC As String = ClsPanorama.FstrPrefijoDcto(txtNroNotaRCr.Content)
                Dim lentIdNotaRRC As Integer = ClsPanorama.FentIdDcto(txtNroNotaRCr.Content)
                SAbraNotaRCr(lstrPrefNotaRRC, lentIdNotaRRC)
            Else
                SLevanteEveNoti("Este Recibo de Caja no ha sido reversado!", "", 0,
                        EnuSeveridadNot.EnuInformacion)
            End If
        End If
    End Sub

    Private Sub OnRatonUp(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TabItem Then
            SEstablezcaDataContext()
        End If
    End Sub
#End Region
End Class
