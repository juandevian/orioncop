Public Class WinNotasReversionCr
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuFecha
        enuCliente
        enuPreAgr
        enuNroDoc
        enuDetalle
        enuDocRev
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsNotaReversionCr = Nothing
    Private MobjCliente As ClsCliente = Nothing
    Private MblnPoblandoCbo As Boolean = False
    Private MstrIdPredioAgr As String = String.Empty
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomNotRevCr
    Private MnuEstadoProvEfac As MenuItem = Nothing
    Private MnuReprocesarEDoc As MenuItem = Nothing
    Private MnuEDoc As MenuItem = Nothing
    Private ReadOnly MwinMW As MWOrionCop = Nothing
#End Region

#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuNotaReverRC
    End Sub
    Public Sub New(awinMW As MWOrionCop)
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuNotaReverRC
        MwinMW = awinMW
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolControlesLlave As New Collection From {
            cboPref,
            txtIdNota
        }
        SAdicioneControlRestringido(dgrMovimiento)
        SAdicioneControlRestringido(bttEncontrarCliente)
        SCargueForma(EnuElementosAdicionalesDef.enuImprimir, 6,
                lcolControlesLlave, dtpFechaNotaRRC, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SVisibiliceCtls()
        SHabiliteMenuEFac()
        SHabiliteReprocesarEDoc()
    End Sub

    Protected Overrides Sub SInicialiceObjeto()
        If IsNothing(ObjObjetoWin) Then
            Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuNotaReversaCr)
            ObjObjetoWin = New ClsNotaReversionCr(lstrPref)
            ObjObjetoWin.SVayaAlUltimo()
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub

    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuFecha) = lblFechaNota
        StcValidaControl(EnuValidEntradaDef.enuCliente) = lblIdCliente
        StcValidaControl(EnuValidEntradaDef.enuPreAgr) = lblPredioAgru
        StcValidaControl(EnuValidEntradaDef.enuNroDoc) = lblNroDoc
        StcValidaControl(EnuValidEntradaDef.enuDocRev) = lblDocRev
        StcValidaControl(EnuValidEntradaDef.enuDetalle) = lblDetalle
        SPuebleComboBoxes()
        '
        HbttAceptar.TabIndex = 20
        HbttCancelar.TabIndex = 21
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

    Protected Overrides Sub SMuestreDatos()
        HblnMostrandoDatos = True
        If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos AndAlso
                EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            SLevanteEveNoti("No hay Notas para ser mostradas!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            cboPref.IsEnabled = False
            txtIdNota.IsEnabled = False
        Else
            With MobjObjetoWin
                cboPref.SelectedItem = .ObjPrefijo_NotaReversaCrStr.ToString()
                txtIdNota.Text = .ObjIdNotaReversaCrEnt.ToString()
                dtpFechaNotaRRC.SelectedDate = .ObjFecha_NotaReversaCrDtm.ObjValorPro
                cboDocRev.SelectedIndex = .ObjTipoDocReversadoByt.ObjValorPro
                txtValorNota.Content = Format(.ObjValor_NotaReversaCrDec.ObjValorPro, "c")
                txtIdCliente.Text = .ObjClienteNota.ObjIdClienteDbl.ObjValorPro
                txtNombreCliente.Content = .ObjClienteNota.ObjNombreCompletoStr.ObjValorPro
                cboPredioAgru.SelectedItem = .ObjIdPredioAgrupador_NotaReversaCrStr.ObjValorPro
                txtPredioAgr.Content = .ObjIdPredioAgrupador_NotaReversaCrStr.ObjValorPro
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    txtNroDoc.Content = .StrNumeroDocRev
                    txtFechaDocRev.Content = Format(.DtmFechaDocRev, GCSTRFMTFECHASIMPLE)
                    txtValorNota.Content = Format(MobjObjetoWin.ObjValor_NotaReversaCrDec.ObjValorPro, "c")
                Else
                    txtNroDoc.Content = .StrNumeroDocRev
                    txtValorNota.Content = Format(.ObjValor_NotaReversaCrDec.ObjValorPro, "c")
                    txtFechaDocRev.Content = Format(.DtmFechaDocRev, GCSTRFMTFECHASIMPLE)
                End If
                txtDetalle.Text = .ObjDetalle_NotaReversaCrStr.ObjValorPro
            End With
        End If
        Title = My.Resources.FichaNRevCr
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            Title &= "Nuevo " & My.Resources.De & txtNombreCliente.Content
        Else
            Title &= MobjObjetoWin.StrNumeroNotaReversaCr & My.Resources.De &
                    txtNombreCliente.Content
        End If
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            SEstablezcaDataContext()
            SMuestreUsuarios()
            SMuestreEstado()
            SHabiliteMenuEFac()
            SHabiliteReprocesarEDoc()
            txtIdNota.Focus()
        End If
        SValide()
        HblnMostrandoDatos = False
    End Sub

    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.cenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntradaDef.enuFecha) = .ObjFecha_NotaReversaCrDtm.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuCliente) = .ObjIdCliente_NotaReversaCrDbl.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuPreAgr) = .ObjIdPredioAgrupador_NotaReversaCrStr.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuDocRev) = .ObjTipoDocReversadoByt.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuNroDoc) = .ObjIdDoc_NotaReversaCrEnt.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuDetalle) = .ObjDetalle_NotaReversaCrStr.BlnEsValido
            End With
        End If
        '
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub

    Protected Overrides Sub SRegistre()
        Dim ldtrPrefRecCaja = String.Empty, lentIdRecCaja = 0
        If Not IsNothing(cboNroDoc.SelectedItem) Then
            ldtrPrefRecCaja = ClsPanorama.FstrPrefijoDcto(cboNroDoc.SelectedItem)
            lentIdRecCaja = ClsPanorama.FentIdDcto(cboNroDoc.SelectedItem)
        End If
        With MobjObjetoWin
            .ObjFecha_NotaReversaCrDtm.ObjValorPro = dtpFechaNotaRRC.SelectedDate
            .ObjIdCliente_NotaReversaCrDbl.ObjValorPro = txtIdCliente.Text
            .ObjIdPredioAgrupador_NotaReversaCrStr.ObjValorPro = MstrIdPredioAgr
            .ObjPrefijo_NotaReversaCrStr.ObjValorPro = ldtrPrefRecCaja
            .ObjIdDoc_NotaReversaCrEnt.ObjValorPro = lentIdRecCaja
            .ObjDetalle_NotaReversaCrStr.ObjValorPro = txtDetalle.Text
        End With
        SValide()
    End Sub

    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Protected Overrides Sub SConfigureMenuesPropios()
        HmnuImprimir = FmnuiMenuItem("MnuImprimir", "Im_primir", "RecMnuItemSec")
        Dim lentPosicion = HmnuAcciones.Items.Count - 2
        Dim lsepSeparad As New Separator
        HmnuAcciones.Items.Insert(lentPosicion, HmnuImprimir)
        HmnuAcciones.Items.Insert(lentPosicion, lsepSeparad)
        If GobjParametros.BlnEFacAutorizado Then
            MnuEDoc = FmnuiMenuItemPan("MnuEDoc", "eDocumento", 0, "", True)
            HmnuMiMenu.Items.Insert(2, MnuEDoc)
            MnuEstadoProvEfac = FmnuiMenuItem("MnuEstadoProvEfac", "Estado en Pro_veedor eFactura",
                    "RecMnuItemSec")
            MnuReprocesarEDoc = FmnuiMenuItem("MnuReprocesarEDoc",
                    "_Reprocesar Documento Electrónico", "RecMnuItemSec")
            MnuEDoc.Items.Add(MnuEstadoProvEfac)
            MnuEDoc.Items.Add(MnuReprocesarEDoc)
        End If
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    Protected Overrides Sub SCree()
        Dim lstrMens = String.Empty
        If ClsOrionCop.FblnPuedeCrear(EnuTipoDocOri.EnuNotaRevCr, False, lstrMens) Then
            MyBase.SCree()
            SVisibiliceCtls()
            SEstablezcaDataContext()
            MobjObjetoWin.ObjIdNotaReversaCrEnt.SValide()
            With GobjParametros
                If .ObjExigeFechaHoyDocsBln.ObjValorPro Then
                    MobjObjetoWin.ObjFecha_NotaReversaCrDtm.ObjValorPro = Date.Today
                    dtpFechaNotaRRC.Style = FindResource("RecCtlNoHabilitado")
                Else
                    If .ObjAnoActual.StrIdPeriodoActual < ClsOrionCop.FstrPeriodoDeFecha(Date.Today) Then
                        MobjObjetoWin.ObjFecha_NotaReversaCrDtm.ObjValorPro =
                                .ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                    End If
                End If
                If .BlnEFacAutorizado Then
                    FblnEstaConectado(GobjParametros.ObjURLStr.ObjValorPro, lstrMens)
                End If
            End With
            If dtpFechaNotaRRC.IsEnabled Then
                dtpFechaNotaRRC.Focus()
            Else
                bttEncontrarCliente.Focus()
            End If
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Protected Overrides Sub SGuarde()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            If MyBase.FblnGravo() Then
                Mouse.OverrideCursor = Cursors.Wait
                lstrMens = "Desea imprimir la Nota?"
                If MsgBox(lstrMens, vbYesNo + MsgBoxStyle.Question, "Imprimir Nota?") = vbYes Then
                    SImprima()
                End If
                If MobjObjetoWin.ObjTipoDocReversadoByt.ObjValorPro =
                        EnuDocReversado.EnuNotaCr Then
                    Dim lobjNCr As ClsNotaCr = MobjObjetoWin.ObjDocReversado
                    If lobjNCr.BlnEsDocEle Then
                        SProceseRevDsctosCrApi()
                    End If
                Else
                    SReverseNsCrDelRC()
                End If
                lstrMens = FstrNombreDoc() & " fue creada exitosamente!"
                SFinaliceOperacion()
                Mouse.OverrideCursor = Cursors.Arrow
            Else
                MobjObjetoWin.SNormaliceEstado(True)
                SHabiliteWin(False)
                SMuestreDatos()
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
                GobjPanDat.SConfirmeTransaccion()
                GobjPanDat.SControleProcesoObj(False)
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            Me.Cursor = Cursors.Arrow
        End Try
    End Sub

    Protected Overrides Sub SFinaliceOperacion()
        MyBase.SFinaliceOperacion()
        SRefrescarClic()
        SVisibiliceCtls()
    End Sub

    Protected Overrides Sub SImprima()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Dim lblnPuede = MobjObjetoWin.BlnEstaRegEFac OrElse Not MobjObjetoWin.BlnEsDocEle
            If lblnPuede Then
                Mouse.OverrideCursor = Cursors.Wait
                If MobjObjetoWin.BlnExiste Then
                    SLevanteEveNoti("Imprimiendo", String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    Dim lentIdNotaPrimera = MobjObjetoWin.ObjIdNotaReversaCrEnt.ObjValorPro
                    Dim lentIdNotaUltima = lentIdNotaPrimera
                    Dim lstrPref = MobjObjetoWin.ObjPrefijo_NotaReversaCrStr.ObjValorPro
                    Dim lobjParaNota As New ClsParametrosReportesDocs(lstrPref,
                            lentIdNotaPrimera, lentIdNotaUltima)
                    Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                        .ObjParRepDocs = lobjParaNota,
                        .EnuReporte = EnuReporteDef.enuNotaReverCr
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
#End Region

#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                Dim largBusqueda As New System.Windows.RoutedEventArgs With {
                    .RoutedEvent = TextBox.LostFocusEvent
                }
                If txtIdCliente.Focus Then
                    txtIdCliente.Text = StrResultadoBusqueda
                    largBusqueda.Source = txtIdCliente
                    OnPierdeFoco(txtIdCliente, largBusqueda)
                End If
            End If
        Else
            If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable Then
                If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable Then
                    If BlnBusquedaOk AndAlso StrResutadosBusqueda.Length > 0 Then
                        cboPref.SelectedItem = StrResutadosBusqueda(0)
                        txtIdNota.Text = StrResutadosBusqueda(1)
                        SAbraNotaRevCr()
                    End If
                End If
            End If
        End If
    End Sub
    ''' <summary>
    ''' Invalida la funcion "fblnDefinioBusqueda" de la clase base.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            If txtIdCliente.Focus Then
                SDefineBusquedaPredioAgr()
                SDefineBusquedaCliente()
            End If
        Else
            SDefineNombreCliente()
            SDefinePredioAgr()
        End If
        Return True
    End Function
    Private Sub SDefineBusquedaPredioAgr()
        Dim lstrTablaPri = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec = ClsPropietario.SstrNombreTabla
        Dim lstrCamSelTablaPri As String() = {"DISTINCT " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd}
        Dim lstrCampSelTablaSec As String() = {ClsIdCliente_PropDbl.SstrNombreCampoBd,
                ClsNombreCompleto_PropStr.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta,
                StrCampoCentroutil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta,
                StrCampoCentroutil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdCliente_PropDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & StrCampoCarpeta &
                " = " & GshrIdCarpeta & " AND P." &
                StrCampoCentroutil & " = " &
                GshrIdCentroUtil & " AND " & "P." & ClsIdPredioStr.SstrNombreCampoBd &
                " = " & ClsIdPredioAgrupadorStr.SstrNombreCampoBd & " AND " &
                lstrCampoBusqueda & " <> ''"
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador - Propietario", lstrTablaPri,
                lstrTablaSec, lstrCamSelTablaPri, lstrCampSelTablaSec, lstrCampRelPri,
                lstrCampRelSec, lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, False)
    End Sub
    Private Sub SDefineBusquedaCliente()
        Dim lstrTabla = ClsCliente.SstrNombreTabla
        Dim lstrCamposMostrar = {ClsIdClienteDbl.SstrNombreCampoBd,
                ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdClienteDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " & lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
    Private Sub SDefineNombreCliente()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsNotaReversionCr.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd, ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgr_NotaReversaCrStr.SstrNombreCampoBd,
                ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd, ClsIdNotaReversaCrEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_NotaReversaCrDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCamposRetornar As String() = {ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd,
                                              ClsIdNotaReversaCrEnt.SstrNombreCampoBd}
        Dim lstrFiltro As String = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCamposRetornar, lstrFiltro, False)
    End Sub
    Private Sub SDefinePredioAgr()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsNotaReversionCr.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgr_NotaReversaCrStr.SstrNombreCampoBd,
                                            ClsFecha_NotaReversaCrDtm.SstrNombreCampoBd,
                                            ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd,
                                            ClsIdNotaReversaCrEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_NotaReversaCrDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsIdPredioAgr_NotaReversaCrStr.SstrNombreCampoBd
        Dim lstrCamposRetornar As String() = {ClsPrefijo_NotaReversaCrStr.SstrNombreCampoBd,
                                            ClsIdNotaReversaCrEnt.SstrNombreCampoBd}
        Dim lstrFiltro = "S." & StrCampoCarpeta & " = " &
                GshrIdCarpeta.ToString & " AND S." & StrCampoCentroUtil &
                " = " & GshrIdCentroUtil.ToString
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCamposRetornar, lstrFiltro, False)
    End Sub
    Private Sub SBuscarCliente()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            StrResultadoBusqueda = String.Empty
            SBuscar()
            If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                txtIdCliente.Text = StrResultadoBusqueda
            End If
        End If
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SInicialiceNota()
        If MobjObjetoWin IsNot Nothing Then
            If cboPref.SelectedItem Is Nothing Then
                cboPref.SelectedIndex = 0
            End If
            Dim lstrPref As String = cboPref.SelectedItem
            If MobjObjetoWin.ObjPrefijo_NotaReversaCrStr.ToString <> lstrPref Then
                ObjObjetoWin = New ClsNotaReversionCr(lstrPref)
                MobjObjetoWin = ObjObjetoWin
                If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                    MobjObjetoWin.SVayaAlUltimo()
                End If
            End If
        End If
    End Sub
    Private Sub SPuebleComboBoxes()
        MblnPoblandoCbo = True
        Dim ldrwConst = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuDocReversado)
        SPuebleComboBox(ldrwConst, cboDocRev)
        ldrwConst = ClsOrionCop.FdrwPrefDoc(EnuTipoDocOri.EnuNotaRevCr)
        SPuebleComboBox(ldrwConst, cboPref)
        MblnPoblandoCbo = False
    End Sub
    Private Sub SPuebleCboPredAgru()
        MblnPoblandoCbo = True
        cboPredioAgru.Items.Clear()
        cboPredioAgru.Items.Add(My.Resources.Ninguno)
        cboPredioAgru.Items.Add(GCSTRSINPA)
        Dim lstrPredAgrs As String() =
                        MobjObjetoWin.ObjClienteNota.FstrPrediosAgruClienteConFacturas(False)
        For Each lstrPreAgr As String In lstrPredAgrs
            If Not String.IsNullOrEmpty(lstrPreAgr) Then
                cboPredioAgru.Items.Add(lstrPreAgr)
            End If
        Next
        MblnPoblandoCbo = False
        cboPredioAgru.SelectedIndex = 0
    End Sub
    Private Sub SPuebleComboDocs()
        With MobjObjetoWin
            MblnPoblandoCbo = True
            cboNroDoc.Items.Clear()
            cboNroDoc.Items.Add(My.Resources.Ninguno)
            If .ObjTipoDocReversadoByt.BlnEsValido AndAlso .ObjIdCliente_NotaReversaCrDbl.BlnEsValido AndAlso
                    .ObjIdPredioAgrupador_NotaReversaCrStr.BlnEsValido Then
                Dim lstrDocs As String() = Array.Empty(Of String)()
                If .ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuReciboC Then
                    lstrDocs = FstrNrosRecCaja()
                ElseIf .ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuNotaCr Then
                    lstrDocs = FstrNrosNotasCr()
                End If
                If lstrDocs.Length > 0 Then
                    For Each lstrDoc As String In lstrDocs
                        cboNroDoc.Items.Add(lstrDoc)
                    Next
                Else
                    Dim lstrNomDoc = "Recibos de Caja"
                    If .ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuNotaCr Then
                        lstrNomDoc = "Notas Crédito"
                    End If
                    Dim lstrMens = "No hay " & lstrNomDoc & " para este Cliente y Predio Agrupador"
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            End If
            cboNroDoc.SelectedIndex = 0
            MblnPoblandoCbo = False
        End With
    End Sub
    Private Function FstrNrosRecCaja() As String()
        Dim lstrRecCaja() As String = Array.Empty(Of String)(), j = -1
        Dim ldtmFechaIni = DateSerial(1900, 1, 1)
        Dim ldtmFechaFin = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo.AddDays(-1)
        Dim ldtbRecibos = MobjCliente.FdtbRecibos(MstrIdPredioAgr, ldtmFechaIni,
                ldtmFechaFin)
        If ldtbRecibos.Rows.Count > 0 Then
            ReDim lstrRecCaja(ldtbRecibos.Rows.Count - 1)
            Dim lstrPrefRc As String, lentIdRecCaja As Integer
            Dim ldrwRecCaja As DataRow
            For i As Integer = ldtbRecibos.Rows.Count - 1 To 0 Step -1
                ldrwRecCaja = ldtbRecibos.Rows(i)
                If Not ClsPanorama.FobjValorCampo(ldrwRecCaja(
                        ClsAnuladoBln.SstrNombreCampoBd), EnuTipoValor.enuBoolean) Then
                    j += 1
                    lstrPrefRc = ldrwRecCaja(ClsPrefijo_RecStr.SstrNombreCampoBd)
                    lentIdRecCaja = ldrwRecCaja(ClsIdRecCajaEnt.SstrNombreCampoBd)
                    lstrRecCaja(j) = ClsPanorama.FstrNumeroDcto(lstrPrefRc, lentIdRecCaja)
                End If
            Next
        End If
        Return lstrRecCaja
    End Function
    Private Function FstrNrosNotasCr() As String()
        Dim lstrNotasCr() As String = Array.Empty(Of String)(), j = -1
        Dim ldtmFechaIni = DateSerial(1900, 1, 1)
        Dim ldtmFechaFin = Date.Today
        Dim ldtbNotasCr = MobjCliente.FdtbNotasCr(MstrIdPredioAgr, ldtmFechaIni, ldtmFechaFin)
        If ldtbNotasCr.Rows.Count > 0 Then
            Dim lstrPrefNcr As String, lentIdNotaCr As Integer
            Dim ldrwNotaCr As DataRow
            Dim lobjNotaCr As New ClsNotaCr()
            Dim lobjValorLlave As Object()
            For i As Integer = ldtbNotasCr.Rows.Count - 1 To 0 Step -1
                ldrwNotaCr = ldtbNotasCr.Rows(i)
                If Not ClsPanorama.FobjValorCampo(ldrwNotaCr(ClsAnuladoBln.SstrNombreCampoBd),
                        EnuTipoValor.enuBoolean) Then
                    lstrPrefNcr = ldrwNotaCr(ClsPrefijo_NotaCrStr.SstrNombreCampoBd)
                    lentIdNotaCr = ldrwNotaCr(ClsIdNotaCrEnt.SstrNombreCampoBd)
                    lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefNcr, lentIdNotaCr}
                    lobjNotaCr.SAbra(lobjValorLlave)
                    If lobjNotaCr.ObjIdTipoNotaCrByt.ObjValorPro <> EnuTipoNotaCrDef.EnuAnulaFac Then
                        j += 1
                        ReDim Preserve lstrNotasCr(j)
                        lstrNotasCr(j) = ClsPanorama.FstrNumeroDcto(lstrPrefNcr, lentIdNotaCr)
                    End If
                End If
            Next
        End If
        Return lstrNotasCr
    End Function
    Private Sub SAbraNotaRevCr()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                If txtIdNota.Text <> MobjObjetoWin.ObjIdNotaReversaCrEnt.ToString() Then
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                            cboPref.SelectedItem, txtIdNota.Text}
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
    Private Sub SEstablezcaDataContext()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            dgrMovimiento.DataContext = MobjObjetoWin.DtbNovedadesIU
        Else
            dgrMovimiento.DataContext = Nothing
        End If
    End Sub
    Private Sub SVisibiliceCtls()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            txtNroDoc.Visibility = Visibility.Visible
            cboNroDoc.Visibility = Visibility.Hidden
            txtPredioAgr.Visibility = Visibility.Visible
            cboPredioAgru.Visibility = Visibility.Hidden
            bttEncontrarCliente.Visibility = Visibility.Hidden
            txtIdNota.IsEnabled = True
            cboPref.IsEnabled = True
            txtIdNota.Focus()
        Else
            txtNroDoc.Visibility = Visibility.Hidden
            cboNroDoc.Visibility = Visibility.Visible
            txtPredioAgr.Visibility = Visibility.Hidden
            cboPredioAgru.Visibility = Visibility.Visible
            bttEncontrarCliente.Visibility = Visibility.Visible
            txtIdNota.IsEnabled = False
            cboPref.IsEnabled = False
            dtpFechaNotaRRC.Focus()
        End If
    End Sub
    Private Sub SMuestreUsuarios()
        With MobjObjetoWin
            If MobjObjetoWin.BlnExiste Then
                txtUsuarioGenero.Content = MobjObjetoWin.ObjIdUsuario_NotaReversaCrStr.ObjValorPro
                If .ObjAnuladoBln.ObjValorPro Then
                    lblUsuarioAnulo.Visibility = Visibility.Visible
                    txtUsuarioAnulo.Visibility = Visibility.Visible
                    txtUsuarioAnulo.Content = .ObjIdUsuarioAnuloStr.ObjValorPro
                Else
                    lblUsuarioAnulo.Visibility = Visibility.Collapsed
                    txtUsuarioAnulo.Visibility = Visibility.Collapsed
                    txtUsuarioAnulo.Content = String.Empty
                End If
            End If
        End With
    End Sub
    Private Sub SMuestreEstado()
        If MobjObjetoWin.ObjAnuladoBln.ObjValorPro Then
            txtEstado.Style = FindResource("RecDocAnulado")
        Else
            txtEstado.Style = FindResource("RecDocNormal")
        End If
    End Sub
    Private Sub SReverseNsCrDelRC()
        Dim lobjRecCaja As ClsReciboCaja = MobjObjetoWin.ObjDocReversado
        Dim lstrMens = String.Empty
        If Not String.IsNullOrEmpty(lobjRecCaja.ObjIdNotasCrStr.ToString) Then
            lobjRecCaja.SReverseNotasCr(MobjObjetoWin.ObjFecha_NotaReversaCrDtm.ObjValorPro)
            If GobjParametros.BlnEFacAutorizado AndAlso MobjObjetoWin.BlnAfectaFrasRegEFac Then
                SProceseEFac(lstrMens)
            End If
            If String.IsNullOrEmpty(lstrMens) Then
                lstrMens = "Fueron Reversadas las correspondientes Notas Crédito!"
            End If
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
#End Region

#Region "Efac"
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
                    Else
                    End If
                End If
            Else
                MnuEDoc.Visibility = Visibility.Collapsed
            End If
        End If
    End Sub

    Private Sub SProceseRevDsctosCrApi()
        If MobjObjetoWin.BlnAfectaFrasRegEFac Then
            Dim lstrMens = String.Empty
            SProceseEFac(lstrMens)
            If Not String.IsNullOrEmpty(lstrMens) Then
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            End If
        End If
    End Sub

    Private Sub SHabiliteReprocesarEDoc()
        If MnuReprocesarEDoc IsNot Nothing Then
            MnuReprocesarEDoc.Visibility = Visibility.Collapsed
        End If
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If Not HblnCargandoForma AndAlso MobjObjetoWin.BlnEsDocEle AndAlso
                    MobjObjetoWin.BlnAfectaFrasRegEFac Then
                MnuReprocesarEDoc.Visibility = Visibility.Visible
            End If
        End If
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            If lelmElemento.Name = "bttEncontrarCliente" Then
                SBuscarCliente()
            End If
        End If

    End Sub

    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                Dim lmnuOpcion As MenuItem = lelmElemento
                If lmnuOpcion.Name = "MnuEstadoProvEfac" Then
                    Mouse.OverrideCursor = Cursors.Wait
                    SRefresqueWin()
                    Dim lstrIdNotaCon As String
                    If String.IsNullOrEmpty(MobjObjetoWin.ObjPrefijo_NotaReversaCrStr.ToString()) Then
                        lstrIdNotaCon = "RCR" & MobjObjetoWin.StrIdObjeto
                    Else
                        lstrIdNotaCon = MobjObjetoWin.StrIdObjeto
                    End If
                    SMuestreEstadoEFac(MobjObjetoWin.ObjCUDocStr.ObjValorPro,
                            EnuTipoDocOri.EnuNotaRevCr, MobjObjetoWin.StrIdObjeto, False,
                            MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro, lstrMens)
                    Mouse.OverrideCursor = Cursors.Arrow
                ElseIf lmnuOpcion.Name = "MnuRegistrarEFac" Then
                    SProceseRevDsctosCrApi()
                ElseIf lmnuOpcion.Name = "MnuReprocesarEDoc" Then
                    If MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                            EnuEstadoEDoc.EnuNoEDoc AndAlso
                            MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                            EnuEstadoEDoc.EnuNoReg AndAlso
                            MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                             EnuEstadoEDoc.EnuInvalida Then
                        If MobjObjetoWin.ObjFecha_NotaReversaCrDtm.ObjValorPro >
                                Today.AddDays(-30) Then
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

    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando AndAlso Not HblnMostrandoDatos Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is DatePicker Then
                Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                Try
                    GobjPanDat.SControleProcesoObj(True)
                    If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
                        With MobjObjetoWin
                            Select Case lelmElemento.Name
                                Case "dtpFechaNotaRRC"
                                    .ObjFecha_NotaReversaCrDtm.ObjValorPro = dtpFechaNotaRRC.SelectedDate
                                Case "txtIdCliente"
                                    .ObjIdCliente_NotaReversaCrDbl.ObjValorPro = txtIdCliente.Text
                                    MobjCliente = .ObjClienteNota
                                    SPuebleCboPredAgru()
                                Case "txtDetalle"
                                    MobjObjetoWin.ObjDetalle_NotaReversaCrStr.ObjValorPro = txtDetalle.Text
                            End Select
                        End With
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
        End If
    End Sub

    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If Not MblnPoblandoCbo AndAlso TypeOf lelmElemento Is ComboBox AndAlso
                Not HblnSeEstaCerrando Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando AndAlso
                    Not HblnMostrandoDatos Then
                Dim lstrMens = String.Empty
                With MobjObjetoWin
                    Select Case lelmElemento.Name
                        Case "cboDocRev"
                            If cboDocRev.SelectedIndex <> .ObjTipoDocReversadoByt.ObjValorPro Then
                                .ObjTipoDocReversadoByt.ObjValorPro = cboDocRev.SelectedIndex
                                SPuebleComboDocs()
                            End If
                        Case "cboPredioAgru"
                            MstrIdPredioAgr = If(cboPredioAgru.SelectedItem = GCSTRSINPA,
                                String.Empty, cboPredioAgru.SelectedItem)
                            If cboPredioAgru.SelectedItem <>
                                    .ObjIdPredioAgrupador_NotaReversaCrStr.ObjValorPro Then
                                .ObjIdPredioAgrupador_NotaReversaCrStr.ObjValorPro = MstrIdPredioAgr
                                SPuebleComboDocs()
                            End If
                        Case "cboNroDoc"
                            If cboNroDoc.SelectedItem IsNot Nothing AndAlso
                                    cboNroDoc.SelectedItem <> My.Resources.Ninguno Then
                                Dim lstrPrefDoc = ClsPanorama.FstrPrefijoDcto(cboNroDoc.SelectedItem)
                                Dim lentIdDoc = ClsPanorama.FentIdDcto(cboNroDoc.SelectedItem)
                                If Not ClsOrionCop.FblnDocReversable(lstrPrefDoc, lentIdDoc,
                                        .ObjTipoDocReversadoByt.ObjValorPro, lstrMens) Then
                                    lstrPrefDoc = String.Empty
                                    lentIdDoc = 0
                                End If
                                .ObjPrefijoDoc_NotaReversaCrStr.ObjValorPro = lstrPrefDoc
                                .ObjIdDoc_NotaReversaCrEnt.ObjValorPro = lentIdDoc
                            Else
                                .ObjPrefijo_NotaReversaCrStr.ObjValorPro = String.Empty
                                .ObjIdDoc_NotaReversaCrEnt.ObjValorPro = 0
                            End If
                    End Select
                End With
                SMuestreDatos()
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            ElseIf EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                If lelmElemento.Name = "cboPref" Then
                    SInicialiceNota()
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub

    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdNota.KeyDown
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If e.Key = Key.Return OrElse e.Key = Key.Tab Then
                SAbraNotaRevCr()
            End If
        End If
    End Sub

    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
                dgrMovimiento.MouseRightButtonUp
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Dim lstrPrefRc As String
        Dim lentIdRC As Integer
        If lelmElemento.Name = "dgrMovimiento" Then
            Dim ldrvFilaActual As DataRowView
            Dim ldgrActual As DataGrid = sender
            ldrvFilaActual = ldgrActual.SelectedItem
            Dim lstrNroFac As String = ldrvFilaActual("NroFac")
            If lstrNroFac <> "0" Then
                Dim lstrPreFac = ClsPanorama.FstrPrefijoDcto(lstrNroFac)
                Dim lentIdFac = ClsPanorama.FentIdDcto(lstrNroFac)
                SAbraFactura(lstrPreFac, lentIdFac)
            Else
                lstrPrefRc = ClsPanorama.FstrPrefijoDcto(txtNroDoc.Content)
                lentIdRC = ClsPanorama.FentIdDcto(txtNroDoc.Content)
                SAbraRecibo(lstrPrefRc, lentIdRC)
            End If
        End If
    End Sub

    Private Sub TxtCtl_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles _
                txtNroDoc.MouseDoubleClick
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Dim lstrPrefDoc = ClsPanorama.FstrPrefijoDcto(txtNroDoc.Content)
        Dim lentIdDoc = ClsPanorama.FentIdDcto(txtNroDoc.Content)
        If lelmElemento.Name = "txtNroDoc" Then
            If MobjObjetoWin.ObjTipoDocReversadoByt.ObjValorPro = EnuDocReversado.EnuReciboC Then
                SAbraRecibo(lstrPrefDoc, lentIdDoc)
            Else
                SAbraNotaCr(lstrPrefDoc, lentIdDoc)
            End If
        End If
    End Sub
#End Region
End Class