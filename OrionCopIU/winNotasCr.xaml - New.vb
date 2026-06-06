Imports System.ComponentModel
Public Class WinNotasCr
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuEstadoItem As Integer
        None = 0
        enuConsultandoItem
        enuCreandoItem
    End Enum
    Private Enum EnuValidEntradaDef As Integer
        enuFecha
        enuCliente
        enuPreAgr
        enuComent
        enuNroFact
        enuItemFra
        enuTipoDscto
        enuValorDscto
        enuItems
        enuCausarInt
        enuModoNota
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsNotaCr = Nothing
    Private MobjItemNotaCrActual As ClsItemNotaCr = Nothing
    Private MobjCliente As ClsCliente = Nothing
    Private ReadOnly MobjFactura As New ClsFactura()
    Private MobjItemFac As ClsItemFactura = Nothing
    Private MenuEstadoItem As EnuEstadoItem = EnuEstadoItem.None
    Private MblnDejoUltimoControl As Boolean = False
    Private MblnPoblandoCbo As Boolean = False
    Private MdtbItems As DataTable = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomNotCr
    Private MblnEliminandoItem As Boolean = False
    Private MdecBaseRet As Decimal = 0
    Private MdblTasaRet As Double = 0
    Private MnuEnviarPorCorreo As MenuItem = Nothing
    Private MnuEstadoProvEfac As MenuItem = Nothing
    Private MnuReprocesarEDoc As MenuItem = Nothing
    Private ReadOnly MwinMW As MWOrionCop = Nothing
    Private MblnCancelaCrear As Boolean = False
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuNotaCr
    End Sub
    Public Sub New(awinMW As MWOrionCop)
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuNotaCr
        MwinMW = awinMW
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Mouse.OverrideCursor = Cursors.Wait
        Dim lcolControlesLlave As New Collection From {
            cboPref,
            txtIdNota
        }
        SAdicioneCtlsRestringidos()
        SCargueForma(EnuElementosAdicionalesDef.enuImprimir, 11,
                lcolControlesLlave, txtIdCliente, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SCambieNombresDocCobro()
        SHabiliteMenuEFac()
        SHabiliteReprocesarEDoc()
        txtIdNota.Focus()
        Mouse.OverrideCursor = Cursors.Arrow
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
            Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuNotaCr)
            ObjObjetoWin = New ClsNotaCr(lstrPref)
            If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                ObjObjetoWin.SVayaAlUltimo()
            End If
        End If
        MobjObjetoWin = ObjObjetoWin
        If Not IsNothing(MobjObjetoWin) Then
            If IsNothing(MdtbItems) Then
                MdtbItems = FdtbItems()
            End If
        End If
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuFecha) = lblFechaNotaCr
        StcValidaControl(EnuValidEntradaDef.enuCliente) = lblIdCliente
        StcValidaControl(EnuValidEntradaDef.enuPreAgr) = lblPredioAgru
        StcValidaControl(EnuValidEntradaDef.enuComent) = lblComentario
        StcValidaControl(EnuValidEntradaDef.enuNroFact) = lblIdFactura
        StcValidaControl(EnuValidEntradaDef.enuItemFra) = lblItemFactura
        StcValidaControl(EnuValidEntradaDef.enuTipoDscto) = lblTipoDcto
        StcValidaControl(EnuValidEntradaDef.enuValorDscto) = lblValorDct
        StcValidaControl(EnuValidEntradaDef.enuCausarInt) = lblIntPorCausar
        StcValidaControl(EnuValidEntradaDef.enuItems) = lblDescuentos
        StcValidaControl(EnuValidEntradaDef.enuModoNota) = lblModoNota
        SPuebleComboBoxes()
        SDeshabiliteControlesActuales()
        txtValorNotaCr.Content = Format(0, "c")
        '
        HbttAceptar.TabIndex = 100
        HbttCancelar.TabIndex = 101
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos AndAlso
                EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            SLevanteEveNoti("No hay Notas para ser mostradas!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            cboPref.IsEnabled = False
            txtIdNota.IsEnabled = False
        Else
            HblnMostrandoDatos = True
            With MobjObjetoWin
                cboPref.SelectedItem = .ObjPrefijo_NotaCrStr.ToString()
                txtIdNota.Text = .ObjIdNotaCrEnt.ToString()
                dtpFechaNotaCr.Text = .ObjFecha_NotaCrDtm.ToString
                txtIdCliente.Text = .ObjIdCliente_NotaCrDbl.ToString
                txtNombreCliente.Content = .ObjClienteNotaCr.ObjNombreCompletoStr.ToString
                txtNombreCliente.ToolTip = txtNombreCliente.Content
                If String.IsNullOrEmpty(.ObjIdPredioAgrupador_NotaCrStr.ToString) Then
                    txtPredioAgr.Content = GCSTRSINPA
                Else
                    If .ObjIdPredioAgrupador_NotaCrStr.ToString = "***" Then
                        txtPredioAgr.Content = My.Resources.Ninguno
                    Else
                        txtPredioAgr.Content = .ObjIdPredioAgrupador_NotaCrStr.ToString
                    End If
                End If
                cboModoNota.SelectedIndex = MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro
                If .ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuAnulaFac Then
                    lblFacAnu.Visibility = Visibility.Visible
                End If
                txtValorNotaCr.Content = Format(.ObjValor_NotaCrDec.ObjValorPro, "c")
                txtComentario.Text = .ObjComentario_NotaCrStr.ObjValorPro
            End With
            If MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro = EnuModoNotaCr.EnuPorFactura AndAlso
                    IsNothing(MobjItemNotaCrActual) Then
                SVacieItem()
            End If
            SEstablezcaDataContext()
            SMuestreUsuarios()
            SMuestreEstado()
            SHabiliteMenuEFac()
            SHabiliteReprocesarEDoc()
            Title = My.Resources.FichaNCr
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
                Title &= "Nuevo " & My.Resources.De & txtNombreCliente.Content
            Else
                Title &= MobjObjetoWin.StrNumeroNotaCr & My.Resources.De &
                txtNombreCliente.Content
                txtIdNota.Focus()
            End If
            SValide()
            HblnMostrandoDatos = False
        End If
        Title = My.Resources.FichaNCr
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If lblnNoHayDatos Then
                SInicialiceValido()
            End If
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntradaDef.enuCliente) =
                    .ObjIdCliente_NotaCrDbl.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuComent) = .ObjComentario_NotaCrStr.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuFecha) = .ObjFecha_NotaCrDtm.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuPreAgr) =
                    .ObjIdPredioAgrupador_NotaCrStr.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuModoNota) = .ObjModoNotaCrByt.BlnEsValido
            End With
            With MobjItemNotaCrActual
                If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
                    If Not IsNothing(MobjItemNotaCrActual) Then
                        StcValidValido(EnuValidEntradaDef.enuNroFact) =
                                .ObjPrefijoFact_ItemNotaCrStr.BlnEsValido AndAlso
                                .ObjIdFactura_ItemNotaCrEnt.BlnEsValido
                        StcValidValido(EnuValidEntradaDef.enuItemFra) =
                                .ObjIdItemFac_ItemNotaCrShr.BlnEsValido
                        StcValidValido(EnuValidEntradaDef.enuTipoDscto) =
                                .ObjIdTipoDscto_ItemNotaCrByt.BlnEsValido
                        StcValidValido(EnuValidEntradaDef.enuValorDscto) =
                                .ObjValor_ItemNotaCrDec.BlnEsValido()
                    Else
                        If MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro =
                                EnuModoNotaCr.EnuPorValor Then
                            StcValidValido(EnuValidEntradaDef.enuNroFact) = True
                            StcValidValido(EnuValidEntradaDef.enuItemFra) = True
                            StcValidValido(EnuValidEntradaDef.enuTipoDscto) = MobjObjetoWin.
                                    FblnTipoDsctoValido()
                            StcValidValido(EnuValidEntradaDef.enuValorDscto) = MobjObjetoWin.
                                    FblnValorDsctoValido()
                        Else
                            StcValidValido(EnuValidEntradaDef.enuNroFact) = False
                            StcValidValido(EnuValidEntradaDef.enuItemFra) = False
                            StcValidValido(EnuValidEntradaDef.enuTipoDscto) = False
                            StcValidValido(EnuValidEntradaDef.enuValorDscto) = False
                        End If
                    End If
                    StcValidValido(EnuValidEntradaDef.enuItems) = dgrDescuentos.Items.Count > 0
                Else
                    StcValidValido(EnuValidEntradaDef.enuNroFact) = True
                    StcValidValido(EnuValidEntradaDef.enuNroFact) = True
                    StcValidValido(EnuValidEntradaDef.enuItemFra) = True
                    StcValidValido(EnuValidEntradaDef.enuTipoDscto) = True
                    StcValidValido(EnuValidEntradaDef.enuValorDscto) = True
                    StcValidValido(EnuValidEntradaDef.enuItems) = True
                    StcValidValido(EnuValidEntradaDef.enuCausarInt) = True
                End If
            End With
            SHabiliteBttDscto()
        End If
        '
        SHabiliteBotonesTlb()
        If FblnEstanTodosBien() Then
            If MenuEstadoItem = EnuEstadoItem.enuCreandoItem Then
                HbttAceptar.Style = FindResource("RecBttAceDesha")
            End If
        End If
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjComentario_NotaCrStr.ObjValorPro = txtComentario.Text
            .ObjIdNotaCrEnt.ObjValorPro = txtIdNota.Text
            .ObjFecha_NotaCrDtm.ObjValorPro = dtpFechaNotaCr.SelectedDate
            .ObjIdCliente_NotaCrDbl.ObjValorPro = txtIdCliente.Text
            If Not IsNothing(cboPredioAgru.SelectedItem) AndAlso
                cboPredioAgru.SelectedItem <> My.Resources.Ninguno Then
                If cboPredioAgru.SelectedItem = GCSTRSINPA Then
                    .ObjIdPredioAgrupador_NotaCrStr.ObjValorPro = ""
                Else
                    .ObjIdPredioAgrupador_NotaCrStr.ObjValorPro = cboPredioAgru.SelectedItem
                End If
            End If
        End With
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
        MnuEnviarPorCorreo = FmnuiMenuItemPan("MnuEnviarPorCorreo", "_Enviar por eMail", 1, "")
        If ClsPanorama.FblnEmailsHabilitado Then
            HmnuAcciones.Items.Insert(lentPosicion, MnuEnviarPorCorreo)
        End If
        HmnuAcciones.Items.Insert(lentPosicion, HmnuImprimir)
        HmnuAcciones.Items.Insert(lentPosicion, lsepSeparad)
        If GobjParametros.BlnEFacAutorizado Then
            Dim lsepSeparadEFac As New Separator
            MnuEstadoProvEfac = FmnuiMenuItem("MnuEstadoProvEfac", "Estado en Pro_veedor eFactura",
                "RecMnuItemSec")
            MnuReprocesarEDoc = FmnuiMenuItem("MnuReprocesarEDoc",
                    "_Reprocesar Nota Electrónica", "RecMnuItemSec")
            lentPosicion = HmnuHerramientas.Items.Count
            HmnuHerramientas.Items.Insert(lentPosicion, MnuReprocesarEDoc)
            HmnuHerramientas.Items.Insert(lentPosicion, MnuEstadoProvEfac)
            HmnuHerramientas.Items.Insert(lentPosicion, lsepSeparadEFac)
        End If
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SHabiliteMenues()
        MyBase.SHabiliteMenues()
    End Sub
    Protected Overrides Sub SCree()
        Dim lstrMens = String.Empty
        If ClsOrionCop.FblnPuedeCrear(EnuTipoDocOri.EnuNotaCr, False, lstrMens) Then
            SVisibiliceControlesNuevos(True)
            SVisibiliceBttDscto(False)
            SVacieItem()
            SHabiliteCtlsItems()
            MyBase.SCree()
            MblnCancelaCrear = False
            MobjObjetoWin.ObjIdTipoNotaCrByt.ObjValorPro = EnuTipoNotaCrDef.EnuDescuento
            With GobjParametros
                If .ObjExigeFechaHoyDocsBln.ObjValorPro Then
                    MobjObjetoWin.ObjFecha_NotaCrDtm.ObjValorPro = Date.Today
                    dtpFechaNotaCr.Style = FindResource("RecCtlNoHabilitado")
                Else
                    If .ObjAnoActual.StrIdPeriodoActual < ClsOrionCop.FstrPeriodoDeFecha(Date.Today) Then
                        MobjObjetoWin.ObjFecha_NotaCrDtm.ObjValorPro =
                            .ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                    End If
                End If
                dtpFechaNotaCr.SelectedDate = MobjObjetoWin.ObjFecha_NotaCrDtm.ObjValorPro
            End With
            MdtbItems.Rows.Clear()
            grdDetalles.DataContext = MdtbItems
            tbiNovedades.DataContext = Nothing
            If GobjParametros.BlnEFacAutorizado Then
                FblnEstaConectado(GobjParametros.ObjURLStr.ObjValorPro, lstrMens)
            End If
        End If
        If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
            bttEncontrarCliente.Focus()
        Else
            dtpFechaNotaCr.Focus()
        End If
        SValide()
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Protected Overrides Sub SGuarde()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                SRegistre()
                SValide()
                Dim lblnGuardo = FblnGravo()
                If lblnGuardo Then
                    SFinaliceOperacion()
                    lstrMens = "Desea imprimir la Nota?"
                    If MsgBox(lstrMens, vbYesNo + MsgBoxStyle.Question, "Imprimir Nota?") = vbYes Then
                        Mouse.OverrideCursor = Cursors.Wait
                        SImprima()
                        Mouse.OverrideCursor = Cursors.Arrow
                    End If
                    lstrMens = FstrNombreDoc()
                    If lstrMens.StartsWith("El") Then
                        lstrMens &= " fue creado exitosamente!"
                    Else
                        lstrMens &= " fue creada exitosamente!"
                    End If
                    SCrearClic()
                End If
            End If
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentNullException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
                SFinaliceOperacion()
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Protected Overrides Sub SCancele()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            MblnCancelaCrear = True
            SFinaliceOperacion()
            SRefrescarClic()
        Else
            SCerrarClic()
        End If
    End Sub
    Protected Overrides Function SAnule() As Boolean
        Dim lblnAnulo = MyBase.SAnule()
        If lblnAnulo Then
            SInserteDbAnulaCr()
        End If
        Return lblnAnulo
    End Function
    Protected Overrides Sub SEstablezcaWinConsultando()
        If Not MblnCancelaCrear AndAlso Not HblnSeEstaCerrando Then
            If GobjParametros.BlnEFacAutorizado AndAlso MobjObjetoWin.BlnAfectaFrasRegEFac Then
                If MobjObjetoWin.BlnTieneDsctos Then
                    SProceseNotasCrApi()
                End If
            End If
        End If
        MblnPoblandoCbo = True
        cboNroFactNuevo.Items.Clear()
        cboItemFactNuevo.Items.Clear()
        cboPredioAgru.Items.Clear()
        cboTipoDsctoNuevo.SelectedIndex = 0
        MblnPoblandoCbo = False
        txtValorDctoNuevo.Text = Format(0, "c")
        MobjItemNotaCrActual = Nothing
        SVisibiliceControlesNuevos(False)
        MyBase.SEstablezcaWinConsultando()
        SVisibiliceCtrlMora(False)
        txtIdNota.SelectAll()
    End Sub
    Protected Overrides Sub SImprima()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Dim lblnPuede = MobjObjetoWin.BlnEstaRegEFac OrElse Not MobjObjetoWin.BlnEsDocEle
            If lblnPuede Then
                Mouse.OverrideCursor = Cursors.Wait
                If MobjObjetoWin.BlnExiste Then
                    SLevanteEveNoti("Imprimiendo", String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    Dim lstrPrefNotaCr = MobjObjetoWin.ObjPrefijo_NotaCrStr.ObjValorPro
                    Dim lentIdNotaPrimera = MobjObjetoWin.ObjIdNotaCrEnt.ObjValorPro
                    Dim lentIdNotaUltima = MobjObjetoWin.ObjIdNotaCrEnt.ObjValorPro
                    Dim lobjParaNotaCr As New ClsParametrosReportesDocs(lstrPrefNotaCr,
                            lentIdNotaPrimera, lentIdNotaUltima)
                    Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                    .ObjParRepDocs = lobjParaNotaCr,
                    .EnuReporte = EnuReporteDef.enuNotaCr
                    }
                    lobjRep.SGenereReporte()
                    SLevanteEveNoti("", String.Empty, 0, EnuSeveridadNot.EnuOk)
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
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
            If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                If txtIdCliente.Focus Then
                    Dim lstrMens = String.Empty
                    txtIdCliente.Text = StrResultadoBusqueda
                    SRegistreCliente(lstrMens)
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                End If
            End If
        Else
            If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.EnuNavegable Then
                If BlnBusquedaOk AndAlso StrResutadosBusqueda.Length > 0 Then
                    cboPref.SelectedItem = StrResutadosBusqueda(0)
                    txtIdNota.Text = StrResutadosBusqueda(1)
                    SAbraNotaCr()
                End If
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
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
            If txtIdCliente.Focus Then
                SDefineBusquedaPredioAgr_Prop()
                SDefineBusquedaPredioAgr_Arren()
                SDefineBusquedaCliente()
            End If
        Else
            SDefineNombreCliente()
            SDefinePredioAgr()
        End If
        Return True
    End Function
    Private Sub SDefineBusquedaPredioAgr_Prop()
        Dim lstrTablaPri = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec = ClsPropietario.SstrNombreTabla
        Dim lstrCamSelTablaPri As String() = {"DISTINCT " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd}
        Dim lstrCampSelTablaSec As String() = {ClsIdCliente_PropDbl.SstrNombreCampoBd,
                ClsNombreCompleto_PropStr.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta,
                StrCampoCentroUtil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdCliente_PropDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & StrCampoCarpeta &
                " = " & GshrIdCarpeta & " AND P." &
                StrCampoCentroUtil & " = " &
                GshrIdCentroUtil & " AND " & "P." & ClsIdPredioStr.SstrNombreCampoBd &
                " = " & ClsIdPredioAgrupadorStr.SstrNombreCampoBd & " AND " &
                lstrCampoBusqueda & " <> ''"
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador - Propietario", lstrTablaPri,
                lstrTablaSec, lstrCamSelTablaPri, lstrCampSelTablaSec, lstrCampRelPri,
                lstrCampRelSec, lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, False)
    End Sub
    Private Sub SDefineBusquedaPredioAgr_Arren()
        Dim lstrTablaSec As String = ClsPredio.SstrNombreTabla
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrCamposTabSec = {ClsIdPredioAgrupadorStr.SstrNombreCampoBd,
                                 ClsIdClienteArrendatarioDbl.SstrNombreCampoBd}
        Dim lstrCamposTabPri As String() = {"DISTINCT " & ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdClienteArrendatarioDbl.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdClienteArrendatarioDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil
        lstrFiltro &= " AND " & ClsIdClienteArrendatarioDbl.SstrNombreCampoBd & " > 0 AND " &
                lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador - Arrendatario", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, False)
    End Sub
    Private Sub SDefineBusquedaCliente()
        Dim lstrTabla = ClsCliente.SstrNombreTabla
        Dim lstrCamposMostrar = {ClsIdClienteDbl.SstrNombreCampoBd, ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdClienteDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " & lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)

    End Sub
    Private Sub SDefineNombreCliente()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsNotaCr.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgrupador_NotaCrStr.SstrNombreCampoBd,
                                            ClsPrefijo_NotaCrStr.SstrNombreCampoBd,
                                            ClsIdNotaCrEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_NotaCrDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCamposRetornar As String() = {ClsPrefijo_NotaCrStr.SstrNombreCampoBd,
                                              ClsIdNotaCrEnt.SstrNombreCampoBd}
        Dim lstrFiltro As String = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCamposRetornar, lstrFiltro, False)
    End Sub
    Private Sub SDefinePredioAgr()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsNotaCr.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgrupador_NotaCrStr.SstrNombreCampoBd,
                                            ClsFecha_NotaCrDtm.SstrNombreCampoBd,
                                            ClsPrefijo_NotaCrStr.SstrNombreCampoBd,
                                            ClsIdNotaCrEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_NotaCrDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsIdPredioAgrupador_NotaCrStr.SstrNombreCampoBd
        Dim lstrCamposRetornar As String() = {ClsPrefijo_NotaCrStr.SstrNombreCampoBd,
                                            ClsIdNotaCrEnt.SstrNombreCampoBd}
        Dim lstrFiltro = "S." & StrCampoCarpeta & " = " &
                GshrIdCarpeta.ToString & " AND S." & StrCampoCentroUtil &
                " = " & GshrIdCentroUtil.ToString
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCamposRetornar, lstrFiltro, False)
    End Sub
    Private Sub SBuscarCliente()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
            StrResultadoBusqueda = String.Empty
            SBuscar()
            If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                txtIdCliente.Text = StrResultadoBusqueda
            End If
        End If
    End Sub
#End Region
#Region "Procedimientos Propios"
#Region "Manejo Controles"
    Private Sub SAdicioneCtlsRestringidos()
        ' Items
        SAdicioneControlRestringido(dgrDescuentos)
        SAdicioneControlRestringido(dgrNovedades)
        ' Descuentos y Retenciones
        SAdicioneControlRestringido(txtNroFactura)
        SAdicioneControlRestringido(txtItemFactura)
        SAdicioneControlRestringido(txtTipoDcto)
        SAdicioneControlRestringido(txtValorDcto)
        SAdicioneControlRestringido(bttNuevoDscto)
        SAdicioneControlRestringido(bttAceptarDscto)
        SAdicioneControlRestringido(bttCancelarDscto)
        SAdicioneControlRestringido(bttEliminarDscto)
        SAdicioneControlRestringido(bttEncontrarCliente)
    End Sub
    Private Sub SDeshabiliteControlesActuales()
        Dim lstyEstiloNoHabilitado As Style = FindResource("RecCtlNoHabilitado")
        ' Descuentos y Retenciones
        txtNroFactura.Style = lstyEstiloNoHabilitado
        txtItemFactura.Style = lstyEstiloNoHabilitado
        txtTipoDcto.Style = lstyEstiloNoHabilitado
        txtValorDcto.Style = lstyEstiloNoHabilitado
        lblDeudaCap.Visibility = Visibility.Collapsed
        lblDeudaMora.Visibility = Visibility.Collapsed
        lblDeudaIva.Visibility = Visibility.Collapsed
        txtDeudaCap.Visibility = Visibility.Collapsed
        txtDeudaIva.Visibility = Visibility.Collapsed
        txtDeudaMora.Visibility = Visibility.Collapsed
        bttNuevoDscto.Visibility = Visibility.Hidden
        bttAceptarDscto.Visibility = Visibility.Hidden
        bttCancelarDscto.Visibility = Visibility.Hidden
        bttEliminarDscto.Visibility = Visibility.Hidden
        bttEncontrarCliente.Visibility = Visibility.Hidden
    End Sub
    Private Sub SVisibiliceControlesNuevos(ablnMuestre As Boolean)
        Dim lvisVisibilidadNuevos As Visibility = Visibility.Hidden
        Dim lvisVisibilidadActual As Visibility = Visibility.Visible
        If ablnMuestre Then
            lvisVisibilidadNuevos = Visibility.Visible
            lvisVisibilidadActual = Visibility.Hidden
            txtIdNota.IsEnabled = False
            cboPref.IsEnabled = False
        End If
        lblUsuarioGenero.Visibility = lvisVisibilidadActual
        txtUsuarioGenero.Visibility = lvisVisibilidadActual
        bttEncontrarCliente.Visibility = lvisVisibilidadNuevos
        ' Predio agrupador
        txtPredioAgr.Visibility = lvisVisibilidadActual
        cboPredioAgru.Visibility = lvisVisibilidadNuevos
        lblNroNotaRCr.Visibility = lvisVisibilidadActual
        txtNroNotaRCr.Visibility = lvisVisibilidadActual
        ' Generales
        lblEstado.Visibility = lvisVisibilidadActual
        txtEstado.Visibility = lvisVisibilidadActual
        lblValorNotaCr.Visibility = lvisVisibilidadActual
        txtValorNotaCr.Visibility = lvisVisibilidadActual
        lblValorNotaCr.Visibility = lvisVisibilidadActual
        txtValorNotaCr.Visibility = lvisVisibilidadActual
        '
        tbiNovedades.IsEnabled = Not ablnMuestre
        tbcNotaCr.SelectedIndex = 0
    End Sub
    Private Sub SVisibiliceBttDscto(ablnMuestre As Boolean)
        Dim lvisVisibilidad = If(ablnMuestre, Visibility.Visible, Visibility.Hidden)
        bttNuevoDscto.Visibility = lvisVisibilidad
        bttAceptarDscto.Visibility = lvisVisibilidad
        bttCancelarDscto.Visibility = lvisVisibilidad
        If ablnMuestre AndAlso MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro =
                EnuModoNotaCr.EnuPorValor Then
            bttEliminarDscto.Visibility = Visibility.Hidden
        Else
            bttEliminarDscto.Visibility = lvisVisibilidad
        End If
        ' Descuentos y Retenciones
        If MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro = EnuModoNotaCr.EnuPorFactura Then
            lblDeudaCap.Visibility = lvisVisibilidad
            lblDeudaMora.Visibility = lvisVisibilidad
            lblDeudaIva.Visibility = lvisVisibilidad
            txtDeudaCap.Visibility = lvisVisibilidad
            txtDeudaIva.Visibility = lvisVisibilidad
            txtDeudaMora.Visibility = lvisVisibilidad
        End If
    End Sub
    Private Sub SHabiliteBttDscto()
        bttNuevoDscto.IsEnabled = FblnGeneralesOk() AndAlso
                StcValidValido(EnuValidEntradaDef.enuCausarInt)
        bttCancelarDscto.IsEnabled = MenuEstadoItem = EnuEstadoItem.enuCreandoItem
        bttEliminarDscto.IsEnabled = dgrDescuentos.SelectedIndex >= 0
        bttAceptarDscto.IsEnabled = FblnItemNotaOk() AndAlso MenuEstadoItem =
                EnuEstadoItem.enuCreandoItem
    End Sub
    Private Sub SHabiliteCtlsItems()
        SVisibiliceCtrlItem()
        Select Case MenuEstadoItem
            Case EnuEstadoItem.enuConsultandoItem
                cboNroFactNuevo.Style = FindResource("RecCtlNoHabilitado")
                cboItemFactNuevo.Style = FindResource("RecCtlNoHabilitado")
                cboTipoDsctoNuevo.Style = FindResource("RecCtlNoHabilitado")
                txtValorDctoNuevo.Style = FindResource("RecCtlNoHabilitado")
                lblDeudaCap.Visibility = Visibility.Collapsed
                lblDeudaMora.Visibility = Visibility.Collapsed
                lblDeudaIva.Visibility = Visibility.Collapsed
                txtDeudaCap.Visibility = Visibility.Collapsed
                txtDeudaIva.Visibility = Visibility.Collapsed
                txtDeudaMora.Visibility = Visibility.Collapsed
                If dgrDescuentos.Items.Count > 0 Then
                    dgrDescuentos.SelectedIndex = 0
                End If
                bttEliminarDscto.IsEnabled = dgrDescuentos.Items.Count > 0
            Case EnuEstadoItem.enuCreandoItem
                If MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro = EnuModoNotaCr.EnuPorFactura Then
                    cboNroFactNuevo.Style = FindResource("RecCtlHabilitado")
                    cboItemFactNuevo.Style = FindResource("RecCtlHabilitado")
                Else
                    cboNroFactNuevo.Style = FindResource("RecCtlNoHabilitado")
                    cboItemFactNuevo.Style = FindResource("RecCtlNoHabilitado")
                End If
                cboTipoDsctoNuevo.Style = FindResource("RecCtlHabilitado")
                txtValorDctoNuevo.Style = FindResource("RecCtlHabilitado")
                lblDeudaCap.Visibility = Visibility.Visible
                lblDeudaMora.Visibility = Visibility.Visible
                lblDeudaIva.Visibility = Visibility.Visible
                txtDeudaCap.Visibility = Visibility.Visible
                txtDeudaIva.Visibility = Visibility.Visible
                txtDeudaMora.Visibility = Visibility.Visible
        End Select
    End Sub
    Private Sub SVisibiliceCtrlItem()
        Dim lvisOculto As Visibility = Visibility.Hidden
        Dim lvisVisible As Visibility = Visibility.Visible
        txtNroFactura.Visibility = lvisOculto
        txtItemFactura.Visibility = lvisOculto
        cboNroFactNuevo.Visibility = lvisOculto
        cboItemFactNuevo.Visibility = lvisOculto
        txtTipoDcto.Visibility = lvisOculto
        cboTipoDsctoNuevo.Visibility = lvisOculto
        txtValorDcto.Visibility = lvisOculto
        txtValorDctoNuevo.Visibility = lvisOculto
        Select Case MenuEstadoItem
            Case EnuEstadoItem.enuConsultandoItem
                txtNroFactura.Visibility = lvisVisible
                txtItemFactura.Visibility = lvisVisible
                txtTipoDcto.Visibility = lvisVisible
                txtValorDcto.Visibility = lvisVisible
            Case EnuEstadoItem.enuCreandoItem
                cboNroFactNuevo.Visibility = lvisVisible
                cboItemFactNuevo.Visibility = lvisVisible
                cboTipoDsctoNuevo.Visibility = lvisVisible
                txtValorDctoNuevo.Visibility = lvisVisible
        End Select
    End Sub
    Private Sub SVisibiliceCtrlMora(ablnHabilite As Boolean)
        If ablnHabilite Then
            bttCausarIntereses.Visibility = Visibility.Visible
            lblIntPorCausar.Visibility = Visibility.Visible
            txtIntPorCausar.Visibility = Visibility.Visible
        Else
            bttCausarIntereses.Visibility = Visibility.Collapsed
            lblIntPorCausar.Visibility = Visibility.Collapsed
            txtIntPorCausar.Visibility = Visibility.Collapsed
        End If
    End Sub
    Private Sub SCambieNombresDocCobro()
        If GobjParametros.BlnEFacAutorizado Then
            lblIdFactura.Content = "Número Factura"
            lblItemFactura.Content = "Item Factura"
            dgtNroFac.Header = "Nro. Factura"
            dgtItemFac.Header = "Item Factura"
            dgtNroFacMC.Header = "Número Factura"
            dgtItemFacMC.Header = "Item Fac"
        Else
            lblIdFactura.Content = "Número Cta. Cobro"
            lblItemFactura.Content = "Item Cta. Cobro"
            dgtNroFac.Header = "Nro. Cuenta Cobro"
            dgtItemFac.Header = "Item Cuenta Cobro"
            dgtNroFacMC.Header = "Nro. Cta. Cobro"
            dgtItemFacMC.Header = "Item C.C."
        End If
    End Sub
    Private Sub SCapturePorFactura()
        SVisibiliceCtrlItem()
        SVisibiliceBttDscto(True)
        SHabiliteBttDscto()
    End Sub
    Private Sub SCapturePorValor()
        MobjObjetoWin.ColItemsNotaCr.Clear()
        If MobjItemNotaCrActual IsNot Nothing Then
            MobjItemNotaCrActual.SVacie()
        End If
        cboNroFactNuevo.Items.Clear()
        cboItemFactNuevo.Items.Clear()
        SHabiliteCtlsItems()
        SVisibiliceBttDscto(True)
    End Sub
#End Region
#Region "Manejo Items"
    Private Sub SVacieItem()
        MblnPoblandoCbo = True
        cboNroFactNuevo.SelectedIndex = 0
        cboItemFactNuevo.SelectedIndex = 0
        cboTipoDsctoNuevo.SelectedIndex = 0
        txtDeudaCap.Content = Format(0, "c")
        txtDeudaIva.Content = Format(0, "c")
        txtDeudaMora.Content = Format(0, "c")
        txtValorDctoNuevo.Text = String.Empty
        txtValorDcto.Text = String.Empty
        txtNroFactura.Text = String.Empty
        txtItemFactura.Text = String.Empty
        txtTipoDcto.Text = String.Empty
        MblnPoblandoCbo = False
        If Me.EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
            MenuEstadoItem = EnuEstadoItem.enuConsultandoItem
        End If
        SValide()
    End Sub
    Private Sub SNuevoItem()
        SVacieItem()
        MenuEstadoItem = EnuEstadoItem.enuCreandoItem
        MblnPoblandoCbo = True
        MdecBaseRet = 0
        MdblTasaRet = 0
        If MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro = EnuModoNotaCr.EnuPorFactura Then
            MobjItemNotaCrActual = MobjObjetoWin.FobjNewItemNotaCr
            ObjHijoObjWin = MobjItemNotaCrActual
        Else
            txtDeudaCap.Content = Format(MobjObjetoWin.DecValorCrMaxCap, "c")
            txtDeudaMora.Content = Format(MobjObjetoWin.DecValorCrMaxInt, "c")
            txtDeudaIva.Content = Format(MobjObjetoWin.DecValorCrMaxIva, "c")
        End If
        SHabiliteCtlsItems()
        cboNroFactNuevo.Focus()
        MblnPoblandoCbo = False
        SValide()
    End Sub
    Private Sub SAcepteItem()
        If MenuEstadoItem = EnuEstadoItem.enuCreandoItem AndAlso FblnItemNotaOk() Then
            SEstablezcaDscto()
            MobjObjetoWin.SAdicioneNuevoItem(MobjItemNotaCrActual)
            MenuEstadoItem = EnuEstadoItem.enuConsultandoItem
            SPuebleDtbItems()
            grdDetalles.DataContext = MdtbItems
            SHabiliteCtlsItems()
            SVisibiliceBttDscto(True)
            SHabiliteBttDscto()
            SMuestreDatos()
            bttNuevoDscto.Focus()
        End If
    End Sub
    Private Sub SAcepteValor()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            MobjObjetoWin.SApliqueDescuentoValor()
            SPuebleDtbItems()
            grdDetalles.DataContext = MdtbItems
            SCanceleItem()
            txtValorDctoNuevo.Text = Format("0", "c")
            cboTipoDsctoNuevo.SelectedIndex = 0
            SValide()
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
    End Sub
    Private Sub SCanceleItem()
        If MenuEstadoItem = EnuEstadoItem.enuCreandoItem Then
            If dgrDescuentos.Items.Count > 0 Then
                dgrDescuentos.SelectedIndex = 0
                MobjItemNotaCrActual = MobjObjetoWin.ColItemsNotaCr(1)
            Else
                MobjItemNotaCrActual = Nothing
                SVacieItem()
            End If
            SLevanteEveOk()
            SMuestreDatos()
            SValide()
            MenuEstadoItem = EnuEstadoItem.enuConsultandoItem
        End If
        SHabiliteCtlsItems()
        bttNuevoDscto.Focus()
        SValide()
    End Sub
    Private Sub SElimineItem()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
            If dgrDescuentos.Items.Count > 0 AndAlso Not IsNothing(dgrDescuentos.SelectedItem) Then
                MblnEliminandoItem = True
                Dim lshrIdItemfac As Short = dgrDescuentos.SelectedIndex + 1
                MobjObjetoWin.SElimineItemNotaCr(lshrIdItemfac)
                SPuebleDtbItems()
                MblnEliminandoItem = False
                If dgrDescuentos.Items.Count > 0 Then
                    dgrDescuentos.SelectedIndex = 0
                Else
                    MobjItemNotaCrActual = Nothing
                    SVacieItem()
                End If
                SMuestreDatos()
                bttNuevoDscto.Focus()
            End If
        End If
    End Sub
    Private Shared Function FdtbItems() As DataTable
        Dim ldtbItems As New DataTable
        Dim ldclNroFactura As New DataColumn("NroFac", System.Type.GetType("System.String"))
        Dim ldclItemFact As New DataColumn("Detalle", System.Type.GetType("System.String"))
        Dim ldclIdTipoDcto As New DataColumn("IdTipoDcto", System.Type.GetType("System.Byte"))
        Dim ldclTipoDcto As New DataColumn("Dato", System.Type.GetType("System.String"))
        Dim ldclBase As New DataColumn("BaseDscto", System.Type.GetType("System.Decimal"))
        Dim ldclTasa As New DataColumn("TasaDscto", System.Type.GetType("System.Double"))
        Dim ldclValorDcto As New DataColumn("Valor", System.Type.GetType("System.Decimal"))
        ldtbItems.Columns.Add(ldclNroFactura)
        ldtbItems.Columns.Add(ldclItemFact)
        ldtbItems.Columns.Add(ldclIdTipoDcto)
        ldtbItems.Columns.Add(ldclTipoDcto)
        ldtbItems.Columns.Add(ldclBase)
        ldtbItems.Columns.Add(ldclTasa)
        ldtbItems.Columns.Add(ldclValorDcto)
        Return ldtbItems
    End Function
    Private Sub SPuebleDtbItems()
        Dim lobjItemNota As ClsItemNotaCr = Nothing
        Dim lcolItemsNota As Collection = MobjObjetoWin.ColItemsNotaCr
        MdtbItems.Rows.Clear()
        For i = 1 To lcolItemsNota.Count
            lobjItemNota = lcolItemsNota(i)
            Dim ldrwNuevoItem As DataRow = MdtbItems.NewRow
            With lobjItemNota
                ldrwNuevoItem("NroFac") = .StrNroFactura_ItemNotaCr
                ldrwNuevoItem("Detalle") = .StrItemFactura
                ldrwNuevoItem("IdTipoDcto") = .ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
                ldrwNuevoItem("Dato") = .ObjIdTipoDscto_ItemNotaCrByt.ToString
                ldrwNuevoItem("BaseDscto") = .ObjBaseDscto_NotaCrDec.ObjValorPro
                ldrwNuevoItem("TasaDscto") = .ObjTasaDscto_ItemNotaCrDbl.ObjValorPro
                ldrwNuevoItem("Valor") = .ObjValor_ItemNotaCrDec.ObjValorPro
            End With
            MdtbItems.Rows.Add(ldrwNuevoItem)
        Next
    End Sub
    Private Sub SEstablezcaDscto()
        Dim lstrNroFact = cboNroFactNuevo.SelectedItem
        Dim lstrItemFac As String = cboItemFactNuevo.SelectedItem.ToString
        Dim lshrIdItemFac As Short = CType(lstrItemFac.Substring(0, lstrItemFac.IndexOf("-")), Short)
        Dim lenuTipoDscto As EnuTipoDescuentoDef = MobjItemNotaCrActual.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro
        Dim ldecValorDscto As Decimal = MobjItemNotaCrActual.ObjValor_ItemNotaCrDec.ObjValorPro
        Dim ldecBase = 0D
        Dim ldblTasa = ClsOrionCop.FdblTasaDscto(lstrNroFact, lshrIdItemFac, lenuTipoDscto,
                ldecValorDscto, ldecBase)
        MobjItemNotaCrActual.ObjBaseDscto_NotaCrDec.ObjValorPro = ldecBase
        MobjItemNotaCrActual.ObjTasaDscto_ItemNotaCrDbl.ObjValorPro = ldblTasa
    End Sub
#End Region
#Region "Generales"
    Private Sub SInicialiceNota()
        If MobjObjetoWin IsNot Nothing Then
            If cboPref.SelectedItem Is Nothing Then
                cboPref.SelectedIndex = 0
            End If
            Dim lstrPref As String = cboPref.SelectedItem
            If MobjObjetoWin.ObjPrefijo_NotaCrStr.ToString <> lstrPref Then
                ObjObjetoWin = New ClsNotaCr(lstrPref)
                MobjObjetoWin = ObjObjetoWin
                If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                    MobjObjetoWin.SVayaAlUltimo()
                End If
            End If
        End If
    End Sub
    Private Sub SAbraNotaCr()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                If txtIdNota.Text <> MobjObjetoWin.ObjIdNotaCrEnt.ToString() Then
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
        If Not IsNothing(MobjObjetoWin) AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.CenuConsultando Then
            grdDetalles.DataContext = MobjObjetoWin.FdtbItemsNotaCr(True)
            tbiNovedades.DataContext = MobjObjetoWin.DtbNovedades
            SOrdeneDataGrid(dgrNovedades, dgrNovedades.Columns(0),
                    ClsIdNovedadShr.SstrNombreCampoBd,
                    ListSortDirection.Ascending)
        End If
    End Sub
    Private Sub SPuebleComboBoxes()
        MblnPoblandoCbo = True
        Dim ldrwConst = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuTipoDescuento)
        SPuebleComboBox(ldrwConst, cboTipoDsctoNuevo)
        cboTipoDsctoNuevo.SelectedIndex = 0
        ldrwConst = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuModoNotaCr)
        SPuebleComboBox(ldrwConst, cboModoNota)
        ldrwConst = ClsOrionCop.FdrwPrefDoc(EnuTipoDocOri.EnuNotaCr)
        SPuebleComboBox(ldrwConst, cboPref)
        MblnPoblandoCbo = False
    End Sub
    Private Sub SPuebleCboPredAgru(ByRef astrMens As String)
        If Not IsNothing(MobjObjetoWin.ObjClienteNotaCr) Then
            Dim lstrPredAgrConSaldo As String() = {}
            MblnPoblandoCbo = True
            cboPredioAgru.Items.Clear()
            cboPredioAgru.Items.Add(My.Resources.Ninguno)
            If MobjObjetoWin.ObjClienteNotaCr.BlnExiste Then
                lstrPredAgrConSaldo =
                        MobjObjetoWin.ObjClienteNotaCr.FstrPrediosAgruClienteConFacturas(True)
                For Each lstrPreAgr As String In lstrPredAgrConSaldo
                    If String.IsNullOrEmpty(lstrPreAgr) Then
                        cboPredioAgru.Items.Add(GCSTRSINPA)
                    Else
                        cboPredioAgru.Items.Add(lstrPreAgr)
                    End If
                Next
            End If
            MblnPoblandoCbo = False
            cboPredioAgru.SelectedIndex = 0
            If MobjCliente IsNot Nothing Then
                If lstrPredAgrConSaldo.Length = 0 Then
                    astrMens = "El Cliente ingresado no tiene deudas por pagar!"
                Else
                    SPuebleComboFacturas()
                End If
            End If
        End If
    End Sub
    Private Sub SPuebleComboFacturas()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
            Dim lstrServicios As String() = {"A"}, lstrIdPreAgr = String.Empty
            MblnPoblandoCbo = True
            cboNroFactNuevo.Items.Clear()
            cboNroFactNuevo.Items.Add(My.Resources.Ninguno)
            If MobjObjetoWin.ObjIdPredioAgrupador_NotaCrStr.BlnEsValido Then
                lstrIdPreAgr = MobjObjetoWin.ObjIdPredioAgrupador_NotaCrStr.ObjValorPro
                If Not IsNothing(MobjCliente) Then
                    Dim lstrNrosFac = MobjCliente.FstrIdFacturasVivas(lstrIdPreAgr, lstrServicios)
                    For Each lstrIdFra As String In lstrNrosFac
                        cboNroFactNuevo.Items.Add(lstrIdFra)
                    Next
                End If
            End If
            MblnPoblandoCbo = False
            cboNroFactNuevo.SelectedIndex = 0
        End If
    End Sub
    Private Sub SPuebleItemsFra()
        MblnPoblandoCbo = True
        cboItemFactNuevo.Items.Clear()
        cboItemFactNuevo.Items.Add(My.Resources.Ninguno)
        If Not IsNothing(cboNroFactNuevo.SelectedItem) Then
            If cboNroFactNuevo.SelectedItem <> My.Resources.Ninguno Then
                Dim lstrPref = ClsPanorama.FstrPrefijoDcto(cboNroFactNuevo.SelectedItem)
                Dim lentIdFra = ClsPanorama.FentIdDcto(cboNroFactNuevo.SelectedItem)
                Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFra}
                MobjFactura.SAbra(lobjValorLlave)
                If Not MobjFactura.BlnExiste Then
                    Throw New ErrorInesperadoPanLException("Factura en RC no existe")
                End If
                Dim lstrItem = String.Empty
                For Each lobjItemFra As ClsItemFactura In MobjFactura.ColItemsFactura
                    lstrItem = lobjItemFra.ObjIdItemFacturaShr.ToString & "-" &
                                lobjItemFra.ObjDetalle_ItemFactStr.ObjValorPro
                    cboItemFactNuevo.Items.Add(lstrItem)
                Next
            End If
        End If
        MblnPoblandoCbo = False
        cboItemFactNuevo.SelectedIndex = 0
    End Sub
    Private Sub SMuestreUsuarios()
        With MobjObjetoWin
            If MobjObjetoWin.BlnExiste AndAlso EnuOperacionEnWin =
                    EnuOperacionEnVentana.CenuConsultando Then
                txtUsuarioGenero.Content = MobjObjetoWin.ObjIdUsuario_NotaCrStr.ObjValorPro
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
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            txtNroNotaRCr.Content = MobjObjetoWin.StrNroNotaRCr
            If MobjObjetoWin.ObjAnuladoBln.ObjValorPro Then
                txtEstado.Style = FindResource("RecDocAnulado")
            Else
                If Not String.IsNullOrEmpty(txtNroNotaRCr.Content) Then
                    txtEstado.Style = FindResource("RecDocRversado")
                    SHabiliteBotonTlb(False, HbttAnular)
                    SHabiliteMenu(False, HmnuAnular)
                Else
                    txtEstado.Style = FindResource("RecDocNormal")
                End If
            End If
        End If
    End Sub
    Private Sub SMuestreVlrDscto()
        If cboTipoDsctoNuevo.SelectedIndex > 0 Then
            MdecBaseRet = 0
            MdblTasaRet = 0
            Dim lenuTipoDscto As EnuTipoDescuentoDef = cboTipoDsctoNuevo.SelectedIndex
            If lenuTipoDscto >= EnuTipoDescuentoDef.EnuReteFuente AndAlso
                    lenuTipoDscto < EnuTipoDescuentoDef.EnuDsctoPP Then
                If Not IsNothing(MobjItemFac) Then
                    If Not MobjItemFac.FblnRetencionAplicada(lenuTipoDscto) Then
                        Dim ldecVlrRet As Decimal = MobjItemFac.FdecValorRetencion(lenuTipoDscto,
                                MdecBaseRet, MdblTasaRet)
                        txtValorDctoNuevo.Text = Format(ldecVlrRet, "c")
                        MobjItemNotaCrActual.ObjBaseDscto_NotaCrDec.ObjValorPro = MdecBaseRet
                        MobjItemNotaCrActual.ObjTasaDscto_ItemNotaCrDbl.ObjValorPro = MdblTasaRet
                        MobjItemNotaCrActual.ObjValor_ItemNotaCrDec.ObjValorPro = ldecVlrRet
                    Else
                        SLevanteEveNoti("La Retención ya fue aplicada!", "", 0,
                                EnuSeveridadNot.EnuInformacion)
                    End If
                End If
            ElseIf lenuTipoDscto = EnuTipoDescuentoDef.EnuCancelaIva Then
                Dim ldecDeudaIva As Decimal = MobjItemFac.FdecDeudaIva
                MobjItemNotaCrActual.ObjValor_ItemNotaCrDec.ObjValorPro = ldecDeudaIva
                txtValorDctoNuevo.Text = Format(ldecDeudaIva, "c")
                MsgBox("ADVERTENCIA" & vbCrLf & "El valor del IVA sera llevado al gasto!",
                        vbOKOnly, "Iva al gasto")
            Else
                txtValorDctoNuevo.Text = Format(0, "c")
                MobjItemNotaCrActual.ObjBaseDscto_NotaCrDec.ObjValorPro = MdecBaseRet
                MobjItemNotaCrActual.ObjTasaDscto_ItemNotaCrDbl.ObjValorPro = MdblTasaRet
            End If
        Else
            txtValorDctoNuevo.Text = Format(0, "c")
        End If
    End Sub
    Private Sub SMuestreDeuda(astrIdItemFac As String)
        If astrIdItemFac <> My.Resources.Ninguno Then
            Dim lstrIdItemFac = astrIdItemFac.Substring(0, astrIdItemFac.IndexOf("-"))
            MobjItemFac = MobjFactura.ColItemsFactura(lstrIdItemFac)
            Dim ldecDeudaCapSinIva = MobjItemFac.FdecDeudaCapital -
                    MobjItemFac.FdecDeudaIvaCapital
            ' AVV Revisar
            Dim ldecDeudaIntSinIva = 0 'MobjItemFac.FdecDeudaIntMora -
            ' MobjItemFac.FdecDeudaIvaInt
            txtDeudaCap.Content = Format(ldecDeudaCapSinIva, "c")
            txtDeudaIva.Content = Format(MobjItemFac.FdecDeudaIva, "c")
            txtDeudaMora.Content = Format(ldecDeudaIntSinIva, "c")
            MobjItemNotaCrActual.ObjIdItemFac_ItemNotaCrShr.ObjValorPro = CType(lstrIdItemFac, Short)
        Else
            txtDeudaCap.Content = Format(0, "c")
            txtDeudaIva.Content = Format(0, "c")
            txtDeudaMora.Content = Format(0, "c")
        End If
    End Sub
    Private Sub SEnvieEmailNotaCr()
        Dim lstrMens = String.Empty
        If ClsPanorama.FblnEmailsHabilitado Then
            If MobjObjetoWin.ObjClienteNotaCr.ObjRecibeDocsPorEmailBln.ObjValorPro Then
                SEnvieCorreo(Me, EnuTipoCorreoE.EnuNCR,
                        MobjObjetoWin.ObjClienteNotaCr.ObjIdClienteDbl.ObjValorPro,
                        MobjObjetoWin.ObjIdPredioAgrupador_NotaCrStr.ObjValorPro,
                        MobjObjetoWin.StrNumeroNotaCr)
            Else
                lstrMens = "El Cliente de la Nota Crédito no tiene habilitado el envio de " &
                                              "Documentos por Email!"
            End If
        Else
            lstrMens = "Aún no tiene instalada la Aplicación para el envío de Documentos por Email!"
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Function FblnGeneralesOk() As Boolean
        Dim lblnGenOk = StcValidValido(EnuValidEntradaDef.enuCliente) AndAlso
                StcValidValido(EnuValidEntradaDef.enuFecha) AndAlso
                StcValidValido(EnuValidEntradaDef.enuPreAgr) AndAlso
                StcValidValido(EnuValidEntradaDef.enuModoNota)
        Return lblnGenOk
    End Function
    Private Sub SRegistreCliente(ByRef astrMens As String)
        With MobjObjetoWin
            If .ObjIdCliente_NotaCrDbl.ToString() <> txtIdCliente.Text Then
                .ObjIdCliente_NotaCrDbl.ObjValorPro = txtIdCliente.Text
                SVacieItem()
                SValide()
                If .ObjIdCliente_NotaCrDbl.BlnEsValido Then
                    MobjCliente = .ObjClienteNotaCr
                    txtNombreCliente.Content = MobjCliente.ObjNombreCompletoStr.ObjValorPro
                Else
                    txtNombreCliente.Content = String.Empty
                    MobjCliente = Nothing
                End If
                SPuebleCboPredAgru(astrMens)
                .ObjIdPredioAgrupador_NotaCrStr.ObjValorPro = cboPredioAgru.SelectedItem
            End If
        End With
    End Sub
    Private Sub SRegistrePredAgru()
        Dim lstrIdPredioAgr = If(cboPredioAgru.SelectedItem = GCSTRSINPA,
                String.Empty, cboPredioAgru.SelectedItem)
        If cboPredioAgru.SelectedItem <>
                MobjObjetoWin.ObjIdPredioAgrupador_NotaCrStr.ObjValorPro Then
            SVisibiliceCtrlMora(False)
            SVacieItem()
            SHabiliteCtlsItems()
            cboModoNota.SelectedIndex = 0
            MobjItemNotaCrActual = Nothing
        End If
        MobjObjetoWin.ObjIdPredioAgrupador_NotaCrStr.ObjValorPro = lstrIdPredioAgr
        SValide()
    End Sub
    Private Sub SRegistreModoNota(ByRef astrMens As String)
        MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro = cboModoNota.SelectedIndex
        StcValidValido(EnuValidEntradaDef.enuModoNota) = MobjObjetoWin.ObjModoNotaCrByt.BlnEsValido
        If FblnGeneralesOk() Then
            SValideMoraPorCausar()
            If Not StcValidValido(EnuValidEntradaDef.enuCausarInt) Then
                SVisibiliceCtrlMora(True)
                astrMens = "Se deben causar los Intereses de Mora a la Fecha!"
            Else
                SVisibiliceCtrlMora(False)
                If MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro = EnuModoNotaCr.EnuPorFactura Then
                    SCapturePorFactura
                Else
                    SCapturePorValor()
                End If
            End If
        End If
    End Sub
    Private Function FblnItemNotaOk() As Boolean
        Dim lblnItemNotaOk = False
        If MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro = EnuModoNotaCr.EnuPorFactura Then
            With MobjItemNotaCrActual
                If Not IsNothing(MobjItemNotaCrActual) Then
                    lblnItemNotaOk = .ObjPrefijoFact_ItemNotaCrStr.BlnEsValido AndAlso
                        .ObjIdFactura_ItemNotaCrEnt.BlnEsValido AndAlso
                        .ObjIdTipoDscto_ItemNotaCrByt.BlnEsValido AndAlso
                        .ObjValor_ItemNotaCrDec.BlnEsValido
                Else
                    lblnItemNotaOk = False
                End If
            End With
        Else
            lblnItemNotaOk = MobjObjetoWin.FblnValorDsctoValido
        End If
        Return lblnItemNotaOk
    End Function
    Private Sub SValideMoraPorCausar()
        MobjObjetoWin.ObjFecha_NotaCrDtm.ObjValorPro = dtpFechaNotaCr.SelectedDate
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
            If Not MobjObjetoWin.FblnDeudaSuspendida Then
                Dim ldecMoraPorCausar = MobjObjetoWin.FdecInteresesMoraPorCausar
                StcValidValido(EnuValidEntradaDef.enuCausarInt) = ldecMoraPorCausar = 0
                txtIntPorCausar.Content = Format(ldecMoraPorCausar, "c")
            Else
                StcValidValido(EnuValidEntradaDef.enuCausarInt) = True
            End If
        End If
    End Sub
    Private Sub SCauseMora()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lenuSevNot As EnuSeveridadNot
        Dim ldecValorCausado = 0D, lblnNoHayError As Boolean
        Try
            ldecValorCausado = MobjObjetoWin.SCauseMoraCliente(lstrMens)
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                SVisibiliceCtrlMora(False)
                StcValidValido(EnuValidEntradaDef.enuCausarInt) = True
                If String.IsNullOrEmpty(lstrMens) Then
                    lstrMens = "Se causaron Intereses de Mora por un valor de " &
                            Format(ldecValorCausado, "c")
                    lenuSevNot = EnuSeveridadNot.EnuInformacion
                Else
                    lenuSevNot = EnuSeveridadNot.EnuAdvertencia
                End If
                If GobjParametros.BlnEFacAutorizado AndAlso
                        MobjObjetoWin.BlnAfectaFrasRegEFac Then
                    If ldecValorCausado > 0 Then
                        SProceseNotasDbApi()
                    End If
                End If
                If MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro = EnuModoNotaCr.EnuPorFactura Then
                    SCapturePorFactura()
                Else
                    SCapturePorValor()

                End If
            Else
                lenuSevNot = EnuSeveridadNot.EnuExcep
            End If
            SValide()
            If Not String.IsNullOrEmpty(lstrMens) Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, lenuSevNot)
            End If
        End Try
    End Sub
#End Region
#End Region
#Region "eFactura"
    Private Sub SHabiliteMenuEFac()
        If MnuEstadoProvEfac IsNot Nothing Then
            MnuEstadoProvEfac.Visibility = Visibility.Collapsed
        End If
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If Not HblnCargandoForma AndAlso MobjObjetoWin.BlnEsDocEle AndAlso
                    MobjObjetoWin.BlnAfectaFrasRegEFac Then
                Dim lenuEstadoEFac As EnuEstadoEDoc = MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro
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
                    Case EnuEstadoEDoc.EnuEnviada
                        lblnProcesar = False
                    Case Else
                        Exit Select
                End Select
                If lblnProcesar Then
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        lblnProcesar = MsgBox(lstrMens, vbYesNo, "Registrar Documento?") = vbYes
                        If lblnProcesar Then
                            If lenuEstadoEFac = EnuEstadoEDoc.EnuInvalida OrElse
                                        lenuEstadoEFac = EnuEstadoEDoc.EnuErrorFtp Then
                                MobjObjetoWin.SHabiliteProcesarEFac()
                            End If
                        End If
                    End If
                End If
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
    Private Sub SProceseNotasCrApi()
        Dim lstrMens = String.Empty
        SProceseEFac(lstrMens)
        SRefresqueWin()
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SProceseNotasDbApi()
        Dim lstrMens = String.Empty
        SProceseEFac(lstrMens)
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SInserteDbAnulaCr()
        If GobjParametros.BlnEFacAutorizado AndAlso MobjObjetoWin.BlnTieneDsctos Then
            If Not IsNothing(MobjObjetoWin.ObjNotaReversionCr) Then
                If MobjObjetoWin.BlnAfectaFrasRegEFac Then
                    Dim lstrMens = String.Empty
                    SProceseEFac(lstrMens)
                    If MsgBox("Desea imprimir la Nota?", vbYesNo + MsgBoxStyle.Question,
                              "Imprimir la Nota anulada?") = vbYes Then
                        SImprima()
                    End If
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                End If
            End If
        End If
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Select Case lelmElemento.Name
                Case "bttCausarIntereses"
                    SCauseMora()
                Case "bttNuevoDscto"
                    SNuevoItem()
                Case "bttCancelarDscto"
                    SCanceleItem()
                Case "bttAceptarDscto"
                    If MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro = EnuModoNotaCr.EnuPorFactura Then
                        SAcepteItem()
                    Else
                        SAcepteValor()
                    End If
                Case "bttEliminarDscto"
                    SElimineItem()
                Case "bttEncontrarCliente"
                    SBuscarCliente()
            End Select
        End If
    End Sub
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                Dim lmnuOpcion As MenuItem = lelmElemento
                If lmnuOpcion.Name = "MnuEnviarPorCorreo" Then
                    SEnvieEmailNotaCr()
                ElseIf lmnuOpcion.Name = "MnuEstadoProvEfac" Then
                    Dim lblnV1 = (MobjObjetoWin.ObjVerEFacEnt.ObjValorPro = EnuVerEFac.EnuV1)
                    Mouse.OverrideCursor = Cursors.Wait
                    SRefresqueWin()
                    SMuestreEstadoEFac(MobjObjetoWin.ObjCUDocStr.ObjValorPro,
                            EnuTipoDocOri.EnuNotaCr, MobjObjetoWin.StrIdObjeto, lblnV1,
                            MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro, lstrMens)
                    Mouse.OverrideCursor = Cursors.Arrow
                ElseIf lmnuOpcion.Name = "MnuRegistrarEFac" Then
                    SProceseNotasCrApi()
                ElseIf lmnuOpcion.Name = "MnuReprocesarEDoc" Then
                    If MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                            EnuEstadoEDoc.EnuNoEDoc AndAlso
                            MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                            EnuEstadoEDoc.EnuNoReg AndAlso
                            MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                             EnuEstadoEDoc.EnuInvalida Then
                        If MobjObjetoWin.ObjFecha_NotaCrDtm.ObjValorPro > Today.AddDays(-30) Then
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
        ElseIf TypeOf lelmElemento Is Button Then
            If MblnDejoUltimoControl Then
                If lelmElemento.Name = "hbttCancelar" Then
                    HbttAceptar.Focus()
                ElseIf lelmElemento.Name = "bttCancelarDscto" Then
                    If bttAceptarDscto.IsEnabled Then bttAceptarDscto.Focus()
                End If
            End If
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando AndAlso Not HblnMostrandoDatos Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is DatePicker Then
                MblnDejoUltimoControl = False
                If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
                    Dim lstrMens = String.Empty
                    With MobjObjetoWin
                        Select Case lelmElemento.Name
                            Case "txtIdCliente"
                                SRegistreCliente(lstrMens)
                                cboPredioAgru.Focus()
                            Case "dtpFechaNotaCr"
                                .ObjFecha_NotaCrDtm.ObjValorPro = dtpFechaNotaCr.SelectedDate
                            Case "txtValorDctoNuevo"
                                If MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro =
                                        EnuModoNotaCr.EnuPorFactura Then
                                    If Not IsNothing(MobjItemNotaCrActual) Then
                                        MobjItemNotaCrActual.ObjValor_ItemNotaCrDec.ObjValorPro =
                                            txtValorDctoNuevo.Text
                                        txtValorDctoNuevo.Text = Format(
                                                MobjItemNotaCrActual.ObjValor_ItemNotaCrDec.
                                                ObjValorPro, "c")
                                    End If
                                Else
                                    MobjObjetoWin.DecDsctoPorValor = txtValorDctoNuevo.Text
                                    If MobjObjetoWin.FblnValorDsctoValido Then
                                        txtValorDctoNuevo.Text =
                                                Format(MobjObjetoWin.DecDsctoPorValor, "c")
                                        bttAceptarDscto.IsEnabled = True
                                    Else
                                        bttAceptarDscto.IsEnabled = False
                                    End If
                                End If
                                MblnDejoUltimoControl = True
                            Case "txtComentario"
                                MobjObjetoWin.ObjComentario_NotaCrStr.ObjValorPro = txtComentario.Text
                        End Select
                    End With
                    SMuestreDatos()
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If Not MblnPoblandoCbo AndAlso TypeOf lelmElemento Is ComboBox AndAlso
                Not HblnSeEstaCerrando AndAlso Not HblnMostrandoDatos Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
                Dim lstrMens = String.Empty
                Select Case lelmElemento.Name
                    Case "cboPredioAgru"
                        SRegistrePredAgru()
                        SPuebleComboFacturas()
                    Case "cboModoNota"
                        SRegistreModoNota(lstrMens)
                    Case "cboNroFactNuevo"
                        If Not IsNothing(cboNroFactNuevo.SelectedItem) Then
                            SLevanteEveOk()
                            If cboNroFactNuevo.SelectedItem <> My.Resources.Ninguno Then
                                Dim lstrPrefFact = ClsPanorama.FstrPrefijoDcto(cboNroFactNuevo.SelectedItem)
                                Dim lentIdFact = ClsPanorama.FentIdDcto(cboNroFactNuevo.SelectedItem)
                                MobjFactura.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPrefFact, lentIdFact})
                                If Not MobjFactura.BlnExiste Then
                                    Throw New ErrorInesperadoPanLException("Factura en NCR no existe")
                                End If
                                MobjItemNotaCrActual.SInicialiceObj()
                                MobjItemNotaCrActual.ObjPrefijoFact_ItemNotaCrStr.ObjValorPro = lstrPrefFact
                                MobjItemNotaCrActual.ObjIdFactura_ItemNotaCrEnt.ObjValorPro = lentIdFact
                            Else
                                If Not IsNothing(MobjItemNotaCrActual) Then
                                    MobjItemNotaCrActual.ObjPrefijoFact_ItemNotaCrStr.ObjValorPro = String.Empty
                                    MobjItemNotaCrActual.ObjIdFactura_ItemNotaCrEnt.ObjValorPro = 0
                                End If
                            End If
                            txtDeudaCap.Content = Format(0, "c")
                            txtDeudaMora.Content = Format(0, "c")
                            txtDeudaIva.Content = Format(0, "c")
                            cboTipoDsctoNuevo.SelectedIndex = 0
                            txtValorDctoNuevo.Text = Format(0, "c")
                            SPuebleItemsFra()
                        End If
                    Case "cboItemFactNuevo"
                        If Not IsNothing(cboItemFactNuevo.SelectedItem) Then
                            Dim lstrIdItemFac As String = cboItemFactNuevo.SelectedItem
                            SMuestreDeuda(lstrIdItemFac)
                        Else
                            txtDeudaCap.Content = Format(0, "c")
                            txtDeudaMora.Content = Format(0, "c")
                            txtDeudaIva.Content = Format(0, "c")
                            If MobjItemNotaCrActual IsNot Nothing Then
                                MobjItemNotaCrActual.ObjIdItemFac_ItemNotaCrShr.ObjValorPro = 0
                            End If
                        End If
                    Case "cboTipoDsctoNuevo"
                        If MobjObjetoWin.ObjModoNotaCrByt.ObjValorPro =
                                EnuModoNotaCr.EnuPorFactura Then
                            MobjItemNotaCrActual.ObjIdTipoDscto_ItemNotaCrByt.ObjValorPro =
                                    cboTipoDsctoNuevo.SelectedIndex
                            If MobjItemNotaCrActual.ObjIdTipoDscto_ItemNotaCrByt.BlnEsValido Then
                                SMuestreVlrDscto()
                            End If
                        Else
                            MobjObjetoWin.EnuTipoDsctoPorValor = cboTipoDsctoNuevo.SelectedIndex
                        End If
                End Select
                SMuestreDatos()
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            ElseIf EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                If lelmElemento.Name = "cboPref" Then
                    SInicialiceNota()
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub
    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdNota.KeyDown,
            txtIdCliente.KeyDown
        If e.Key = Key.Return OrElse e.Key = Key.Tab Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If lelmElemento.Name = "txtIdNota" AndAlso EnuOperacionEnWin =
                    EnuOperacionEnVentana.CenuConsultando Then
                SAbraNotaCr()
            End If
            If lelmElemento.Name = "txtIdCliente" AndAlso EnuOperacionEnWin =
                    EnuOperacionEnVentana.CenuCreando AndAlso e.Key = Key.Return Then
                SRegistreCliente(String.Empty)
                If MobjObjetoWin.ObjIdCliente_NotaCrDbl.BlnEsValido Then
                    cboPredioAgru.Focus()
                Else
                    bttEncontrarCliente.Focus()
                End If
                SMuestreDatos()
            End If
        End If
    End Sub
    Private Sub DgrDescuentos_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dgrDescuentos.SelectionChanged
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
            If MenuEstadoItem = EnuEstadoItem.enuConsultandoItem AndAlso Not MblnEliminandoItem Then
                Dim ldrvItemItemNotaCr As DataRowView = dgrDescuentos.SelectedItem
                If Not IsNothing(ldrvItemItemNotaCr) AndAlso MobjObjetoWin.ColItemsNotaCr.Count > 0 Then
                    MobjItemNotaCrActual = MobjObjetoWin.ColItemsNotaCr(dgrDescuentos.SelectedIndex + 1)
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub
    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
            dgrDescuentos.MouseRightButtonUp, dgrNovedades.MouseRightButtonUp
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            Dim lstrNroFac As String, lstrPrefFac As String, lentIdFac As Integer
            ldrvFilaActual = ldgrActual.SelectedItem
            lstrNroFac = ldrvFilaActual("NroFac")
            lstrPrefFac = ClsPanorama.FstrPrefijoDcto(lstrNroFac)
            lentIdFac = ClsPanorama.FentIdDcto(lstrNroFac)
            SAbraFactura(lstrPrefFac, lentIdFac)
        End If
    End Sub
    Private Sub Txt_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtNroNotaRCr.MouseDoubleClick
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Dim lstrPrefDoc = ClsPanorama.FstrPrefijoDcto(txtNroNotaRCr.Content)
        Dim lentIdDoc = ClsPanorama.FentIdDcto(txtNroNotaRCr.Content)
        If lelmElemento.Name = "txtNroNotaRCr" Then
            SAbraNotaRCr(lstrPrefDoc, lentIdDoc)
        End If
    End Sub
#End Region
End Class