Imports OPT.OrionP.PanL
Imports OPT.OrionP.OrionCopL
Public Class WinFacturas
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuIdCliente
        enuIdPreAgru
        enuFechaFact
        enuFechaVenc
        enuServicio
        enuPredio
        enuDetalle
        enuValor
        enuItems
        enuMedPag
        enuForPag
    End Enum
    Private Enum EnuEstadoItemDef As Integer
        None = 0
        enuConsultandoItem
        enuCreandoItem
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsFactura = Nothing
    Private MobjCliente As ClsCliente = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomFacts
    Private MstrNumeroFac As String = String.Empty
    Private MnuEstadoProvEfac As MenuItem = Nothing
    Private MnuReprocesarEDoc As MenuItem = Nothing
    Private MnuEnviarPorCorreo As MenuItem = Nothing
    ' Creacion
    Private MblnPoblandoCombo As Boolean = False
    Private MblnDejoUltimoControl As Boolean = False
    Private MobjItemFactActual As ClsItemFactura = Nothing
    Private MdtbItems As DataTable = Nothing
    Private MenuEstadoItem As EnuEstadoItemDef = EnuEstadoItemDef.None
    Private MblnEliminandoItem As Boolean = False
    Private MblnCapturoFechaFac As Boolean = False
    Private MblnDejoComboBox As Boolean = False
    Private MblnCancelaCrear As Boolean = False
    Private MblnAplicoAjuste As Boolean = False
    Private MblnCausoMora As Boolean = False
    Private MblnCreando As Boolean = False
    Private MstrMensAnulo As String = String.Empty
    Private ReadOnly MwinMW As MWOrionCop = Nothing
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuFacturas
    End Sub
    Public Sub New(awinMW As MWOrionCop)
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuFacturas
        MwinMW = awinMW
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolControlesLlave As New Collection From {
                cboPref, txtIdFactura
        }
        SAdicioneControlesRestringidos()
        SCargueForma(EnuElementosAdicionalesDef.enuBuscar + EnuElementosAdicionalesDef.enuTercero +
                EnuElementosAdicionalesDef.enuImprimir, 11, lcolControlesLlave, Nothing, False)
        BlnSiempreCreando = True
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        txtIdFactura.SelectAll()
        SHabiliteMenuEFac()
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
            Dim lstrPref As String
            If ClsOrionCop.FblnHayPrefacturas Then
                lstrPref = GCSTRPREFPREFACTURA
            Else
                lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuFacturaVenta)
                Dim lstrPrefAnt = "*", lblnExistePref = False
                Dim ldrwPrefs = ClsOrionCop.FdrwPrefDoc(EnuTipoDocOri.EnuFactura)
                For Each ldrwPref As DataRow In ldrwPrefs
                    lstrPrefAnt = ldrwPref("Dato")
                    lblnExistePref = lstrPref = lstrPrefAnt
                    If lblnExistePref Then Exit For
                Next
                If Not lblnExistePref AndAlso lstrPrefAnt <> "*" Then
                    lstrPref = lstrPrefAnt
                End If
            End If
            ObjObjetoWin = New ClsFactura(lstrPref)
            If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                ObjObjetoWin.SVayaAlUltimo()
            End If
        End If
        MobjObjetoWin = ObjObjetoWin
        If Not IsNothing(MobjObjetoWin) Then
            MdtbItems = MobjObjetoWin.FdtbItems
        End If
        SVisibiliceCtrls(False)
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        If Not MobjObjetoWin.FblnEstaVacioOrigenDatos Then
            cnvDatosFactura.DataContext = MobjObjetoWin.DtbItemsFact
        End If
        StcValidaControl(EnuValidEntradaDef.enuIdCliente) = lblIdClienteNue
        StcValidaControl(EnuValidEntradaDef.enuIdPreAgru) = lblIdPredioAgr
        StcValidaControl(EnuValidEntradaDef.enuFechaFact) = lblFechaFac
        StcValidaControl(EnuValidEntradaDef.enuFechaVenc) = lblFechaVen
        StcValidaControl(EnuValidEntradaDef.enuServicio) = lblServicio
        StcValidaControl(EnuValidEntradaDef.enuPredio) = lblPredio
        StcValidaControl(EnuValidEntradaDef.enuDetalle) = lblDetalle
        StcValidaControl(EnuValidEntradaDef.enuValor) = lblValor
        StcValidaControl(EnuValidEntradaDef.enuItems) = lblItems
        StcValidaControl(EnuValidEntradaDef.enuMedPag) = lblIdMedPago
        StcValidaControl(EnuValidEntradaDef.enuForPag) = lblIdFormaPago
        dtpFechaFactura.SelectedDate = MobjObjetoWin.ObjFechaFacturaDtm.ObjValorPro
        dtpFechaVen.SelectedDate = MobjObjetoWin.ObjFechaVencimientoDtm.ObjValorPro
        SPuebleCboPref()
        cnvItemFac.DataContext = MdtbItems
        tbiFactura.IsSelected = True
        '
        HbttAceptar.TabIndex = 20
        HbttCancelar.TabIndex = 21
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            HblnMostrandoDatos = True
            If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                SLevanteEveNoti("No hay Facturas para ser mostradas!", "", 0,
                EnuSeveridadNot.EnuInformacion)
                cboPref.IsEnabled = False
                txtIdFactura.IsEnabled = False
            Else
                With MobjObjetoWin
                    MstrNumeroFac = .StrNumeroFactura
                    cboPref.SelectedItem = .ObjPrefijo_FactStr.ObjValorPro
                    txtIdFactura.Text = .ObjIdFacturaEnt.ToString()
                    txtIdCliente.Content = .ObjIdCliente_FactDbl.ObjValorPro
                    txtNombreCliente.Content = .ObjClienteFactura.ObjNombreCompletoStr.ToString
                    txtNombreCliente.ToolTip = txtNombreCliente.Content
                    txtModoFacturacion.Content = .ObjIdModoFacturacionByt.ToString
                    If String.IsNullOrEmpty(.ObjIdPredioAgrupador_FacStr.ObjValorPro) Then
                        txtIdPredioAgru.Content = GCSTRSINPA
                    Else
                        txtIdPredioAgru.Content = .ObjIdPredioAgrupador_FacStr.ObjValorPro
                    End If
                    txtSaldo.Content = Format(.DecDeuda, "c")
                    txtSaldoCap.Content = Format(.FdecDeudaCapital, "c")
                    txtSaldoMora.Content = Format(.FdecDeudaIntMora, "c")
                    txtValorFactura.Content = Format(.ObjValor_FactDec.ObjValorPro, "c")
                    txtValorF.Content = Format(.ObjValor_FactDec.ObjValorPro, "c")
                    txtDebitos.Content = Format(.ObjDebitos_FactDec.ObjValorPro, "c")
                    txtCreditos.Content = Format(.ObjCreditos_FactDec.ObjValorPro, "c")
                    txtFechaFactura.Content = Format(.ObjFechaFacturaDtm.ObjValorPro,
                            GCSTRFMTFECHASIMPLE)
                    txtFechaFacturaP.Content = Format(.ObjFechaFacturaDtm.ObjValorPro,
                            GCSTRFMTFECHASIMPLE)
                    txtFechaVencimiento.Content = Format(.ObjFechaVencimientoDtm.ObjValorPro,
                            GCSTRFMTFECHASIMPLE)
                    txtIdMedPag.Content = ClsOrionCop.FstrNombreDatoConstanteOri(
                            EnuGrupoConstantesOriDef.EnuMediosPago, .ObjIdMedioPagoByt.ObjValorPro)
                    txtFormaPago.Content = ClsOrionCop.FstrNombreDatoConstanteOri(
                            EnuGrupoConstantesOriDef.EnuFormaPago, .ObjIdFormaPagoByt.ObjValorPro)
                End With
                SMuestreFechas()
                SMuestreEstado()
                SMuestreEstadoCuenta()
                SHabiliteMenuEFac()
                SEstablezcaDataContext()
            End If
        Else
            SMuestreDatosNuevos()
        End If
        txtRefPago.Content = MobjObjetoWin.ObjReferenciaPago_FacStr.ToString()
        SMuestreUsuario()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando OrElse
            MobjObjetoWin.BlnExiste Then
            SValide()
        End If
        SMuestreTitulo
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If txtIdFactura.Focus Then
                txtIdFactura.SelectAll()
            End If
        End If
        HblnMostrandoDatos = False
    End Sub
    Protected Overrides Sub SRefresqueWin()
        SInicialiceObjeto()
        MyBase.SRefresqueWin()
        SEstablezcaDataContext()
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
        EnuOperacionEnVentana.CenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                If .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                    StcValidValido(EnuValidEntradaDef.enuFechaVenc) = True
                    StcValidValido(EnuValidEntradaDef.enuIdCliente) = True
                    StcValidValido(EnuValidEntradaDef.enuIdPreAgru) = True
                    StcValidValido(EnuValidEntradaDef.enuFechaFact) = True
                    StcValidValido(EnuValidEntradaDef.enuFechaVenc) = True
                    StcValidValido(EnuValidEntradaDef.enuServicio) = True
                    StcValidValido(EnuValidEntradaDef.enuPredio) = True
                    StcValidValido(EnuValidEntradaDef.enuDetalle) = True
                    StcValidValido(EnuValidEntradaDef.enuValor) = True
                    StcValidValido(EnuValidEntradaDef.enuItems) = True
                    StcValidValido(EnuValidEntradaDef.enuMedPag) = True
                    StcValidValido(EnuValidEntradaDef.enuForPag) = True
                Else
                    StcValidValido(EnuValidEntradaDef.enuIdCliente) = .ObjIdCliente_FactDbl.BlnEsValido
                    StcValidValido(EnuValidEntradaDef.enuIdPreAgru) = .ObjIdPredioAgrupador_FacStr.BlnEsValido
                    StcValidValido(EnuValidEntradaDef.enuFechaFact) = .ObjFechaFacturaDtm.BlnEsValido
                    StcValidValido(EnuValidEntradaDef.enuFechaVenc) = .ObjFechaVencimientoDtm.BlnEsValido
                    If Not IsNothing(MobjItemFactActual) Then
                        With MobjItemFactActual
                            StcValidValido(EnuValidEntradaDef.enuServicio) = .ObjIdServicio_ItemFactShr.BlnEsValido
                            StcValidValido(EnuValidEntradaDef.enuPredio) = .ObjIdPredio_ItemFactStr.BlnEsValido
                            StcValidValido(EnuValidEntradaDef.enuDetalle) = .ObjDetalle_ItemFactStr.BlnEsValido
                            StcValidValido(EnuValidEntradaDef.enuValor) = .ObjValor_ItemFactDec.BlnEsValido
                        End With
                    Else
                        StcValidValido(EnuValidEntradaDef.enuServicio) = True
                        StcValidValido(EnuValidEntradaDef.enuPredio) = True
                        StcValidValido(EnuValidEntradaDef.enuDetalle) = True
                        StcValidValido(EnuValidEntradaDef.enuValor) = True
                    End If
                    StcValidValido(EnuValidEntradaDef.enuItems) = (dgrItemsFac.Items.Count > 0)
                    StcValidValido(EnuValidEntradaDef.enuMedPag) = .ObjIdMedioPagoByt.BlnEsValido
                    StcValidValido(EnuValidEntradaDef.enuForPag) = .ObjIdFormaPagoByt.BlnEsValido
                End If
            End With
        End If
        '
        SHabiliteBotonesTlb()
        SHabiliteBotonesWin()
        If FblnEstanTodosBien() Then
            If MenuEstadoItem = EnuEstadoItemDef.enuCreandoItem Then
                HbttAceptar.Style = FindResource("RecBttAceDesha")
            End If
        End If
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjIdCliente_FactDbl.ObjValorPro = txtIdClienteNue.Text
            .ObjFechaFacturaDtm.ObjValorPro = dtpFechaFactura.SelectedDate
            .ObjFechaVencimientoDtm.ObjValorPro = dtpFechaVen.SelectedDate
            .ObjFechaDctoProntoPagoDtm.ObjValorPro = GCDTMFECHANULA
            .ObjIdMedioPagoByt.ObjValorPro = cboIdMedPag.SelectedIndex
            .ObjIdFormaPagoByt.ObjValorPro = cboIdFormaPago.SelectedIndex
            If cboPredioAgrupador.SelectedItem = GCSTRSINPA Then
                .ObjIdPredioAgrupador_FacStr.ObjValorPro = String.Empty
            Else
                .ObjIdPredioAgrupador_FacStr.ObjValorPro = cboPredioAgrupador.SelectedItem
            End If
        End With
        If Not IsNothing(MobjItemFactActual) Then
            With MobjItemFactActual
                .ObjIdPredio_ItemFactStr.ObjValorPro = If(cboPredio.Text = My.Resources.Ninguno,
                        String.Empty, cboPredio.Text)
                .ObjIdServicio_ItemFactShr.ObjValorPro = FshrIdServicioActual()
                .ObjDetalle_ItemFactStr.ObjValorPro = txtDetalle.Text
                .ObjValor_ItemFactDec.ObjValorPro = txtValor.Text
                .ObjFechaGraciaIFDtm.ObjValorPro = dtpFechaVen.SelectedDate
                .ObjFechaVencimientoIFDtm.ObjValorPro = dtpFechaVen.SelectedDate
            End With
        End If
        SValide()
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        HmnuImprimir = FmnuiMenuItem("MnuImprimir", "Im_primir", "RecMnuItemSec")
        MnuEnviarPorCorreo = FmnuiMenuItemPan("MnuEnviarPorCorreo", "_Enviar por eMail", 1,
        "")
        Dim lentPosicion = HmnuAcciones.Items.Count - 1
        Dim lsepSeparad As New Separator
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
                    "_Reprocesar Factura Electrónica", "RecMnuItemSec")
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
        If ClsOrionCop.FblnPuedeCrear(EnuTipoDocOri.EnuFactura, True, lstrMens) Then
            ObjObjetoWin = Nothing
            MblnCreando = True
            SInicialiceObjeto()
            MblnCreando = False
            MstrNumeroFac = ""
            MyBase.SCree()
            SPuebleCombos()
            MobjItemFactActual = Nothing
            MdtbItems.Rows.Clear()
            SPuebleComboPrediosAgru()
            dgrItemsFac.DataContext = Nothing
            SVacieItem()
            MblnCancelaCrear = False
            MobjObjetoWin.ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuManual
            SVisibiliceCtrls(True)
            SEstablezcaFechasFac()
            MenuEstadoItem = EnuEstadoItemDef.enuCreandoItem
            SHabiliteCtlsItems()
            bttEncontrarCliente.Focus()
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Protected Overrides Sub SGuarde()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Mouse.OverrideCursor = Cursors.Wait
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            If Not MobjObjetoWin.FblnFacturaIntegra Then
                Throw New ErrorInesperadoPanLException("Error inesperado en documento de cobro. " &
                            "Reinicie Orión y vuelva a digitarlo!")
            End If
            HbttAceptar.IsEnabled = False
            MblnAplicoAjuste = False
            Dim lobjOrionCop = New ClsOrionCop(GCOBJREGISTRO, False)
            lobjOrionCop.SGenereFacturaManual(MobjObjetoWin, MblnCausoMora, lstrMens)
            GobjPanDat.SConfirmeTransaccion()
            If String.IsNullOrEmpty(lstrMens) Then
                lstrMens = FstrNombreDoc() & " fue creada exitosamente!"
            End If
            GobjPanDat.SControleProcesoObj(False)
            SFinaliceOperacion()
            HbttAceptar.IsEnabled = True
            SCrearClic()
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ConexionBdPanException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                bttEncontrarCliente.Focus()
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub
    Protected Overrides Sub SCancele()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            MblnCancelaCrear = True
            SFinaliceOperacion()
            SRefresqueWin()
        Else
            SCerrarClic()
        End If
    End Sub
    Protected Overrides Function SAnule() As Boolean
        Dim lobjOrionCop = New ClsOrionCop(GCOBJREGISTRO, False)
        Dim larlNrosFacts = lobjOrionCop.FarlFactsPropietarios(MobjObjetoWin)
        Dim lblnAnulo As Boolean
        MstrMensAnulo = String.Empty
        If larlNrosFacts.Count > 1 Then
            Dim lstrPrefAct = MobjObjetoWin.ObjPrefijo_FactStr.ObjValorPro
            Dim lentIdFacAct = MobjObjetoWin.ObjIdFacturaEnt.ObjValorPro
            Dim lobjVlrLlaveAct As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefAct,
                    lentIdFacAct}
            Dim lstrFacts = String.Empty, i = 0
            For Each lstrNroFac As String In larlNrosFacts
                If i = 0 Then
                    lstrFacts &= lstrNroFac
                ElseIf i = larlNrosFacts.Count - 1 Then
                    lstrFacts &= " y " & lstrNroFac & ", "
                Else
                    lstrFacts &= ", " & lstrNroFac
                End If
                i += 1
            Next
            Dim lstrMens = "Serán anuladas las facturas " & lstrFacts & " correspondientes " &
                    " a los propietarios del predio " &
                    MobjObjetoWin.ObjIdPredioAgrupador_FacStr.ToString() & ". Desea continuar?"
            If MsgBox(lstrMens, MsgBoxStyle.YesNo, "Confirmar Anulación") = MsgBoxResult.Yes Then
                lblnAnulo = SAnuleFactsPropietarios(larlNrosFacts)
            End If
            If lblnAnulo Then
                MobjObjetoWin.SAbra(lobjVlrLlaveAct)
                SFinaliceOperacion()
                SRefresqueWin()
            Else
                If Not String.IsNullOrEmpty(MstrMensAnulo) Then
                    SLevanteEveNoti(MstrMensAnulo, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            End If
        Else
            lblnAnulo = MyBase.SAnule()
            If lblnAnulo Then
                If MobjObjetoWin.BlnEstaRegEFac Then
                    SProceseNotaCrAPI()
                End If
                SImprimaNcr()
            End If
        End If
        Return lblnAnulo
    End Function
    Private Function SAnuleFactsPropietarios(aarlFacturasProps As ArrayList) As Boolean
        Dim lstrPref As String, lentIdFact As Integer
        Dim lblnAnulo = False
        Dim lobjValorLlave As Object()
        GobjPanDat.SControleProcesoObj(True)
        GobjPanDat.SInicialiceTransaccion()
        For Each lstrNroFac As String In aarlFacturasProps
            lstrPref = ClsPanorama.FstrPrefijoDcto(lstrNroFac)
            lentIdFact = ClsPanorama.FentIdDcto(lstrNroFac)
            lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFact}
            MobjObjetoWin.SAbra(lobjValorLlave)
            If Not CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuAnular, Boolean) Then
                lblnAnulo = False
                MstrMensAnulo = "El usuario no tiene permisos para esta acción!"
                Exit For
            End If
            lblnAnulo = MobjObjetoWin.FblnEsAnulable
            If Not lblnAnulo Then
                MstrMensAnulo = "La factura número " & lstrNroFac & " no puede ser anulada!"
                Exit For
            End If
        Next
        If lblnAnulo Then
            lblnAnulo = False
            For Each lstrNroFac As String In aarlFacturasProps
                lstrPref = ClsPanorama.FstrPrefijoDcto(lstrNroFac)
                lentIdFact = ClsPanorama.FentIdDcto(lstrNroFac)
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, lstrPref, lentIdFact}
                With MobjObjetoWin
                    .SAbra(lobjValorLlave)
                    .SAnule()
                    lblnAnulo = True
                    If MobjObjetoWin.BlnEstaRegEFac Then
                        SProceseNotaCrAPI()
                    End If
                    SImprimaNcr()
                End With
            Next
        End If
        If lblnAnulo Then
            GobjPanDat.SConfirmeTransaccion()
        Else
            GobjPanDat.SAborteTransaccion()
        End If
        GobjPanDat.SControleProcesoObj(False)
        Return lblnAnulo
    End Function
    Protected Overrides Sub SEstablezcaWinConsultando()
        If Not MblnCancelaCrear AndAlso Not HblnSeEstaCerrando Then
            If GobjParametros.BlnEFacAutorizado Then
                SRegistreFacAPI()
            Else
                If MsgBox("Desea imprimir La Factura generada?", vbYesNo + MsgBoxStyle.Question,
                        "Imprimir Factura?") = vbYes Then
                    SImprima()
                End If
            End If
        End If
        SVisibiliceCtrls(False)
        MyBase.SEstablezcaWinConsultando()
        txtIdFactura.SelectAll()
    End Sub
    Protected Overrides Sub SImprima()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Mouse.OverrideCursor = Cursors.Wait
            If MobjObjetoWin.ObjPrefijo_FactStr.ObjValorPro <> GCSTRPREFPREFACTURA Then
                If MobjObjetoWin.BlnExiste Then
                    SImprimaFac()
                End If
            Else
                lstrMens = "Las Pre-Facturas no se pueden imprimir"
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
            ElseIf Not String.IsNullOrEmpty(lstrMens) Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuInformacion)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub
    Private Sub SImprimaEstadoCta()
        Dim lobjEstCta = MobjObjetoWin.ObjEstadoCuenta
        Dim lstrMens = String.Empty
        If Not IsNothing(lobjEstCta) Then
            SLevanteEveNoti("Imprimiendo", String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Dim lentIdCtaCobroPrimera = lobjEstCta.ObjIdEstadoCuentaEnt.ObjValorPro
            Dim lentIdCtaCobroUltima = lobjEstCta.ObjIdEstadoCuentaEnt.ObjValorPro
            Dim lobjParaCuentaCobro As New ClsParametrosReportesDocs("",
                    lentIdCtaCobroPrimera, lentIdCtaCobroUltima)
            Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
    .ObjParRepDocs = lobjParaCuentaCobro,
    .EnuReporte = EnuReporteDef.enuCtaCobroDet,
    .BlnEstadoDeCuenta = True
    }
            lobjRep.SGenereReporte()
        Else
            lstrMens = "No hay Estado de Cuenta para Imprimir!"
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        Else
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuOk)
        End If
    End Sub
    Private Sub SImprimaNcr()
        Dim lstrMens = "Desea imprimir la Nota Cr generada?"
        If MsgBox(lstrMens, vbYesNo, "Imprimir Documento") = vbYes Then
            Dim lobjNotaCrAnu As ClsNotaCr = MobjObjetoWin.ObjNotaCrAnulo
            If Not IsNothing(lobjNotaCrAnu) Then
                lstrMens = "Imprimiendo"
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Dim lstrPrefNotaCr = lobjNotaCrAnu.ObjPrefijo_NotaCrStr.ObjValorPro
                Dim lentIdNotaPrimera = lobjNotaCrAnu.ObjIdNotaCrEnt.ObjValorPro
                Dim lentIdNotaUltima = lobjNotaCrAnu.ObjIdNotaCrEnt.ObjValorPro
                Dim lobjParaNotaCr As New ClsParametrosReportesDocs(lstrPrefNotaCr,
                        lentIdNotaPrimera, lentIdNotaUltima)
                Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                    .ObjParRepDocs = lobjParaNotaCr,
                    .EnuReporte = EnuReporteDef.enuNotaCr
                }
                lobjRep.SGenereReporte()
            End If
        End If
    End Sub
    Private Sub SImprimaFac()
        Dim lstrMens = String.Empty
        Dim lblnPuede = MobjObjetoWin.BlnEstaRegEFac OrElse Not MobjObjetoWin.BlnEsFacEle
        If lblnPuede Then
            Dim lstrPrefFact = MobjObjetoWin.ObjPrefijo_FactStr.ObjValorPro
            Dim lentIdFacPrimera = MobjObjetoWin.ObjIdFacturaEnt.ObjValorPro
            Dim lentIdFacUltima = MobjObjetoWin.ObjIdFacturaEnt.ObjValorPro
            Dim lobjParaFact As New ClsParametrosReportesDocs(lstrPrefFact, lentIdFacPrimera,
                    lentIdFacUltima)
            Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                    .ObjParRepDocs = lobjParaFact
                }
            If MobjObjetoWin.BlnEsFacEle Then
                lobjRep.EnuReporte = EnuReporteDef.enuFacturaEFac
            Else
                If MobjObjetoWin.ObjIdModoFacturacionByt.ObjValorPro = EnuModoFacturacionDef.EnuImportada Then
                    lobjRep.EnuReporte = EnuReporteDef.enuFactImportada
                Else
                    If GobjParametros.ObjAnoActual.ObjAplicaDsctoPPBln.ObjValorPro Then
                        lobjRep.EnuReporte = EnuReporteDef.enuFacturaDscto
                    Else
                        lobjRep.EnuReporte = EnuReporteDef.enuFactura
                    End If
                End If
            End If
            lobjRep.SGenereReporte()
        Else
            lstrMens = "La Factura no se puede imprimir porque aún no está registrada en API de EFac!"
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
#End Region
#Region "Busqueda"
    Private Sub SBusqueCliente()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
            StrResultadoBusqueda = String.Empty
            SBuscar()
            If MobjObjetoWin.ObjIdCliente_FactDbl.BlnEsValido Then
                cboPredioAgrupador.Focus()
            End If
        End If
    End Sub
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If BlnBusquedaOk Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
                If Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                    txtIdClienteNue.Text = StrResultadoBusqueda
                    SAbraCliente()
                End If
            Else
                If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable Then
                    If BlnBusquedaOk AndAlso StrResutadosBusqueda.Length > 0 Then
                        cboPref.SelectedItem = StrResutadosBusqueda(0)
                        txtIdFactura.Text = StrResutadosBusqueda(1)
                        SAbraFact()
                    End If
                End If
            End If
        End If
        SValide()
    End Sub
    ''' <summary>
    ''' Invalida la funcion "fblnDefinioBusqueda" de la clase base.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
            SDefineBusquedaPredioAgr_Prop()
            SDefineBusquedaPredioAgr_Arren()
            SDefineBusquedaCliente()
        Else
            SDefineNombreCliente()
            SDefinePredioAgr()
        End If
        Return True
    End Function
    Private Sub SDefineBusquedaCliente()
        Dim lstrTabla = ClsCliente.SstrNombreTabla
        Dim lstrCamposMostrar = {ClsIdClienteDbl.SstrNombreCampoBd, ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdClienteDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " And " & lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTabla, lstrCamposMostrar,
lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
    Private Sub SDefineBusquedaPredioAgr_Prop()
        Dim lstrTablaPri = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec = ClsPropietario.SstrNombreTabla
        Dim lstrCamSelTablaPri As String() = {"DISTINCT " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd}
        Dim lstrCampSelTablaSec As String() = {ClsIdCliente_PropDbl.SstrNombreCampoBd,
                ClsNombreCompleto_PropStr.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {OrionP.OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd,
                OrionP.OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {OrionP.OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd,
                OrionP.OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdCliente_PropDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & OrionP.OrionCopL.ClsIdCarpetaShr.SstrNombreCampoBd &
                " = " & GshrIdCarpeta & " AND P." &
                OrionP.OrionCopL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = " &
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
        Dim lstrFiltro As String = "P." & OPT.OrionP.PanL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " & GshrIdCarpeta &
                " AND P." & OPT.OrionP.PanL.ClsIdCentroUtilShr.SstrNombreCampoBd & " = " &
                GshrIdCentroUtil
        lstrFiltro &= " AND " & ClsIdClienteArrendatarioDbl.SstrNombreCampoBd & " > 0 AND " &
                lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador - Arrendatario", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, False)
    End Sub
    Private Sub SDefineNombreCliente()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsFactura.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd,
                            ClsFechaFacturaDtm.SstrNombreCampoBd,
                            ClsPrefijo_FactStr.SstrNombreCampoBd,
                            ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_FactDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCamposRetornar As String() = {ClsPrefijo_FactStr.SstrNombreCampoBd,
                            ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = "S." & OrionP.PanL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " &
                GshrIdCarpeta.ToString & " AND S." &
                OrionP.PanL.ClsIdCentroUtilShr.SstrNombreCampoBd &
                " = " & GshrIdCentroUtil.ToString
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCamposRetornar, lstrFiltro, False)
    End Sub
    Private Sub SDefinePredioAgr()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsFactura.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd,
                ClsFechaFacturaDtm.SstrNombreCampoBd, ClsPrefijo_FactStr.SstrNombreCampoBd,
                ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_FactDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd
        Dim lstrCamposRetornar As String() = {ClsPrefijo_FactStr.SstrNombreCampoBd,
                            ClsIdFacturaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = "S." & OPT.OrionP.PanL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " &
                GshrIdCarpeta.ToString & " AND S." & OPT.OrionP.PanL.ClsIdCentroUtilShr.
                SstrNombreCampoBd & " = " & GshrIdCentroUtil.ToString
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCamposRetornar, lstrFiltro, False)
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SInicialiceFac()
        If MobjObjetoWin IsNot Nothing Then
            If cboPref.SelectedItem Is Nothing Then
                cboPref.SelectedIndex = 0
            End If
            Dim lstrPref As String = cboPref.SelectedItem
            If MobjObjetoWin.ObjPrefijo_FactStr.ToString <> lstrPref Then
                ObjObjetoWin = New ClsFactura(lstrPref)
                MobjObjetoWin = ObjObjetoWin
                If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                    MobjObjetoWin.SVayaAlUltimo()
                End If
                MdtbItems = MobjObjetoWin.FdtbItems
            End If
        End If
    End Sub
    Private Sub SVisibiliceCtrls(ablnVisibleCrear As Boolean)
        Dim lvsbConsulta As Visibility
        Dim lvsbCrea As Visibility
        If ClsPredio.FblnConRefPago Then
            lblRefPago.Visibility = Visibility.Visible
            txtRefPago.Visibility = Visibility.Visible
        Else
            lblRefPago.Visibility = Visibility.Collapsed
            txtRefPago.Visibility = Visibility.Collapsed
        End If
        If ablnVisibleCrear Then
            lvsbConsulta = Visibility.Collapsed
            lvsbCrea = Visibility.Visible
            lblUsuarioGenero.Visibility = Visibility.Hidden
            txtUsuarioGenero.Visibility = Visibility.Hidden
            If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
                dtpFechaFactura.Style = FindResource("RecCtlNoHabilitado")
            End If
            If GobjParametros.BlnEFacAutorizado Then
                rwdGenerales.Height = New GridLength(100)
                lblIdFormaPago.Visibility = Visibility.Visible
                cboIdFormaPago.Visibility = Visibility.Visible
                If MobjObjetoWin.ObjIdFormaPagoByt.ObjValorPro = EnuFormaPago.EnuContado Then
                    lblIdMedPago.Visibility = Visibility.Visible
                    cboIdMedPag.Visibility = Visibility.Visible
                Else
                    lblIdMedPago.Visibility = Visibility.Hidden
                    cboIdMedPag.Visibility = Visibility.Hidden
                End If
            Else
                lblIdFormaPago.Visibility = Visibility.Hidden
                cboIdFormaPago.Visibility = Visibility.Hidden
                lblIdMedPago.Visibility = Visibility.Hidden
                cboIdMedPag.Visibility = Visibility.Hidden
                rwdGenerales.Height = New GridLength(75)
            End If
        Else
            lvsbConsulta = Visibility.Visible
            lvsbCrea = Visibility.Collapsed
            lblUsuarioGenero.Visibility = Visibility.Visible
            txtUsuarioGenero.Visibility = Visibility.Visible
            If GobjParametros.BlnEFacAutorizado Then
                rwdGralCons.Height = New GridLength(100)
                lblFormaPagoCons.Visibility = Visibility.Visible
                txtFormaPago.Visibility = Visibility.Visible
                If MobjObjetoWin.ObjIdFormaPagoByt.ObjValorPro = EnuFormaPago.EnuContado Then
                    lblMedPagoCons.Visibility = Visibility.Visible
                    txtIdMedPag.Visibility = Visibility.Visible
                End If
            Else
                If ClsPredio.FblnConRefPago Then
                    rwdGralCons.Height = New GridLength(100)
                Else
                    rwdGralCons.Height = New GridLength(70)
                End If
                lblFormaPagoCons.Visibility = Visibility.Collapsed
                txtFormaPago.Visibility = Visibility.Collapsed
                lblMedPagoCons.Visibility = Visibility.Collapsed
                txtIdMedPag.Visibility = Visibility.Collapsed
            End If
        End If
        cnvConsulta.Visibility = lvsbConsulta
        cnvNuevo.Visibility = lvsbCrea
    End Sub
    Private Sub SHabiliteCtlsItems()
        Select Case MenuEstadoItem
            Case EnuEstadoItemDef.enuConsultandoItem
                cboPredio.Style = FindResource("RecCtlNoHabilitado")
                cboServicio.Style = FindResource("RecCtlNoHabilitado")
                txtDetalle.Style = FindResource("RecCtlNoHabilitado")
                txtValor.Style = FindResource("RecCtlNoHabilitado")
                bttNuevoItem.IsEnabled = True
                bttAceptarItem.IsEnabled = False
                bttCancelarItem.Visibility = Visibility.Collapsed
                bttEliminarItem.Visibility = Visibility.Visible
                dgrItemsFac.IsEnabled = True
                If dgrItemsFac.Items.Count > 0 Then
                    dgrItemsFac.SelectedIndex = 0
                End If
                bttEliminarItem.IsEnabled = (dgrItemsFac.Items.Count > 0)
            Case EnuEstadoItemDef.enuCreandoItem
                Dim lstrEstilo As String
                If Not IsNothing(MobjCliente) AndAlso MobjCliente.BlnExiste Then
                    lstrEstilo = "RecCtlHabilitado"
                    bttCancelarItem.IsEnabled = True
                Else
                    lstrEstilo = "RecCtlNoHabilitado"
                    bttNuevoItem.IsEnabled = False
                    bttCancelarItem.IsEnabled = False
                End If
                SHabiliteCtrls(lstrEstilo)
                If Not IsNothing(cboPredioAgrupador.SelectedItem) AndAlso
                        cboPredioAgrupador.SelectedItem <> My.Resources.Ninguno AndAlso
                        cboPredioAgrupador.SelectedItem <> GCSTRSINPA Then
                    cboPredio.Style = FindResource("RecCtlHabilitado")
                Else
                    cboPredio.Style = FindResource("RecCtlNoHabilitado")
                    If Not IsNothing(MobjItemFactActual) Then
                        MobjItemFactActual.ObjIdPredio_ItemFactStr.ObjValorPro = String.Empty
                    End If
                End If
                txtValor.Text = Format(0, "c")
                bttNuevoItem.IsEnabled = False
                bttAceptarItem.IsEnabled = False
                bttCancelarItem.Visibility = Visibility.Visible
                bttEliminarItem.Visibility = Visibility.Collapsed
                dgrItemsFac.IsEnabled = False
        End Select
    End Sub
    Private Sub SHabiliteCtrls(astrEstilo As String)
        cboPredio.Style = FindResource(astrEstilo)
        cboServicio.Style = FindResource(astrEstilo)
        txtDetalle.Style = FindResource(astrEstilo)
        txtValor.Style = FindResource(astrEstilo)
        bttNuevoItem.Style = FindResource(astrEstilo)
    End Sub
    Private Sub SHabiliteBotonesWin()
        If MenuEstadoItem = EnuEstadoItemDef.enuConsultandoItem Then
            bttNuevoItem.IsEnabled = FblnEncabezadoOk()
        End If
        If MenuEstadoItem = EnuEstadoItemDef.enuCreandoItem Then
            bttAceptarItem.IsEnabled = FblnItemFacturaOk()
        End If
    End Sub
    Private Sub SHabilteAcciones(ablnHabil As Boolean)
        If ablnHabil Then
            HbttImprimir.Visibility = Visibility.Visible
            HbttCrear.Visibility = Visibility.Visible
            HbttAnular.Visibility = Visibility.Visible
            HmnuImprimir.Visibility = Visibility.Visible
            HmnuCrear.Visibility = Visibility.Visible
            HmnuAnular.Visibility = Visibility.Visible
        Else
            HbttImprimir.Visibility = Visibility.Collapsed
            HbttCrear.Visibility = Visibility.Collapsed
            HbttAnular.Visibility = Visibility.Collapsed
            HmnuImprimir.Visibility = Visibility.Collapsed
            HmnuCrear.Visibility = Visibility.Collapsed
            HmnuAnular.Visibility = Visibility.Collapsed
        End If
    End Sub
    Private Function FblnEncabezadoOk() As Boolean
        Dim lblnEncabOk As Boolean = StcValidValido(EnuValidEntradaDef.enuIdCliente) AndAlso
StcValidValido(EnuValidEntradaDef.enuFechaFact) AndAlso
StcValidValido(EnuValidEntradaDef.enuFechaVenc) AndAlso
StcValidValido(EnuValidEntradaDef.enuIdPreAgru)
        Return lblnEncabOk
    End Function
    Private Function FblnItemFacturaOk() As Boolean
        Dim lblnItemFacOk = False
        With MobjItemFactActual
            If Not IsNothing(MobjItemFactActual) Then
                lblnItemFacOk = .ObjIdServicio_ItemFactShr.BlnEsValido AndAlso
                .ObjIdPredio_ItemFactStr.BlnEsValido AndAlso
                .ObjDetalle_ItemFactStr.BlnEsValido AndAlso
                .ObjValor_ItemFactDec.BlnEsValido
            Else
                lblnItemFacOk = False
            End If
        End With
        Return lblnItemFacOk
    End Function
    Private Sub SAdicioneControlesRestringidos()
        SAdicioneControlRestringido(bttNuevoItem)
        SAdicioneControlRestringido(bttAceptarItem)
        SAdicioneControlRestringido(bttCancelarItem)
        SAdicioneControlRestringido(bttEliminarItem)
        SAdicioneControlRestringido(bttEncontrarCliente)
        SAdicioneControlRestringido(dgrItemsFactura)
        SAdicioneControlRestringido(dgrNovedades)
        SAdicioneControlRestringido(dgrFacturasEstado)
        SAdicioneControlRestringido(dgrItemsEstadoResum)
        SAdicioneControlRestringido(chkVistaRes)
    End Sub
    Private Sub SEstablezcaFechasFac()
        MobjObjetoWin.ObjFechaFacturaDtm.ObjValorPro = Date.Today
        With GobjParametros
            If Not .ObjExigeFechaHoyDocsBln.ObjValorPro Then
                If .ObjAnoActual.StrIdPeriodoActual < ClsOrionCop.FstrPeriodoDeFecha(Date.Today) Then
                    MobjObjetoWin.ObjFechaFacturaDtm.ObjValorPro =
                    .ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                End If
            End If
            Dim lshrPlazo As Short = GobjParametros.ObjPlazoDefectoFacManualShr.ObjValorPro
            MobjObjetoWin.ObjFechaVencimientoDtm.ObjValorPro = Today.AddDays(lshrPlazo)
        End With
        dtpFechaFactura.SelectedDate = MobjObjetoWin.ObjFechaFacturaDtm.ObjValorPro
        dtpFechaVen.SelectedDate = MobjObjetoWin.ObjFechaVencimientoDtm.ObjValorPro
    End Sub
    Private Sub SRegistreFechas()
        MobjObjetoWin.ObjFechaFacturaDtm.ObjValorPro = dtpFechaFactura.SelectedDate
        If MobjObjetoWin.ObjIdFormaPagoByt.ObjValorPro = EnuFormaPago.EnuContado Then
            MobjObjetoWin.ObjFechaVencimientoDtm.ObjValorPro = dtpFechaFactura.SelectedDate
        Else
            MobjObjetoWin.ObjFechaVencimientoDtm.ObjValorPro = dtpFechaVen.SelectedDate
        End If
    End Sub
    Private Sub SMuestreDatosNuevos()
        With MobjObjetoWin
            txtIdClienteNue.Text = .ObjIdCliente_FactDbl.ObjValorPro
            txtNombreClienteNue.Content = .ObjClienteFactura.ObjNombreCompletoStr.ObjValorPro
            dtpFechaFactura.SelectedDate = .ObjFechaFacturaDtm.ObjValorPro
            dtpFechaVen.SelectedDate = .ObjFechaVencimientoDtm.ObjValorPro
            cboIdMedPag.SelectedIndex = .ObjIdMedioPagoByt.ObjValorPro
            cboIdFormaPago.SelectedIndex = .ObjIdFormaPagoByt.ObjValorPro
        End With
        If Not IsNothing(MobjItemFactActual) Then
            With MobjItemFactActual
                cboServicio.SelectedIndex = FentIndiceServicio(.ObjIdServicio_ItemFactShr.ObjValorPro)
                cboServicio.ToolTip = cboServicio.SelectedItem
                If String.IsNullOrEmpty(.ObjIdPredio_ItemFactStr.ObjValorPro) Then
                    cboPredio.SelectedIndex = 0
                Else
                    cboPredio.SelectedItem = .ObjIdPredio_ItemFactStr.ObjValorPro
                End If
                txtDetalle.Text = .ObjDetalle_ItemFactStr.ObjValorPro
                txtValor.Text = Format(.ObjValor_ItemFactDec.ObjValorPro, "c")
                If Not IsNothing(.ObjServicio) Then
                    txtIva.Content = .ObjServicio.ObjTarifaIvaDbl.ToString
                Else
                    txtIva.Content = String.Empty
                End If
            End With
            dgrItemsFac.DataContext = MdtbItems
        Else
            SVacieItem()
        End If
        SCalculeTotales()
        SValide()
    End Sub
    Private Sub SMuestrePropietario()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If FblnEstanTodosBien() Then
                Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
                lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, txtIdCliente.Content})
                If lobjCliente.BlnExiste Then
                    Dim lwinCliente As New WinClientes With {
        .WinPadre = Me,
        .ObjObjetoWin = lobjCliente
    }
                    lwinCliente.ShowDialog()
                End If
            End If
        End If
    End Sub
    Private Sub SEstablezcaDataContext()
        If Not IsNothing(MobjObjetoWin) Then
            Select Case True
                Case tbiFactura.IsSelected
                    dgrItemsFactura.DataContext = MobjObjetoWin.DtbItemsFact
                    SOrdeneDataGrid(dgrItemsFactura, dgrItemsFactura.Columns(0),
                            ClsIdItemFacturaShr.SstrNombreCampoBd,
                            ListSortDirection.Ascending)
                Case tbiMovimiento.IsSelected
                    dgrNovedades.DataContext = MobjObjetoWin.DtbNovedadesFact
                    SOrdeneDataGrid(dgrNovedades, dgrNovedades.Columns(0),
                            ClsIdNovedadShr.SstrNombreCampoBd,
                            ListSortDirection.Ascending)
                Case tbiEstadoCuenta.IsSelected
                    chkVistaRes.IsChecked = False
                    dgrFacturasEstado.Visibility = Visibility.Visible
                    dgrItemsEstadoResum.Visibility = Visibility.Hidden
                    If Not IsNothing(MobjObjetoWin.ObjEstadoCuenta) Then
                        Using ldtbFactEstado As DataTable = MobjObjetoWin.ObjEstadoCuenta.DtbFacturasEstado
                            If Not IsNothing(MobjObjetoWin.ObjEstadoCuenta) Then
                                dgrFacturasEstado.DataContext = ldtbFactEstado
                                SOrdeneDataGrid(dgrFacturasEstado, dgrFacturasEstado.Columns(1),
                                        ClsIdFacturaVivaEnt.SstrNombreCampoBd,
                                        ListSortDirection.Ascending)
                                SMuestreEstadoCuenta()
                            Else
                                dgrFacturasEstado.DataContext = ldtbFactEstado
                                txtFechaEstado.Content = String.Empty
                                txtDeudaCaiptal.Content = Format(0, "c")
                                txtDeudaMora.Content = Format(0, "c")
                                txtTotalDeuda.Content = Format(0, "c")
                            End If
                        End Using
                    End If
            End Select
        End If
    End Sub
    Private Sub SAbraFact()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                If txtIdFactura.Text <> MobjObjetoWin.ObjIdFacturaEnt.ToString() Then
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, cboPref.SelectedItem,
            txtIdFactura.Text}
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
    Private Sub SMuestreFechas()
        txtFechaEstadoFac.Visibility = Visibility.Visible
        Select Case MobjObjetoWin.EnuEstadoFactura
            Case EnuEstadoFacturaDef.EnuAnulada
                lblFechaEstadoFac.Content = My.Resources.FechaAnulacion
                txtFechaEstadoFac.Content = MobjObjetoWin.ObjFechaAnulacion_FactDtm.ToString
            Case EnuEstadoFacturaDef.EnuCancelada
                lblFechaEstadoFac.Content = My.Resources.FechaCancelacion
                txtFechaEstadoFac.Content = MobjObjetoWin.ObjFechaCancelacion_FactDtm.ToString
            Case Else
                If MobjObjetoWin.ObjFechaGraciaDtm.ObjValorPro <> GCDTMFECHANULA Then
                    lblFechaEstadoFac.Content = My.Resources.FechaGracia
                    txtFechaEstadoFac.Content = MobjObjetoWin.ObjFechaGraciaDtm.ToString
                Else
                    lblFechaEstadoFac.Content = My.Resources.NoCausa
                    txtFechaEstadoFac.Visibility = Visibility.Hidden
                End If
        End Select
    End Sub
    Private Sub SMuestreEstado()
        Select Case MobjObjetoWin.EnuEstadoFactura
            Case EnuEstadoFacturaDef.EnuNormal
                txtEstado.Style = FindResource("RecDocNormal")
            Case EnuEstadoFacturaDef.EnuAnulada
                txtEstado.Style = FindResource("RecDocAnulado")
            Case EnuEstadoFacturaDef.EnuPeriodoGracia
                txtEstado.Style = FindResource("RecFacPerGracia")
            Case EnuEstadoFacturaDef.EnuVencida
                txtEstado.Style = FindResource("RecFacVencida")
            Case EnuEstadoFacturaDef.EnuCancelada
                txtEstado.Style = FindResource("RecFacCancelada")
            Case Else
                txtEstado.Style = FindResource("RecDocNoExiste")
        End Select
    End Sub
    Private Sub SMuestreEstadoCuenta()
        If cnvEstadoCuenta.Visibility = Visibility.Visible Then
            Dim lobjEstCta = MobjObjetoWin.ObjEstadoCuenta
            If Not IsNothing(lobjEstCta) Then
                With lobjEstCta
                    txtFechaEstado.Content = Format(.ObjFechaEstadoDtm.ObjValorPro, "dd/MM/yyyy")
                    txtAntPorApl.Content = Format(.ObjAntPorAplDec.ObjValorPro, "c")
                    txtDeudaCaiptal.Content = Format(.ObjDeudaCapitalDec.ObjValorPro, "c")
                    txtDeudaMora.Content = Format(.ObjDeudaIntMoraDec.ObjValorPro, "c")
                    txtTotalDeuda.Content = Format(.DecTotalDeuda, "c")
                End With
            Else
                txtFechaEstado.Content = String.Empty
                txtDeudaCaiptal.Content = String.Empty
                txtDeudaMora.Content = String.Empty
                txtTotalDeuda.Content = String.Empty
            End If
        End If
    End Sub
    Private Sub SMuestreUsuario()
        With MobjObjetoWin
            If MobjObjetoWin.BlnExiste Then
                txtUsuarioGenero.Content = MobjObjetoWin.ObjIdUsuario_FactStr.ObjValorPro
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
    Private Sub SMuestreTitulo()
        Dim lstrTitulo = String.Empty
        If GobjParametros.BlnEFacAutorizado Then
            If MobjObjetoWin.ObjEsPreFacturaBln.ObjValorPro Then
                lstrTitulo = "Pre-Factura Número: "
            Else
                lstrTitulo = "Factura Número: "
            End If
        Else
            If MobjObjetoWin.ObjEsPreFacturaBln.ObjValorPro Then
                lstrTitulo = "Pre-Cuenta de Cobro Número: "
            Else
                lstrTitulo = "Cuenta de Cobro Número: "
            End If
        End If
        If Not String.IsNullOrEmpty(MstrNumeroFac) Then
            lstrTitulo &= MstrNumeroFac & My.Resources.De & txtNombreCliente.Content
        End If
        Title = lstrTitulo
    End Sub
    Private Sub SAbraCliente()
        Dim ldblIdCliente = 0.0, lstrNombreCliente As String
        If IsNumeric(txtIdClienteNue.Text) Then
            ldblIdCliente = CDbl(txtIdClienteNue.Text)
        End If
        MobjObjetoWin.ObjIdCliente_FactDbl.ObjValorPro = txtIdClienteNue.Text
        If Not String.IsNullOrEmpty(txtIdClienteNue.Text) Then
            If Not IsNothing(MobjObjetoWin.ObjClienteFactura) AndAlso
                    MobjObjetoWin.ObjClienteFactura.BlnExiste Then
                MobjCliente = MobjObjetoWin.ObjClienteFactura
                lstrNombreCliente = MobjCliente.ObjNombreCompletoStr.ObjValorPro
                If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                    txtNombreClienteNue.Content = lstrNombreCliente
                    MenuEstadoItem = EnuEstadoItemDef.enuConsultandoItem
                    SHabiliteCtlsItems()
                    SPuebleComboPrediosAgru()
                End If
            Else
                SLevanteEveNoti("La Id. del Cliente ingresada, '" & ldblIdCliente.ToString &
        ", no existe!", "", 0, EnuSeveridadNot.EnuInformacion)
            End If
        End If
    End Sub
    Private Sub SEnvieEmailFra(ByRef astrMens As String)
        If MobjObjetoWin.BlnEstaRegEFac OrElse Not MobjObjetoWin.BlnEsFacEle Then
            If MobjObjetoWin.ObjClienteFactura.ObjRecibeDocsPorEmailBln.ObjValorPro Then
                SEnvieCorreo(Me, EnuTipoCorreoE.EnuFac,
                        MobjObjetoWin.ObjClienteFactura.ObjIdClienteDbl.ObjValorPro,
                        MobjObjetoWin.ObjIdPredioAgrupador_FacStr.ObjValorPro,
                        MobjObjetoWin.StrNumeroFactura)
            Else
                astrMens = "El Cliente de la Factura no tiene habilitado el envio de " &
                        "Documentos por Email!"
            End If
        Else
            astrMens = "La Factura no puede ser publicada debido a que " &
                    "no tiene asignado el CUFE!"
        End If
    End Sub
#End Region
#Region "Manejo Items"
    Private Sub SVacieItem()
        MblnPoblandoCombo = True
        cboServicio.SelectedIndex = 0
        cboPredio.SelectedIndex = 0
        txtDetalle.Text = String.Empty
        txtValor.Text = Format(0, "c")
        MblnPoblandoCombo = False
        MobjItemFactActual = Nothing
    End Sub
    Private Sub SNuevoItem()
        MenuEstadoItem = EnuEstadoItemDef.enuCreandoItem
        MobjItemFactActual = MobjObjetoWin.FobjNuevoItemFactura
        ObjHijoObjWin = MobjItemFactActual
        MobjItemFactActual.ObjIdAno_ServicioItemFactShr.ObjValorPro = 0
        MobjItemFactActual.ObjFechaVencimientoIFDtm.ObjValorPro = dtpFechaVen.SelectedDate
        MobjItemFactActual.ObjFechaGraciaIFDtm.ObjValorPro = dtpFechaVen.SelectedDate
        SHabiliteCtlsItems()
        SMuestreDatos()
        If cboPredio.IsEnabled Then
            cboPredio.Focus()
        Else
            cboServicio.Focus()
            MobjItemFactActual.ObjIdPredio_ItemFactStr.ObjValorPro = String.Empty
        End If
    End Sub
    Private Sub SAcepteItem()
        If MenuEstadoItem = EnuEstadoItemDef.enuCreandoItem AndAlso
                FblnItemFacturaOk() Then
            SRegistre()
            MobjObjetoWin.SAdicioneNuevoItem(MobjItemFactActual)
            MenuEstadoItem = EnuEstadoItemDef.enuConsultandoItem
            SPuebleDtbItems()
            SHabiliteCtlsItems()
            SMuestreDatos()
            bttNuevoItem.Focus()
        End If
    End Sub
    Private Sub SCanceleItem()
        If dgrItemsFac.Items.Count > 0 Then
            dgrItemsFac.SelectedIndex = 0
            MobjItemFactActual = MobjObjetoWin.ColItemsFactura(1)
            SMuestreDatos()
        Else
            SVacieItem()
        End If
        MenuEstadoItem = EnuEstadoItemDef.enuConsultandoItem
        SHabiliteCtlsItems()
        bttNuevoItem.Focus()
    End Sub
    Private Sub SElimineItem()
        If dgrItemsFac.Items.Count > 0 AndAlso Not IsNothing(dgrItemsFac.SelectedItem) Then
            MblnEliminandoItem = True
            Dim lshrIdItemfac As Short = dgrItemsFac.SelectedIndex + 1
            MobjObjetoWin.SElimineItem(lshrIdItemfac)
            SPuebleDtbItems()
            MblnEliminandoItem = False
            If dgrItemsFac.Items.Count > 0 Then
                dgrItemsFac.SelectedIndex = 0
            Else
                MobjItemFactActual = Nothing
                SVacieItem()
            End If
            SMuestreDatos()
            bttNuevoItem.Focus()
        End If
    End Sub
    Private Sub SPuebleDtbItems()
        Dim lobjItemFac As ClsItemFactura = Nothing
        Dim lcolItemsFact As Collection = MobjObjetoWin.ColItemsFactura
        MdtbItems.Rows.Clear()
        For i = 1 To lcolItemsFact.Count
            lobjItemFac = lcolItemsFact(i)
            Dim ldrwNuevoItem As DataRow = MdtbItems.NewRow
            With lobjItemFac
                ldrwNuevoItem(.ObjIdCarpeta_ItemFactShr.StrNombreCampoBD) = .ObjIdCarpeta_ItemFactShr.ObjValorPro
                ldrwNuevoItem(.ObjIdCentroUtil_ItemFactShr.StrNombreCampoBD) = .ObjIdCentroUtil_ItemFactShr.ObjValorPro
                ldrwNuevoItem(ClsIdFacturaEnt.SstrNombreCampoBd) = .ObjIdFactura_ItemFactEnt.ObjValorPro
                ldrwNuevoItem(ClsIdItemFacturaShr.SstrNombreCampoBd) = 0
                ldrwNuevoItem(ClsIdServicio_ItemFactShr.SstrNombreCampoBd) = .ObjIdServicio_ItemFactShr.ObjValorPro
                ldrwNuevoItem(ClsIdPredio_ItemFactStr.SstrNombreCampoBd) = .ObjIdPredio_ItemFactStr.ObjValorPro
                ldrwNuevoItem(ClsTarifaIva_ItemFactDbl.SstrNombreCampoBd) = .ObjServicio.ObjTarifaIvaDbl.ObjValorPro
                ldrwNuevoItem(ClsDetalle_ItemFactStr.SstrNombreCampoBd) = .ObjDetalle_ItemFactStr.ObjValorPro
                ldrwNuevoItem(ClsPrefijo_FactStr.SstrNombreCampoBd) = .ObjPrefijo_ItemFactStr.ToString
                ldrwNuevoItem(ClsValor_ItemFactDec.SstrNombreCampoBd) = .ObjValor_ItemFactDec.ObjValorPro
            End With
            MdtbItems.Rows.Add(ldrwNuevoItem)
        Next
    End Sub
    Private Sub SCalculeTotales()
        Dim ldecTotalFactura = MobjObjetoWin.ObjValor_FactDec.ObjValorPro
        If IsNothing(ldecTotalFactura) Then ldecTotalFactura = 0D
        txtTotalFra.Content = Format(ldecTotalFactura, "c")
    End Sub
#End Region
#Region "eFac"
    Private Sub SProceseNotaCrAPI()
        If Not IsNothing(MobjObjetoWin.ObjNotaCrAnulo) Then
            If MobjObjetoWin.ObjNotaCrAnulo.BlnAfectaFrasRegEFac Then
                Dim lstrMens = String.Empty
                SProceseEFac(lstrMens)
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            End If
        End If
    End Sub
    Private Sub SRegistreFacAPI()
        If MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro < EnuEstadoEDoc.EnuEnviada Then
            Dim lstrMens = String.Empty
            BlnFactAuto = False
            SProceseEFac(lstrMens)
            If MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro >= EnuEstadoEDoc.EnuRegi Then
                If MsgBox("Desea imprimir La Factura generada?", vbYesNo + MsgBoxStyle.Question,
    "Imprimir Factura?") = vbYes Then
                    SImprima()
                End If
            End If
            If Not String.IsNullOrEmpty(lstrMens) Then
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            End If
        End If
    End Sub
    Private Sub SHabiliteMenuEFac()
        If MnuEstadoProvEfac IsNot Nothing Then
            MnuEstadoProvEfac.Visibility = Visibility.Collapsed
            MnuReprocesarEDoc.Visibility = Visibility.Collapsed
        End If
        If GobjParametros.BlnEFacAutorizado AndAlso MobjObjetoWin.BlnExiste AndAlso
        MobjObjetoWin.ObjPrefijo_FactStr.ObjValorPro <> GCSTRPREFPREFACTURA Then
            If Not HblnCargandoForma AndAlso MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                    EnuEstadoEDoc.EnuNoEDoc Then
                Dim lenuEstadoEFac As EnuEstadoEDoc = MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro
                Dim lstrMens = String.Empty, lblnProcesar = True
                MnuEstadoProvEfac.Visibility = Visibility.Visible
                Select Case lenuEstadoEFac
                    Case EnuEstadoEDoc.EnuErrorFtp
                        lstrMens = "El Documento parece tener Problemas en su Información. " &
                    "Ya esta corregido?"
                    Case EnuEstadoEDoc.EnuNoReg
                        MsgBox("El Documento debe ser procesado!", vbOKOnly, "Actualizar Documento!")
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
                        MnuReprocesarEDoc.Visibility = Visibility.Visible
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
#End Region
#Region "Manejo Combos"
    Private Sub SPuebleCombos()
        MblnPoblandoCombo = True
        Dim ldrwConst As DataRow() = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuMediosPago)
        SPuebleComboBox(ldrwConst, cboIdMedPag)
        ldrwConst = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuFormaPago)
        SPuebleComboBox(ldrwConst, cboIdFormaPago)
        SPuebleComboServicio()
        MblnPoblandoCombo = False
    End Sub
    Private Sub SPuebleCboPref()
        MblnPoblandoCombo = True
        Dim ldrwConst As DataRow() = ClsOrionCop.FdrwPrefDoc(EnuTipoDocOri.EnuFactura)
        SPuebleComboBox(ldrwConst, cboPref)
        Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuFacturaVenta)
        If Not cboPref.Items.Contains(lstrPref) Then
            cboPref.Items.Add(lstrPref)
        End If
        MblnPoblandoCombo = False
    End Sub
    Private Sub SPuebleComboServicio()
        Dim lcolServicios = GobjParametros.ColServiciosPer
        MblnPoblandoCombo = True
        cboServicio.Items.Clear()
        cboServicio.Items.Add(My.Resources.Ninguno)
        If Not IsNothing(lcolServicios) AndAlso lcolServicios.Count > 0 Then
            For Each lobjSer As ClsServicio In lcolServicios
                If Not lobjSer.ObjGeneraProgramBln.ObjValorPro AndAlso
                        lobjSer.ObjEstaActivoServicioBln.ObjValorPro Then
                    cboServicio.Items.Add(lobjSer.ObjNombreServicioStr.ObjValorPro)
                End If
            Next
        End If
        cboServicio.SelectedIndex = 0
        MblnPoblandoCombo = False
    End Sub
    Private Sub SPuebleComboPrediosAgru()
        MblnPoblandoCombo = True
        cboPredioAgrupador.Items.Clear()
        cboPredioAgrupador.Items.Add(GCSTRSINPA)
        If Not IsNothing(MobjCliente) Then
            Dim lstrPrediosAgru As String() = MobjCliente.FstrPrediosAgruClienteTodos(False)
            For Each lstrPreAgru As String In lstrPrediosAgru
                cboPredioAgrupador.Items.Add(lstrPreAgru)
            Next
        End If
        cboPredioAgrupador.SelectedIndex = 0
        MblnPoblandoCombo = False
    End Sub
    Private Sub SPuebleComboPredios()
        MblnPoblandoCombo = True
        cboPredio.Items.Clear()
        cboPredio.Items.Add(My.Resources.Ninguno)
        Dim lstrPredioAgrupador As String = cboPredioAgrupador.SelectedItem
        If lstrPredioAgrupador <> GCSTRSINPA Then
            Dim lobjPredioAgr As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
            lobjPredioAgr.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrPredioAgrupador})
            If lobjPredioAgr.BlnExiste Then
                For Each lobjPredio As ClsPredio In lobjPredioAgr.ColPrediosAgrupados
                    cboPredio.Items.Add(lobjPredio.ObjIdPredioStr.ObjValorPro)
                Next
            End If
        End If
        cboPredio.SelectedIndex = 0
        MblnPoblandoCombo = False
    End Sub
    ''' <summary>
    ''' Devuelve el indice en el combobox de un servicio permanente identificado con "ashrIdServicio"
    ''' </summary>
    ''' <param name="ashrIdServicio">Id del Servicio a ubicar en el combobox</param>
    Private Function FentIndiceServicio(ashrIdServicio As Short) As Integer
        Dim lentIndice = 0
        If ashrIdServicio <> 0 Then
            Dim lstrNombreServicio = GobjParametros.FstrNombreServicio(ashrIdServicio)
            For i = 0 To cboServicio.Items.Count - 1
                If cboServicio.Items(i).Equals(lstrNombreServicio) Then
                    lentIndice = i
                    Exit For
                End If
            Next
        End If
        Return lentIndice
    End Function
    Private Function FshrIdServicioActual() As Short
        Dim lshrIdServicio As Short = 0
        If cboServicio.SelectedItem <> My.Resources.Ninguno Then
            lshrIdServicio = GobjParametros.FshrIdServicio(cboServicio.SelectedItem, True)
        End If
        Return lshrIdServicio
    End Function
    Private Sub SDetalle()
        Dim lstrDetalle = cboServicio.SelectedItem.ToString
        If cboPredio.SelectedIndex > 0 Then
            lstrDetalle &= " " & cboPredio.SelectedItem
        End If
        MobjItemFactActual.ObjDetalle_ItemFactStr.ObjValorPro = lstrDetalle
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Select Case lelmElemento.Name
                Case "bttEncontrarCliente"
                    SBusqueCliente()
                Case "bttNuevoItem"
                    SNuevoItem()
                Case "bttAceptarItem"
                    SAcepteItem()
                Case "bttCancelarItem"
                    SCanceleItem()
                Case "bttEliminarItem"
                    SElimineItem()
                Case "bttTercero"
                    SMuestrePropietario()
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
                    SEnvieEmailFra(lstrMens)
                ElseIf lmnuOpcion.Name = "MnuEstadoProvEfac" Then
                    Dim lblnV1 = MobjObjetoWin.FenuVerFacEFac = EnuVerEFac.EnuV1
                    SRefresqueWin()
                    SMuestreEstadoEFac(MobjObjetoWin.ObjCUDocStr.ObjValorPro,
                            EnuTipoDocOri.EnuFactura, MobjObjetoWin.StrNumeroFactura,
                            lblnV1, MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro, lstrMens)
                ElseIf lmnuOpcion.Name = "MnuReprocesarEDoc" Then
                    If MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                            EnuEstadoEDoc.EnuNoEDoc AndAlso
                            MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                             EnuEstadoEDoc.EnuNoReg AndAlso
                             MobjObjetoWin.ObjIdEstadoEDocEnt.ObjValorPro <>
                             EnuEstadoEDoc.EnuInvalida Then
                        If MobjObjetoWin.ObjFechaFacturaDtm.ObjValorPro >= Today.AddDays(-30) Then
                            If MsgBox("Realmente desea reprocesar esta Factura?", vbYesNo,
                                "Reprocesar Factura?") = vbYes Then
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
        Select Case True
            Case TypeOf lelmElemento Is TextBox
                Dim ltxtTextBox As TextBox = lelmElemento
                ltxtTextBox.SelectAll()
            Case TypeOf lelmElemento Is ComboBox
                Dim lcboCombo As ComboBox = lelmElemento
                lcboCombo.SelectedItem = lcboCombo.SelectedItem
            Case TypeOf lelmElemento Is Button
                If lelmElemento.Equals(bttCancelarItem) AndAlso MblnDejoUltimoControl Then
                    If FblnItemFacturaOk() Then
                        bttAceptarItem.Focus()
                        MblnDejoUltimoControl = False
                    End If
                End If
        End Select
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando AndAlso Not HblnMostrandoDatos Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            MblnDejoUltimoControl = False
            If TypeOf lelmElemento Is TextBox Then
                GobjPanDat.SControleProcesoObj(True)
                If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                    Select Case lelmElemento.Name
                        Case "txtIdClienteNue"
                            SAbraCliente()
                        Case "txtDetalle"
                            MobjItemFactActual.ObjDetalle_ItemFactStr.ObjValorPro = txtDetalle.Text
                        Case "txtValor"
                            MobjItemFactActual.ObjValor_ItemFactDec.ObjValorPro = txtValor.Text
                            MblnDejoUltimoControl = True
                    End Select
                    SMuestreDatos()
                End If
                GobjPanDat.SControleProcesoObj(False)
            End If
        End If
    End Sub
    Private Sub DtpFechaVen_LostFocus(sender As Object, e As RoutedEventArgs) Handles _
dtpFechaVen.LostFocus
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If Not HblnSeEstaCerrando AndAlso TypeOf lelmElemento Is DatePicker AndAlso
    EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando AndAlso
    Not HblnMostrandoDatos Then
            MblnDejoUltimoControl = False
            If Not MblnCapturoFechaFac AndAlso lelmElemento.Name = "dtpFechaVen" Then
                SRegistreFechas()
            Else
                MblnCapturoFechaFac = False
            End If
            SMuestreDatos()
        End If
    End Sub
    Private Sub DtpFechaFactura_LostFocus(sender As Object, e As RoutedEventArgs) Handles _
dtpFechaFactura.LostFocus
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If Not HblnSeEstaCerrando AndAlso TypeOf lelmElemento Is DatePicker AndAlso
    EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando AndAlso
    Not HblnMostrandoDatos Then
            If lelmElemento.Name = dtpFechaFactura.Name Then
                MblnDejoUltimoControl = False
                If Not MblnDejoComboBox Then
                    SRegistreFechas()
                    MblnCapturoFechaFac = True
                Else
                    SRegistreFechas()
                    MblnDejoComboBox = False
                End If
                SMuestreDatos()
            End If
        End If
    End Sub
    Private Sub CboPredioAgrupador_LostFocus(sender As Object, e As RoutedEventArgs) _
            Handles cboPredioAgrupador.LostFocus
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando AndAlso
                Not HblnMostrandoDatos Then
            MblnDejoComboBox = True
        End If
    End Sub
    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        If Not MblnPoblandoCombo AndAlso Not HblnMostrandoDatos Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is ComboBox AndAlso Not HblnSeEstaCerrando Then
                Dim lblnNoHayError = False
                If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
                    GobjPanDat.SControleProcesoObj(True)
                    Select Case lelmElemento.Name
                        Case "cboPredioAgrupador"
                            If cboPredioAgrupador.SelectedItem = GCSTRSINPA Then
                                MobjObjetoWin.ObjIdPredioAgrupador_FacStr.ObjValorPro = String.Empty
                            Else
                                MobjObjetoWin.ObjIdPredioAgrupador_FacStr.ObjValorPro =
                                        cboPredioAgrupador.SelectedItem
                            End If
                            SPuebleComboPredios()
                        Case "cboPredio"
                            If MobjItemFactActual IsNot Nothing Then
                                If cboPredio.SelectedItem = My.Resources.Ninguno Then
                                    MobjItemFactActual.ObjIdPredio_ItemFactStr.ObjValorPro = String.Empty
                                Else
                                    MobjItemFactActual.ObjIdPredio_ItemFactStr.ObjValorPro = cboPredio.SelectedItem
                                End If
                                SDetalle()
                            End If
                        Case "cboServicio"
                            If MobjItemFactActual IsNot Nothing Then
                                MobjItemFactActual.ObjIdServicio_ItemFactShr.ObjValorPro = FshrIdServicioActual()
                                If MobjItemFactActual.ObjIdServicio_ItemFactShr.BlnEsValido Then
                                    SDetalle()
                                    txtIva.Content = Format(MobjItemFactActual.ObjServicio.ObjTarifaIvaDbl.ObjValorPro, "p")
                                Else
                                    MobjItemFactActual.ObjDetalle_ItemFactStr.ObjValorPro = String.Empty
                                End If
                            End If
                        Case "cboIdFormaPago"
                            MobjObjetoWin.ObjIdFormaPagoByt.ObjValorPro = cboIdFormaPago.SelectedIndex
                            If MobjObjetoWin.ObjIdFormaPagoByt.ObjValorPro = EnuFormaPago.EnuContado Then
                                lblIdMedPago.Visibility = Visibility.Visible
                                cboIdMedPag.Visibility = Visibility.Visible
                                dtpFechaVen.Style = FindResource("RecCtlNoHabilitado")
                                dtpFechaVen.SelectedDate = dtpFechaFactura.SelectedDate
                                MobjObjetoWin.ObjFechaVencimientoDtm.ObjValorPro = dtpFechaFactura.SelectedDate
                            Else
                                lblIdMedPago.Visibility = Visibility.Hidden
                                cboIdMedPag.Visibility = Visibility.Hidden
                                dtpFechaVen.Style = FindResource("RecCtlHabilitado")
                            End If
                        Case "cboIdMedPag"
                            MobjObjetoWin.ObjIdMedioPagoByt.ObjValorPro = cboIdMedPag.SelectedIndex
                    End Select
                    SMuestreDatos()
                    GobjPanDat.SControleProcesoObj(False)
                Else
                    If lelmElemento.Name = "cboPref" Then
                        SHabilteAcciones(cboPref.SelectedItem <> GCSTRPREFPREFACTURA)
                        SInicialiceFac()
                        SMuestreDatos()
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdFactura.KeyDown
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If e.Key = Key.Return OrElse e.Key = Key.Tab Then
                SAbraFact()
            End If
        End If
    End Sub
    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
    dgrFacturasEstado.MouseRightButtonUp, dgrNovedades.MouseRightButtonUp
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            Dim lstrPrefDocOri As String, lentIdDocOri As Integer
            ldrvFilaActual = ldgrActual.SelectedItem
            If ldgrActual.Name = "dgrFacturasEstado" Then
                If Not IsNothing(ldrvFilaActual) Then
                    lstrPrefDocOri = ldrvFilaActual("PrefijoFactViva")
                    lentIdDocOri = ldrvFilaActual("IdFacturaViva")
                    SAbraFactura(lstrPrefDocOri, lentIdDocOri)
                End If
            ElseIf ldgrActual.Name = "dgrNovedades" Then
                Dim lenuTipoDocOrigen As EnuTipoDocOri =
            ldrvFilaActual(ClsIdTipoDocOrigenByt.SstrNombreCampoBd)
                If lenuTipoDocOrigen <> EnuTipoDocOri.EnuFactura Then
                    lstrPrefDocOri = ldrvFilaActual(ClsPrefijoDocOrigen_NovStr.SstrNombreCampoBd)
                    lentIdDocOri = ldrvFilaActual(ClsIdDocOrigenEnt.SstrNombreCampoBd)
                    Select Case lenuTipoDocOrigen
                        Case EnuTipoDocOri.EnuNotaAjuste
                            SAbraNotaAju(lstrPrefDocOri, lentIdDocOri)
                        Case EnuTipoDocOri.EnuNotaCon
                            SAbraNotaCon(lstrPrefDocOri, lentIdDocOri)
                        Case EnuTipoDocOri.EnuNotaCr
                            SAbraNotaCr(lstrPrefDocOri, lentIdDocOri)
                        Case EnuTipoDocOri.EnuNotaDb
                            SAbraNotaDb(lstrPrefDocOri, lentIdDocOri)
                        Case EnuTipoDocOri.EnuNotaRevCr
                            SAbraNotaRCr(lstrPrefDocOri, lentIdDocOri)
                        Case EnuTipoDocOri.EnuReciboCaja
                            SAbraRecibo(lstrPrefDocOri, lentIdDocOri)
                    End Select
                End If
            End If
        End If
    End Sub
    Private Sub DgrItemsFac_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dgrItemsFac.SelectionChanged
        If MenuEstadoItem = EnuEstadoItemDef.enuConsultandoItem AndAlso Not MblnEliminandoItem Then
            Dim ldrvFilaActual As DataRowView = dgrItemsFac.SelectedItem
            If Not IsNothing(ldrvFilaActual) Then
                MobjItemFactActual = MobjObjetoWin.ColItemsFactura(dgrItemsFac.SelectedIndex + 1)
                SMuestreDatos()
            End If
        End If
    End Sub
    Private Sub OnRatonUp(sender As Object, e As RoutedEventArgs)
        If TypeOf e.Source Is TabItem Then
            SEstablezcaDataContext()
        End If
    End Sub
    Private Sub ChkVistaRes_Click(sender As Object, e As RoutedEventArgs) Handles chkVistaRes.Click
        If chkVistaRes.IsChecked Then
            dgrItemsEstadoResum.Visibility = Visibility.Visible
            dgrFacturasEstado.Visibility = Visibility.Hidden
            If Not IsNothing(MobjObjetoWin.ObjEstadoCuenta) Then
                Using ldtbFactEstado As DataTable = MobjObjetoWin.ObjEstadoCuenta.FdtbFactEstadoResum
                    If Not IsNothing(MobjObjetoWin.ObjEstadoCuenta) Then
                        tbiEstadoCuenta.DataContext = ldtbFactEstado
                    End If
                End Using
            End If
        Else
            SEstablezcaDataContext()
        End If
    End Sub
#End Region
End Class