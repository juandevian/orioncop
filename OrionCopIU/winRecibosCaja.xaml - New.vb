Imports System.Threading.Tasks
Public Class WinRecibosCaja
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuEstadoCapturaDef As Integer
        None
        enuGenerales
        enuCreditos
        enuValores
    End Enum
    Private Enum EnuEstadoDsctoDef As Integer
        None = 0
        enuConsultandoDscto
        enuCreandoDscto
    End Enum
    Private Enum EnuEstadoMedPagoDef As Integer
        None = 0
        enuConsultandoMedPago
        enuCreandoMedPago
    End Enum
    Private Enum EnuValidEntradaDef As Integer
        enuFecha
        enuCliente
        enuPreAgr
        enuServicios
        enuValor
        enuDscto
        enuTipoDscto
        enuFraDscto
        enuItemDscto
        enuTipoMedPago
        enuNroMedPago
        enuNroCtaCon
        enuValorMedPago
        enuMediosPago
        enuComent
        enuCausarInt
        enuAnticipos
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsReciboCaja = Nothing
    Private MobjCliente As ClsCliente = Nothing
    Private MobjMedioPago As ClsMedioPago = Nothing
    Private ReadOnly MobjFactura As New ClsFactura()
    Private MnuEnviarPorCorreo As MenuItem = Nothing
    Private MnuExportarRecs As MenuItemPan = Nothing
    Private MnuMostrarLista As MenuItem = Nothing
    Private MenuEstadoCaptura As EnuEstadoCapturaDef = EnuEstadoCapturaDef.None
    Private MenuEstadoDscto As EnuEstadoDsctoDef = EnuEstadoDsctoDef.None
    Private MenuEstadoMedPago As EnuEstadoMedPagoDef = EnuEstadoMedPagoDef.None
    Private ReadOnly MdtbMediosPago As DataTable = Nothing
    Private MdtbDescuentos As DataTable = Nothing
    Private MblnDejoUltimoControl As Boolean = False
    Private MstrUltimoServiSel As String = String.Empty
    Private MdtmUltimaFechaIng As Date = GCDTMFECHANULA
    ' Indica que se acaba de pasar al modo de crear
    Private MblnPrimeraVezGrales As Boolean = False
    Private MblnPrimeraVezMediosPago As Boolean = False
    Private MblnPoblandoCbo As Boolean = False
    Private MblnCreando As Boolean = False
    Private MblnEliminandoMedPago As Boolean = False
    Private MdecIntPorCausar As Decimal = 0
    Private MdecDeuda As Decimal = 0
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomRecCaj
    Private MblnCausoMora As Boolean = False
    '
    Private WithEvents MobjReportesOrion As ClsRepOrionCop = Nothing
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuRecCaja
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolControlesLlave As New Collection From {
            cboPref,
            txtNroRecCaja
        }
        SAdicioneCtlsRestringidos()
        SCargueForma(EnuElementosAdicionalesDef.enuImprimir, 17,
                lcolControlesLlave, dtpFechaRec, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SCambieNombresDocCobro()
        SMuestreRecibos(Not BlnVentanaAux)
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
            Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuReciboCaja)
            Dim lstrPrefAnt = "*", lblnExistePref = False
            Dim ldrwPrefs As DataRow() = Nothing
            If Not MblnCreando Then
                ldrwPrefs = ClsOrionCop.FdrwPrefDoc(EnuTipoDocOri.EnuReciboCaja)
                For Each ldrwPref As DataRow In ldrwPrefs
                    lstrPrefAnt = ldrwPref("Dato")
                    lblnExistePref = lstrPref = lstrPrefAnt
                    If lblnExistePref Then Exit For
                Next
                If Not lblnExistePref AndAlso lstrPrefAnt <> "*" Then
                    lstrPref = lstrPrefAnt
                End If
            End If
            ObjObjetoWin = New ClsReciboCaja(lstrPref)
            If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                ObjObjetoWin.SVayaAlUltimo()
            Else
                If Not MblnCreando AndAlso ldrwPrefs.Length > 1 Then
                    lstrPref = ldrwPrefs(ldrwPrefs.Length - 2)("Dato")
                    ObjObjetoWin = New ClsReciboCaja(lstrPref)
                    ObjObjetoWin.SVayaAlUltimo()
                End If
            End If
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuServicios) = lblServicios
        StcValidaControl(EnuValidEntradaDef.enuCliente) = lblIdCliente
        StcValidaControl(EnuValidEntradaDef.enuFecha) = lblFechaRec
        StcValidaControl(EnuValidEntradaDef.enuPreAgr) = lblPredioAgru
        StcValidaControl(EnuValidEntradaDef.enuValor) = lblValorRec
        StcValidaControl(EnuValidEntradaDef.enuTipoDscto) = lblTipoDcto
        StcValidaControl(EnuValidEntradaDef.enuFraDscto) = lblIdFactura
        StcValidaControl(EnuValidEntradaDef.enuItemDscto) = lblItemFactura
        StcValidaControl(EnuValidEntradaDef.enuDscto) = lblValorDct
        StcValidaControl(EnuValidEntradaDef.enuTipoMedPago) = lblTipoMedPago
        StcValidaControl(EnuValidEntradaDef.enuNroCtaCon) = lblCtaContIngreso
        StcValidaControl(EnuValidEntradaDef.enuNroMedPago) = lblNumeroMedPago
        StcValidaControl(EnuValidEntradaDef.enuValorMedPago) = lblValorMedPago
        StcValidaControl(EnuValidEntradaDef.enuMediosPago) = lblTotalMedPago
        StcValidaControl(EnuValidEntradaDef.enuComent) = lblComentario
        StcValidaControl(EnuValidEntradaDef.enuCausarInt) = lblIntPorCausar
        StcValidaControl(EnuValidEntradaDef.enuAnticipos) = lblVlrAnt
        SEstablezcaControlesActuales()
        SProceseControlesNuevos(False)
        SPuebleCboCtasIngresos()
        SPuebleComboBoxes()
        SPuebleCboPref()
        If GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
            MdtmUltimaFechaIng = Today
        End If
        SAdicioneControlRestringido(dgrRecsCaja)
        '
        HbttAceptar.TabIndex = 100
        HbttCancelar.TabIndex = 101
    End Sub
    Protected Overrides Sub SMuestreDatos()
        GobjPanDat.SControleProcesoObj(True)
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                SLevanteEveNoti("No hay Recibos de Caja para ser mostrados!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
                cboPref.IsEnabled = False
                txtNroRecCaja.IsEnabled = False
            End If
        End If
        HblnMostrandoDatos = True
        With MobjObjetoWin
            cboPref.SelectedItem = .ObjPrefijo_RecStr.ObjValorPro
            txtNroRecCaja.Text = .ObjIdRecCajaEnt.ToString()
            dtpFechaRec.SelectedDate = .ObjFechaRecDtm.ObjValorPro
            txtIdCliente.Text = .ObjIdCliente_RecDbl.ToString
            txtNombreCliente.Content = .ObjClienteRecibo.ObjNombreCompletoStr.ToString
            txtNombreCliente.ToolTip = txtNombreCliente.Content
            txtComentario.Text = .ObjComentario_RecStr.ObjValorPro
            txtValorR.Content = Format(.ObjValor_RecDec.ObjValorPro, "c")
            txtAnticipo.Content = Format(.ObjValorAnticipoDec.ObjValorPro, "c")
            txtNroNotaRCr.Content = .StrNroNotaRCr
            txtNroNotasCr.Content = .ObjIdNotasCrStr.ToString.Replace(",", "/")
            txtComentarioRec.Text = .ObjComentario_RecStr.ObjValorPro
            txtFechaElaboracon.Content = Format(MobjObjetoWin.ObjFechaCreacionDtm.ObjValorPro, "dd/MM/yyyy hh:mm:ss tt")
            txtValorRec.Text = Format(.ObjValor_RecDec.ObjValorPro, "c")
            txtRecibido.Content = txtValorRec.Text
            txtValorAnt.Content = Format(.ObjValorAnticipoDec.ObjValorPro, "c")
            txtAnticipo.Content = txtValorAnt.Content
            txtIdAnticipo.Content = .ObjIdAnticipo_RecEnt.ToString()
            txtSaldo.Content = Format(MobjObjetoWin.ObjSaldo_RecDec.ObjValorPro, "c")
            txtSaldoPen.Content = txtSaldo.Content
            SMuestreServicios()
            SMuestrePredAgru()
            SMuestreUsuarios()
            SEstablezcaDataContext()
            SMuestreEstado()
            If txtNroRecCaja.Focus() Then
                txtNroRecCaja.SelectAll()
            End If
        End With
        SValide()
        HblnMostrandoDatos = False
        Title = My.Resources.FichaRec
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            Title &= "Nuevo " & My.Resources.De & txtNombreCliente.Content
        Else
            Title &= MobjObjetoWin.StrNumeroRecCaja & My.Resources.De &
                txtNombreCliente.Content
            txtDeudaRecibo.Content = MobjObjetoWin.ObjValorDeudaAlPagoDec.ToString()
        End If
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.cenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntradaDef.enuServicios) = .ObjServicios_RecStr.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuCliente) = .ObjIdCliente_RecDbl.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuFecha) = .ObjFechaRecDtm.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuPreAgr) = .ObjIdPredioAgrupador_RecStr.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuValor) = .ObjValor_RecDec.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuAnticipos) = .ObjValorAnticipoDec.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuComent) = .ObjComentario_RecStr.BlnEsValido
            End With
            If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuCreando Then
                StcValidValido(EnuValidEntradaDef.enuCausarInt) = True
                SValideDscto()
                SValideMedPago()
            End If
        End If
        '
        SHabiliteBotonesTlb()
        SHabiliteBotonesWin()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjComentario_RecStr.ObjValorPro = txtComentario.Text
            .ObjFechaRecDtm.ObjValorPro = dtpFechaRec.SelectedDate
            MdtmUltimaFechaIng = dtpFechaRec.SelectedDate
            .ObjIdCliente_RecDbl.ObjValorPro = txtIdCliente.Text
            If lsbPrediosAgru.SelectedItems.Count > 0 Then
                MobjObjetoWin.ObjIdPredioAgrupador_RecStr.ObjValorPro = FstrPrediosAgrSele()
            End If
            SRegServicios()
        End With
        SValide()
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        Dim lentPosicion = HmnuAcciones.Items.Count - 2
        Dim lsepSeparad As New Separator, lsepSeparadLista As New Separator
        MnuEnviarPorCorreo = FmnuiMenuItemPan("MnuEnviarPorCorreo", "_Enviar por eMail", 1,
            "")
        MnuExportarRecs = FmnuiMenuItemPan("MnuExportarRecs", "Ex_portar Recibos de Caja", 2, "")
        HmnuImprimir = FmnuiMenuItem("MnuImprimir", "Im_primir", "RecMnuItemSec")
        HmnuAcciones.Items.Insert(lentPosicion, MnuExportarRecs)
        If ClsPanorama.FblnEmailsHabilitado Then
            HmnuAcciones.Items.Insert(lentPosicion, MnuEnviarPorCorreo)
        End If
        HmnuAcciones.Items.Insert(lentPosicion, HmnuImprimir)
        HmnuAcciones.Items.Insert(lentPosicion, lsepSeparad)
        '
        MnuMostrarLista = FmnuiMenuItem("MnuMostrarLista", "Mostrar Lista RC", "RecMnuItemSec")
        HmnuHerramientas.Items.Add(lsepSeparadLista)
        HmnuHerramientas.Items.Add(MnuMostrarLista)
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SCree()
        Dim lstrMens = String.Empty
        If ClsOrionCop.FblnPuedeCrear(EnuTipoDocOri.EnuReciboCaja, False, lstrMens) Then
            ObjObjetoWin = Nothing
            MblnCreando = True
            SInicialiceObjeto()
            MblnCreando = False
            If cnvRecsCaja.Visibility = Visibility.Visible Then
                SMuestreRecibos(False)
            End If
            MyBase.SCree()
            MblnPrimeraVezGrales = True
            MenuEstadoCaptura = EnuEstadoCapturaDef.enuGenerales
            MenuEstadoDscto = EnuEstadoDsctoDef.None
            MenuEstadoMedPago = EnuEstadoMedPagoDef.None
            SProceseControlesNuevos(True)
            SHabiliteControlesEstado()
            SEstablezcaCtrlIni()
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Protected Overrides Sub SGuarde()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Static lentUltRC As Integer = 0
        Static lblnGuardando As Boolean = False
        If Not lblnGuardando Then
            lblnGuardando = True
            Try
                GobjPanDat.SControleProcesoObj(True)
                GobjPanDat.SInicialiceTransaccion()
                If MyBase.FblnGravo Then
                    If lentUltRC = 0 Then
                        lentUltRC = MobjObjetoWin.ObjIdRecCajaEnt.ObjValorPro
                    ElseIf lentUltRC <> MobjObjetoWin.ObjIdRecCajaEnt.ObjValorPro Then
                        lentUltRC = MobjObjetoWin.ObjIdRecCajaEnt.ObjValorPro
                    Else
                        Throw New ErrorInesperadoPanLException("Recibo anterior no fue guardado. Avisa a soporte!")
                    End If
                    Dim lstrMensQ = "Desea imprimir El Recibo de Caja?"
                    If MsgBox(lstrMensQ, vbYesNo + MsgBoxStyle.Question, "Imprimir Recibo Caja?") = vbYes Then
                        SImprima()
                    End If
                End If
                lstrMens = FstrNombreDoc() & " fue creado exitosamente!"
                If GobjParametros.BlnEFacAutorizado Then
                    If MblnCausoMora Then
                        SProceseNdb()
                    End If
                    If Not String.IsNullOrEmpty(MobjObjetoWin.ObjIdNotasCrStr.ToString) Then
                        SProceseNsCr()
                    End If
                End If
                SFinaliceOperacion()
                GobjPanDat.SConfirmeTransaccion()
                SCrearClic()
                dtpFechaRec.SelectedDate = MdtmUltimaFechaIng
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
                    GobjPanDat.SControleProcesoObj(False)
                Else
                    GobjPanDat.SAborteTransaccion()
                    GobjPanDat.SControleProcesoObj(False, True)
                    SFinaliceOperacion()
                    MobjObjetoWin.SNormaliceEstado(True)
                    SHabiliteWin(False)
                    SMuestreDatos()
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
            lblnGuardando = False
        End If
    End Sub
    Protected Overrides Sub SCancele()
        If MblnCausoMora Then
            SProceseNdb()
        End If
        MyBase.SCancele()
    End Sub
    Protected Overrides Sub SRefresqueWin()
        MyBase.SRefresqueWin()
        SCambieNombresDocCobro()
        SLevanteEveNoti(String.Empty, String.Empty, 0, EnuSeveridadNot.EnuOk)
        lblIntPorCausar.Visibility = Visibility.Collapsed
        txtIntPorCausar.Visibility = Visibility.Collapsed
        bttCausarIntereses.Visibility = Visibility.Collapsed
    End Sub
    Protected Overrides Function SAnule() As Boolean
        Dim lblnAnulo = MyBase.SAnule()
        If lblnAnulo Then
            Dim lstrMens = String.Empty
            SImprimaNRcr()
            If Not String.IsNullOrEmpty(MobjObjetoWin.ObjIdNotasCrStr.ToString) Then
                If GobjParametros.BlnEFacAutorizado Then
                    SProceseEFac(lstrMens)
                End If
                If lstrMens = String.Empty Then
                    lstrMens = "Fueron anuladas las correspondientes Notas Crédito!"
                End If
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            End If
        End If
        Return lblnAnulo
    End Function
    Protected Overrides Sub SEstablezcaWinConsultando()
        MyBase.SEstablezcaWinConsultando()
        cboNroFactNuevo.Items.Clear()
        cboItemFactNuevo.Items.Clear()
        cboTipoDsctoNuevo.SelectedIndex = 0
        txtValorDctoNuevo.Text = Format(0, "c")
        SProceseControlesNuevos(False)
        cnvDeuda.DataContext = Nothing
        MdtbDescuentos = Nothing
        If Not IsNothing(MdtbMediosPago) Then
            MdtbMediosPago.Rows.Clear()
        End If
        chkAnticipo.IsChecked = False
        grdConsulta.Visibility = Visibility.Visible
        grdPagoNuevo.Visibility = Visibility.Collapsed
        cnvValores.Visibility = Visibility.Visible
        cnvDescuentos.Visibility = Visibility.Collapsed
        SPuebleLsbPredAgru()
        SValide()
    End Sub
    Protected Overrides Sub SImprima()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Mouse.OverrideCursor = Cursors.Wait
            If MobjObjetoWin.BlnExiste Then
                MobjObjetoWin.SRefresqueObj()
                SLevanteEveNoti("Imprimiendo", String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Dim lstrPrefRec = MobjObjetoWin.ObjPrefijo_RecStr.ObjValorPro
                Dim lentIdRecPrimera = MobjObjetoWin.ObjIdRecCajaEnt.ObjValorPro
                Dim lentIdRecUltima = MobjObjetoWin.ObjIdRecCajaEnt.ObjValorPro
                Dim lobjParaFact As New ClsParametrosReportesDocs(lstrPrefRec,
                            lentIdRecPrimera, lentIdRecUltima)
                Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                    .ObjParRepDocs = lobjParaFact,
                    .EnuReporte = EnuReporteDef.enuRecCaja
                    }
                lobjRep.SGenereReporte()
                SLevanteEveNoti(String.Empty, String.Empty, 0, EnuSeveridadNot.EnuOk)
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
            End If
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub
    Private Sub SImprimaNRcr()
        Dim lstrMens = "Desea imprimir la Nota Reversión Cr generada?"
        If MsgBox(lstrMens, vbYesNo, "Imprimir Documento") = vbYes Then
            Dim lobjNotaRCrAnu As ClsNotaReversionCr = MobjObjetoWin.ObjNotaReversionCr
            If Not IsNothing(lobjNotaRCrAnu) Then
                lstrMens = "Imprimiendo"
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Dim lstrPref As String = lobjNotaRCrAnu.ObjPrefijo_NotaReversaCrStr.ObjValorPro
                Dim lentIdNotaPrimera As Integer = lobjNotaRCrAnu.ObjIdNotaReversaCrEnt.ObjValorPro
                Dim lentIdNotaUltima As Integer = lentIdNotaPrimera
                Dim lobjParaNota As New ClsParametrosReportesDocs(lstrPref,
                        lentIdNotaPrimera, lentIdNotaUltima)
                Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
                    .ObjParRepDocs = lobjParaNota,
                    .EnuReporte = EnuReporteDef.enuNotaReverCr
                    }
                lobjRep.SGenereReporte()
            End If
        End If
    End Sub
#End Region
#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        If cnvRecsCaja.Visibility = Visibility.Visible Then
            SMuestreRecibos(False)
        End If
        MyBase.SBuscar()
        If BlnBusquedaOk Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
                If Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                    txtIdCliente.Text = StrResultadoBusqueda
                    SRegCliente()
                End If
            Else
                If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.EnuNavegable Then
                    If BlnBusquedaOk AndAlso StrResutadosBusqueda.Length > 0 Then
                        cboPref.SelectedItem = StrResutadosBusqueda(0)
                        txtNroRecCaja.Text = StrResutadosBusqueda(1)
                        SAbraRecCaja()
                    End If
                End If
            End If
        End If
    End Sub
    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            If txtIdCliente.Focus Then
                If ClsPredio.FblnConRefPago Then
                    SDefineBusquedaRefPago_Prop()
                    SDefineBusquedaRefPago_Arren()
                End If
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
    Private Sub SDefineBusquedaCliente()
        Dim lstrTabla = ClsCliente.SstrNombreTabla
        Dim lstrCamposMostrar = {ClsIdClienteDbl.SstrNombreCampoBd,
                ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdClienteDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " And " &
                lstrCampoBusqueda & " <> ''"
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
    Private Sub SDefineBusquedaRefPago_Prop()
        Dim lstrTablaPri = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec = ClsPropietario.SstrNombreTabla
        Dim lstrCamSelTablaPri As String() = {"DISTINCT " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd, ClsReferenciaPagoStr.SstrNombreCampoBd}
        Dim lstrCampSelTablaSec As String() = {ClsIdCliente_PropDbl.SstrNombreCampoBd,
                ClsNombreCompleto_PropStr.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdPredio_PropStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsReferenciaPagoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdCliente_PropDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & StrCampoCarpeta &
                " = " & GshrIdCarpeta & " AND P." &
                StrCampoCentroUtil & " = " &
                GshrIdCentroUtil & " AND " & "P." & ClsIdPredioStr.SstrNombreCampoBd &
                " = " & ClsIdPredioAgrupadorStr.SstrNombreCampoBd & " AND " &
                lstrCampoBusqueda & " <> ''"
        HwinBusqueda.SDefinaBusqueda("Referencia de Pago - Propietario", lstrTablaPri,
                lstrTablaSec, lstrCamSelTablaPri, lstrCampSelTablaSec, lstrCampRelPri,
                lstrCampRelSec, lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, False)
    End Sub
    Private Sub SDefineBusquedaRefPago_Arren()
        Dim lstrTablaPri As String = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec As String = ClsCliente.SstrNombreTabla
        Dim lstrCamposTabPri = {ClsIdPredioAgrupadorStr.SstrNombreCampoBd,
                ClsReferenciaPagoStr.SstrNombreCampoBd, ClsIdClienteArrendatarioDbl.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteArrendatarioDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsReferenciaPagoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdClienteArrendatarioDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta & " AND P." &
                StrCampoCentroUtil & " = " & GshrIdCentroUtil
        lstrFiltro &= " AND " & ClsIdClienteArrendatarioDbl.SstrNombreCampoBd & " > 0 AND " &
                lstrCampoBusqueda & " <> ''"
        HwinBusqueda.SDefinaBusqueda("Referencia de Pago - Arrendatario", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, True)
    End Sub
    Private Sub SDefineBusquedaPredioAgr_Prop()
        Dim lstrTablaPri = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec = ClsPropietario.SstrNombreTabla
        Dim lstrCamSelTablaPri As String() = {"DISTINCT " &
                ClsIdPredioAgrupadorStr.SstrNombreCampoBd}
        Dim lstrCampSelTablaSec As String() = {ClsIdCliente_PropDbl.SstrNombreCampoBd,
                ClsNombreCompleto_PropStr.SstrNombreCampoBd}
        Dim lstrCampRelPri As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdPredioStr.SstrNombreCampoBd}
        Dim lstrCampRelSec As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdPredio_PropStr.SstrNombreCampoBd}
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
    Private Sub SDefineBusquedaPredioAgr_Arren()
        Dim lstrTablaPri As String = ClsPredio.SstrNombreTabla
        Dim lstrTablaSec As String = ClsCliente.SstrNombreTabla
        Dim lstrCamposTabPri = {ClsIdPredioAgrupadorStr.SstrNombreCampoBd,
                                 ClsIdClienteArrendatarioDbl.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteArrendatarioDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdClienteArrendatarioDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta & " AND P." &
                StrCampoCentroUtil & " = " & GshrIdCentroUtil
        lstrFiltro &= " AND " & ClsIdClienteArrendatarioDbl.SstrNombreCampoBd & " > 0 AND " &
                lstrCampoBusqueda & " <> ''"
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador - Arrendatario", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, True)
    End Sub
    Private Sub SDefineBusquedaPredAgru()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsPredio.SstrNombreTabla
        Dim lstrCamposTabSec = {ClsIdPredioAgrupadorStr.SstrNombreCampoBd,
                                 ClsIdClienteArrendatarioDbl.SstrNombreCampoBd}
        Dim lstrCamposTabPri As String() = {"DISTINCT " & ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdClienteArrendatarioDbl.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda = ClsIdPredioAgrupadorStr.SstrNombreCampoBd
        Dim lstrCampoRetornar = ClsIdClienteArrendatarioDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = "P." & StrCampoCarpeta & " = " &
                GshrIdCarpeta & " AND P." & StrCampoCentroUtil &
                " = " & GshrIdCentroUtil
        lstrFiltro &= " AND " & ClsIdClienteArrendatarioDbl.SstrNombreCampoBd & " > 0 AND " &
                lstrCampoBusqueda & " <> ''"
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador - Arrendatario", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, False)
    End Sub
    Private Sub SDefineNombreCliente()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsReciboCaja.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd,
                "DATE_FORMAT(" & ClsFechaRecDtm.SstrNombreCampoBd & ", '%d/%m/%Y') AS Fecha",
                ClsPrefijo_RecStr.SstrNombreCampoBd, ClsIdRecCajaEnt.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdCliente_RecDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCamposRetornar As String() = {ClsPrefijo_RecStr.SstrNombreCampoBd,
                                            ClsIdRecCajaEnt.SstrNombreCampoBd}
        Dim lstrFiltro As String = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND S." & StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND P." & StrCampoCentroUtil & " = " & GshrIdCentroUtil
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCamposRetornar, lstrFiltro, False)
    End Sub
    Private Sub SDefinePredioAgr()  ' Este
        Dim lstrTablaSec As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaPri As String = ClsReciboCaja.SstrNombreTabla
        Dim lstrCamposTabSec As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabPri As String() = {ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd,
                "DATE_FORMAT(" & ClsFechaRecDtm.SstrNombreCampoBd & ", '%d/%m/%Y') AS Fecha",
                ClsPrefijo_RecStr.SstrNombreCampoBd, ClsIdRecCajaEnt.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {ClsIdCliente_RecDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd
        Dim lstrCamposRetornar As String() = {ClsPrefijo_RecStr.SstrNombreCampoBd,
                                            ClsIdRecCajaEnt.SstrNombreCampoBd}
        Dim lstrFiltro = "S." & StrCampoCarpeta & " = " &
                GshrIdCarpeta.ToString & " AND S." & StrCampoCentroUtil &
                " = " & GshrIdCentroUtil.ToString
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCamposRetornar, lstrFiltro, True)
    End Sub
    Private Sub SBusqueCliente()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            StrResultadoBusqueda = String.Empty
            SBuscar()
        End If
    End Sub
#End Region
#Region "Procedimientos Propios"
#Region "Complementos de sMuestreDatos"
    Private Sub SMuestreMediosPago()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            If Not IsNothing(MobjMedioPago) Then
                txtCuentaIngreso.Content = MobjMedioPago.StrCuentaIngreso
            Else
                MblnPoblandoCbo = True
                cboTipoMedPagoNuevo.SelectedIndex = 0
                MblnPoblandoCbo = False
                txtNroMedPagoNuevo.Text = String.Empty
                cboCuentaIngresoNuevo.SelectedIndex = 0
                txtCuentaIngreso.Content = String.Empty
                txtValorMedPagoNuevo.Text = String.Empty
            End If
        Else
            If Not IsNothing(MobjMedioPago) Then
                txtNomCtaIngreso.Content = MobjMedioPago.StrCuentaIngreso
            End If
        End If
    End Sub
    Private Sub SMuestreEstado()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
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
    Private Sub SMuestreUsuarios()
        With MobjObjetoWin
            If MobjObjetoWin.BlnExiste Then
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                    txtUsuarioGenero.Visibility = Visibility.Visible
                    lblUsuarioGenero.Visibility = Visibility.Visible
                    txtUsuarioGenero.Content = MobjObjetoWin.ObjIdUsuario_RecStr.ObjValorPro
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
            Else
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
                    txtUsuarioGenero.Visibility = Visibility.Collapsed
                    lblUsuarioGenero.Visibility = Visibility.Collapsed
                    lblUsuarioAnulo.Visibility = Visibility.Collapsed
                    txtUsuarioAnulo.Visibility = Visibility.Collapsed
                End If
            End If
        End With
    End Sub
    Private Sub SMuestreRecibos(ablnMostrar As Boolean)
        If MobjObjetoWin IsNot Nothing Then
            Dim ldtbRecsCaja = MobjObjetoWin.FdtbRecsCajaUltimoMes
            If ldtbRecsCaja.Rows.Count = 0 Then
                cnvRecsCaja.Visibility = Visibility.Hidden
                grdRecCaja.Visibility = Visibility.Visible
            Else
                If ablnMostrar Then
                    cnvRecsCaja.Visibility = Visibility.Visible
                    grdRecCaja.Visibility = Visibility.Hidden
                    dgrRecsCaja.DataContext = ldtbRecsCaja
                    HbttAnular.IsEnabled = False
                    HbttAlPrimero.IsEnabled = False
                    HbttAlAnterior.IsEnabled = False
                    HbttAlSiguiente.IsEnabled = False
                    HbttAlUltimo.IsEnabled = False
                    HbttImprimir.IsEnabled = False
                    HmnuAnular.IsEnabled = False
                    HmnuAlPrimero.IsEnabled = False
                    HmnuAlSiguiente.IsEnabled = False
                    HmnuAlAnterior.IsEnabled = False
                    HmnuAlUltimo.IsEnabled = False
                    HmnuImprimir.IsEnabled = False
                    SLevanteEveNoti("Doble clic o clic contratio abre el recibo de caja seleccionado!",
                            String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Else
                    cnvRecsCaja.Visibility = Visibility.Hidden
                    grdRecCaja.Visibility = Visibility.Visible
                    HbttAnular.IsEnabled = True
                    HbttAlPrimero.IsEnabled = True
                    HbttAlAnterior.IsEnabled = True
                    HbttAlSiguiente.IsEnabled = True
                    HbttAlUltimo.IsEnabled = True
                    HbttImprimir.IsEnabled = True
                    HmnuAnular.IsEnabled = True
                    If Not BlnVentanaAux Then
                        HmnuAlPrimero.IsEnabled = True
                        HmnuAlSiguiente.IsEnabled = True
                        HmnuAlAnterior.IsEnabled = True
                        HmnuAlUltimo.IsEnabled = True
                        HmnuImprimir.IsEnabled = True
                    End If
                    SLevanteEveNoti(String.Empty, String.Empty, 0, EnuSeveridadNot.EnuOk)
                End If
            End If
        End If
    End Sub
#End Region
#Region "Manejo Controles"
    Private Sub SAdicioneCtlsRestringidos()
        'Generales
        SAdicioneControlRestringido(tbcRecibosCaja)
        SAdicioneControlRestringido(lsbPrediosAgru)
        SAdicioneControlRestringido(lsbServicios)
        ' Items
        SAdicioneControlRestringido(dgrItems)
        ' Novedades
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
        SAdicioneControlRestringido(bttSiguiente)
        SAdicioneControlRestringido(bttAnterior)
        ' Medios de Pago
        SAdicioneControlRestringido(dgrMediosPago)
        SAdicioneControlRestringido(bttNuevoMedPago)
        SAdicioneControlRestringido(bttAceptarMedPago)
        SAdicioneControlRestringido(bttCancelarMedPago)
        SAdicioneControlRestringido(bttEliminarMedPago)
        SAdicioneControlRestringido(bttEncontrarCliente)
    End Sub
    Private Sub SEstablezcaControlesActuales()
        Dim lstyEstiloNoHabilitado As Style = FindResource("RecCtlNoHabilitado")
        Dim lobjDocRec As ClsDocumento =
                GobjParametros.ObjDocumento(EnuIdDocumentoDef.EnuReciboCaja)
        cnvDeuda.Visibility = Visibility.Collapsed
        ' Descuentos y Retenciones
        txtNroFactura.Style = lstyEstiloNoHabilitado
        txtItemFactura.Style = lstyEstiloNoHabilitado
        txtTipoDcto.Style = lstyEstiloNoHabilitado
        txtValorDcto.Style = lstyEstiloNoHabilitado
        cnvDescuentos.Visibility = Visibility.Collapsed
        ' Medios de Pago
        grdPagoNuevo.Visibility = Visibility.Collapsed
        ' Generales
        bttAnterior.Visibility = Visibility.Collapsed
        bttSiguiente.Visibility = Visibility.Collapsed
    End Sub
    Private Sub SProceseControlesNuevos(ablnHabilite As Boolean)
        Dim lvisVisibilidadNuevos As Visibility
        Dim lvisVisibilidadActual As Visibility
        If ablnHabilite Then
            MblnPoblandoCbo = True
            cboPref.IsEnabled = False
            txtNroRecCaja.IsEnabled = False
            lsbServicios.Items.Clear()
            lsbPrediosAgru.Items.Clear()
            txtDeudaRecibo.Content = Format(0, "c")
            lblIntPorCausar.Visibility = Visibility.Collapsed
            txtIntPorCausar.Visibility = Visibility.Collapsed
            bttCausarIntereses.Visibility = Visibility.Collapsed
            lvisVisibilidadNuevos = Visibility.Visible
            lvisVisibilidadActual = Visibility.Collapsed
            MblnPoblandoCbo = False
        Else
            lvisVisibilidadNuevos = Visibility.Collapsed
            lvisVisibilidadActual = Visibility.Visible
        End If
        ' Consulta
        grdConsulta.Visibility = lvisVisibilidadActual
        grbDocs.Visibility = lvisVisibilidadActual
        ' 
        lblValor.Visibility = lvisVisibilidadActual
        txtValorR.Visibility = lvisVisibilidadActual
        bttEncontrarCliente.Visibility = lvisVisibilidadNuevos
        lblEstado.Visibility = lvisVisibilidadActual
        txtEstado.Visibility = lvisVisibilidadActual
        chkAnticipo.Visibility = lvisVisibilidadNuevos
        ' Captura
        grdPagoNuevo.Visibility = lvisVisibilidadNuevos
        cnvDeuda.Visibility = lvisVisibilidadNuevos
        cnvDescuentos.Visibility = Visibility.Collapsed
        cnvMediosPagoNuevo.Visibility = Visibility.Collapsed
        cnvDetallePago.Visibility = Visibility.Collapsed
        grsPago.Visibility = Visibility.Collapsed
        bttSiguiente.Visibility = lvisVisibilidadNuevos
        bttAnterior.Visibility = lvisVisibilidadNuevos
    End Sub
    ''' <summary>
    ''' Establece el estado de los controles (habilitado o deshabilitado) según el estado de captura
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SHabiliteControlesEstado()
        grdConsulta.Visibility = Visibility.Collapsed
        grdPagoNuevo.Visibility = Visibility.Visible
        cnvDescuentos.Visibility = Visibility.Collapsed
        cnvDeuda.Visibility = Visibility.Collapsed
        cnvMediosPagoNuevo.Visibility = Visibility.Collapsed
        cnvDetallePago.Visibility = Visibility.Collapsed
        grsPago.Visibility = Visibility.Collapsed
        If MenuEstadoCaptura = EnuEstadoCapturaDef.enuGenerales Then
            SHabiliteCtlsGrales(True)
        Else
            SHabiliteCtlsGrales(False)
        End If
        Select Case MenuEstadoCaptura
            Case EnuEstadoCapturaDef.enuGenerales
                MblnPrimeraVezGrales = True
                cnvDeuda.Visibility = Visibility.Visible
                bttAnterior.IsEnabled = False
                bttSiguiente.IsEnabled = FblnGeneralesOk()
            Case EnuEstadoCapturaDef.enuCreditos
                cnvDescuentos.Visibility = Visibility.Visible
                SHabiliteCtlsDesctos()
                bttAnterior.IsEnabled = True
                bttSiguiente.IsEnabled = True
                SValideDscto()
                bttNuevoDscto.Focus()
            Case EnuEstadoCapturaDef.enuValores
                cnvDetallePago.Visibility = Visibility.Visible
                cnvMediosPagoNuevo.Visibility = Visibility.Visible
                grsPago.Visibility = Visibility.Visible
                bttAnterior.IsEnabled = True
                bttSiguiente.IsEnabled = False
                cnvMediosPagoNuevo.DataContext = MobjObjetoWin.DtbMediosPago
                SHabiliteCtlsValores()
                SHabiliteCtlsMedPago()
        End Select
    End Sub
    Private Sub SHabiliteCtlsGrales(ablnHabilite As Boolean)
        Dim lstyCtlsDscto As Style = FindResource("RecCtlNoHabilitado")
        If ablnHabilite Then
            lstyCtlsDscto = FindResource("RecCtlHabilitado")
        End If
        If Not GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro Then
            dtpFechaRec.Style = lstyCtlsDscto
        End If
        bttEncontrarCliente.IsEnabled = ablnHabilite
        txtIdCliente.Style = lstyCtlsDscto
        lsbPrediosAgru.Style = lstyCtlsDscto
        lsbServicios.Style = lstyCtlsDscto
        If lsbPrediosAgru.SelectedItems.Count = 0 OrElse
                MdecIntPorCausar > 0 Then
            If ablnHabilite Then
                SHabiliteControl(chkAnticipo, False)
                chkAnticipo.IsChecked = False
            End If
        Else
            chkAnticipo.Style = lstyCtlsDscto
            lsbServicios.Style = lstyCtlsDscto
        End If
    End Sub
    Private Sub SHabiliteCtlsDesctos()
        txtNroFactura.Visibility = Visibility.Collapsed
        txtItemFactura.Visibility = Visibility.Collapsed
        txtTipoDcto.Visibility = Visibility.Collapsed
        txtValorDcto.Visibility = Visibility.Collapsed
        cboNroFactNuevo.Visibility = Visibility.Collapsed
        cboItemFactNuevo.Visibility = Visibility.Collapsed
        cboTipoDsctoNuevo.Visibility = Visibility.Collapsed
        txtValorDctoNuevo.Visibility = Visibility.Collapsed
        Select Case MenuEstadoDscto
            Case EnuEstadoDsctoDef.enuConsultandoDscto
                txtNroFactura.Visibility = Visibility.Visible
                txtItemFactura.Visibility = Visibility.Visible
                txtTipoDcto.Visibility = Visibility.Visible
                txtValorDcto.Visibility = Visibility.Visible
                bttNuevoDscto.IsEnabled = True
                bttAceptarDscto.IsEnabled = False
                bttCancelarDscto.Visibility = Visibility.Collapsed
                bttEliminarDscto.Visibility = Visibility.Visible
                If dgrDescuentos.Items.Count > 0 Then
                    dgrDescuentos.SelectedIndex = 0
                End If
                bttEliminarDscto.IsEnabled = (dgrDescuentos.Items.Count > 0)
            Case EnuEstadoDsctoDef.enuCreandoDscto
                cboNroFactNuevo.Visibility = Visibility.Visible
                cboItemFactNuevo.Visibility = Visibility.Visible
                cboTipoDsctoNuevo.Visibility = Visibility.Visible
                txtValorDctoNuevo.Visibility = Visibility.Visible
                txtValorDctoNuevo.Text = Format(0, "c")
                bttNuevoDscto.IsEnabled = False
                bttAceptarDscto.IsEnabled = False
                bttCancelarDscto.Visibility = Visibility.Visible
                bttEliminarDscto.Visibility = Visibility.Collapsed
                bttCancelarDscto.IsEnabled = True
        End Select
    End Sub
    Private Sub SHabiliteCtlsValores()
        cnvDetallePago.Visibility = Visibility.Visible
        SHabiliteControl(txtValorRec, True)
        SHabiliteControl(txtComentario, True)
        txtValorRec.Focus()
    End Sub
    Private Sub SInicialiceCtlsMediosPago(ablnHabilite As Boolean)
        SHabiliteControl(cboTipoMedPagoNuevo, ablnHabilite)
        SHabiliteControl(txtNroMedPagoNuevo, ablnHabilite)
        SHabiliteControl(cboCuentaIngresoNuevo, ablnHabilite)
        SHabiliteControl(txtValorMedPagoNuevo, ablnHabilite)
    End Sub
    Private Sub SHabiliteCtlsMedPago()
        cboTipoMedPagoNuevo.Visibility = Visibility.Collapsed
        txtTipoMedPagoCap.Visibility = Visibility.Collapsed
        cboCuentaIngresoNuevo.Visibility = Visibility.Collapsed
        txtCuentaIngreso.Visibility = Visibility.Collapsed
        txtNroMedPagoNuevo.Visibility = Visibility.Collapsed
        txtNroMedPagoCap.Visibility = Visibility.Collapsed
        txtValorMedPagoNuevo.Visibility = Visibility.Collapsed
        txtValorMedPagoCap.Visibility = Visibility.Collapsed
        Select Case MenuEstadoMedPago
            Case EnuEstadoMedPagoDef.enuConsultandoMedPago
                txtTipoMedPagoCap.Visibility = Visibility.Visible
                txtCuentaIngreso.Visibility = Visibility.Visible
                txtNroMedPagoCap.Visibility = Visibility.Visible
                txtValorMedPagoCap.Visibility = Visibility.Visible
                bttNuevoMedPago.IsEnabled = True
                bttAceptarMedPago.IsEnabled = False
                bttCancelarMedPago.Visibility = Visibility.Collapsed
                bttEliminarMedPago.Visibility = Visibility.Visible
                If dgrMediosPagoNuevo.Items.Count > 0 Then
                    dgrMediosPagoNuevo.SelectedIndex = 0
                End If
                bttEliminarMedPago.IsEnabled = (dgrMediosPagoNuevo.Items.Count > 0)
            Case EnuEstadoMedPagoDef.enuCreandoMedPago, EnuEstadoMedPagoDef.None
                cboTipoMedPagoNuevo.Visibility = Visibility.Visible
                cboCuentaIngresoNuevo.Visibility = Visibility.Visible
                txtNroMedPagoNuevo.Visibility = Visibility.Visible
                txtValorMedPagoNuevo.Visibility = Visibility.Visible
                Dim lblnHabilite As Boolean = MobjObjetoWin.ObjValor_RecDec.BlnEsValido AndAlso
                                        MenuEstadoMedPago = EnuEstadoMedPagoDef.enuCreandoMedPago
                SInicialiceCtlsMediosPago(lblnHabilite)
                bttNuevoMedPago.IsEnabled = False
                bttAceptarMedPago.IsEnabled = False
                bttCancelarMedPago.Visibility = Visibility.Visible
                bttEliminarMedPago.Visibility = Visibility.Collapsed
                bttCancelarMedPago.IsEnabled = True
        End Select
    End Sub
    Private Sub SHabiliteControl(actlControl As Control, ablnHabilite As Boolean)
        If ablnHabilite Then
            actlControl.Style = FindResource("RecCtlHabilitado")
        Else
            actlControl.Style = FindResource("RecCtlNoHabilitado")
        End If
    End Sub
    Private Sub SHabiliteBotonesWin()
        If MenuEstadoDscto = EnuEstadoDsctoDef.enuCreandoDscto Then
            SValideDscto()
            bttAceptarDscto.IsEnabled = StcValidValido(EnuValidEntradaDef.enuDscto)
        End If
        If MenuEstadoMedPago = EnuEstadoMedPagoDef.enuCreandoMedPago Then
            bttAceptarMedPago.IsEnabled = FblnEsValidoMedPago()
        End If
    End Sub
    Private Sub TbcRecibosCaja_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles tbcRecibosCaja.SelectionChanged
        If Not IsNothing(MobjObjetoWin) Then
            If TypeOf e.Source Is TabControl Then
                SEstablezcaDataContext()
            End If
        End If
    End Sub
    Private Sub SCambieNombresDocCobro()
        If GobjParametros.BlnEFacAutorizado Then
            dgtNroFac.Header = "Nro. Factura"
            dgtNroFacNov.header = "Nro. Factura"
            dgtNroFacDeuda.header = "Nro. Factura"
            dgtFechaFac.Header = "Fecha Fac."
            dgtValor.header = "Valor Factura"
            lblIdFactura.Content = "Número Factura"
            lblItemFactura.Content = "Item Factura"
            dgtNroFacDcto.Header = "Nro. Factura"
            dgtItemFacDcto.Header = "Item Factura"
        Else
            dgtNroFac.Header = "Nro. C.Cobro"
            dgtNroFacNov.header = "Nro. C.Cobro"
            dgtNroFacDeuda.Header = "Nro. Cuenta Cobro"
            dgtFechaFac.Header = "Fecha C.Cobro"
            dgtValor.header = "Valor Cta.Cobro"
            lblIdFactura.Content = "Número Cuenta Cobro"
            lblItemFactura.Content = "Item Cuenta Cobro"
            dgtNroFacDcto.Header = "Nro. Cuenta Cobro"
            dgtItemFacDcto.Header = "Item Cuenta Cobro"
        End If
    End Sub
#End Region
#Region "Descuentos"
    Private Sub SCargueDescuentos()
        txtValorDctoNuevo.Text = Format(0, "c")
        cboTipoDsctoNuevo.SelectedIndex = 0
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando AndAlso
                MobjObjetoWin.ObjServicios_RecStr.BlnEsValido Then
            SMuestreDescuentos()
        End If
    End Sub
    Private Sub SMuestreDescuentos()
        Dim ldtbDsctos = MobjObjetoWin.DtbDescuentos
        cnvDescuentos.DataContext = ldtbDsctos
        If dgrDescuentos.Items.Count > 0 Then dgrDescuentos.SelectedIndex = 0
        If Not IsNothing(ldtbDsctos) AndAlso ldtbDsctos.Rows.Count > 0 Then
            SMuestreComentario()
        End If
        If MenuEstadoCaptura = EnuEstadoCapturaDef.enuCreditos Then
            SMuestreSaldoAPagar()
        End If
    End Sub
    Private Sub SNuevoDscto()
        If FblnTienePermisoParaDscto() Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
                MenuEstadoDscto = EnuEstadoDsctoDef.enuCreandoDscto
                txtDeudaCap.Content = Format(0, "c")
                txtDeudaMora.Content = Format(0, "c")
                SHabiliteCtlsDesctos()
                cboNroFactNuevo.SelectedIndex = 0
                cboItemFactNuevo.SelectedIndex = 0
                cboTipoDsctoNuevo.SelectedIndex = 0
                txtValorDctoNuevo.Text = Format(0, "c")
                SValideTipoDscto()
                SValideFraDscto()
                SValideItemDscto()
                SValideDscto()
                cboTipoDsctoNuevo.Focus()
            End If
        Else
            Dim lstrMens = "No tiene Permiso para hacer Descuentos!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SAcepteDscto()
        SRegDscto()
        If MenuEstadoDscto = EnuEstadoDsctoDef.enuCreandoDscto AndAlso
                            StcValidValido(EnuValidEntradaDef.enuDscto) Then
            Dim lstrItemFac As String = cboItemFactNuevo.SelectedItem.ToString
            MobjObjetoWin.SAdicionesDscto(cboNroFactNuevo.SelectedItem, lstrItemFac,
                        cboTipoDsctoNuevo.SelectedIndex, txtValorDctoNuevo.Text)
            SMuestreSaldoAPagar()
            MenuEstadoDscto = EnuEstadoDsctoDef.enuConsultandoDscto
            SHabiliteCtlsDesctos()
            bttNuevoDscto.Focus()
        End If
    End Sub
    Private Sub SCanceleDscto()
        If dgrDescuentos.Items.Count > 0 Then
            dgrDescuentos.SelectedIndex = 0
        Else
            cboNroFactNuevo.SelectedIndex = 0
            cboTipoDsctoNuevo.SelectedIndex = 0
            txtValorDctoNuevo.Text = Format(0, "c")
            txtDeudaCap.Content = Format(0, "c")
            txtDeudaMora.Content = Format(0, "c")
        End If
        StcValidValido(EnuValidEntradaDef.enuTipoDscto) = True
        StcValidValido(EnuValidEntradaDef.enuFraDscto) = True
        StcValidValido(EnuValidEntradaDef.enuItemDscto) = True
        StcValidValido(EnuValidEntradaDef.enuDscto) = True
        MenuEstadoDscto = EnuEstadoDsctoDef.enuConsultandoDscto
        SValideDscto()
        SHabiliteCtlsDesctos()
    End Sub
    Private Sub SElimineDscto()
        If dgrDescuentos.Items.Count > 0 AndAlso Not IsNothing(dgrDescuentos.SelectedItem) Then
            Dim ldrvItemDscto As DataRowView = dgrDescuentos.SelectedItem
            Dim lentOrinal As Integer = ldrvItemDscto("Ordinal")
            MobjObjetoWin.SElimineDscto(lentOrinal)
            SCanceleDscto()
            SMuestreSaldoAPagar()
        End If
    End Sub
    Private Function FblnTienePermisoParaDscto() As Boolean
        Dim lblnTienePermisoDscto = FblnHabilitarMenuPan(HenuIdVentana, 3)
        If Not lblnTienePermisoDscto Then
            Dim lwinAutoDscto As New WinAutorizaDscto
            lwinAutoDscto.ShowDialog()
            lblnTienePermisoDscto = GblnOK
        End If
        Return lblnTienePermisoDscto
    End Function
    Private Sub SMuestreComentario()
        Dim lstrComentario = "", lstrDscto As String
        Dim lenuTipodscto As EnuTipoDescuentoDef
        Dim lblnEfectuoDscto As Boolean() = {False, False, False, False}
        Dim ldtbDsctos = MobjObjetoWin.DtbDescuentos
        For Each ldrwDscto As DataRow In ldtbDsctos.Rows
            lstrDscto = String.Empty
            lenuTipodscto = ClsPanorama.FobjValorCampo(ldrwDscto("IdTipoDcto"), EnuTipoValor.enuByte)
            If lenuTipodscto = EnuTipoDescuentoDef.EnuReteFuente Then
                If Not lblnEfectuoDscto(0) Then
                    lstrDscto = "Reteción en la fuente"
                    lblnEfectuoDscto(0) = True
                End If
            ElseIf lenuTipodscto = EnuTipoDescuentoDef.EnuReteIva Then
                If Not lblnEfectuoDscto(1) Then
                    lstrDscto = "Retención del IVA"
                    lblnEfectuoDscto(1) = True
                End If
            ElseIf lenuTipodscto = EnuTipoDescuentoDef.EnuReteIca Then
                If Not lblnEfectuoDscto(2) Then
                    lstrDscto = "Retención de Industria y Comercio"
                    lblnEfectuoDscto(2) = True
                End If
            ElseIf lenuTipodscto = EnuTipoDescuentoDef.EnuDsctoPP Then
                If Not lblnEfectuoDscto(3) Then
                    lstrDscto = "descuento por Pronto Pago"
                    lblnEfectuoDscto(3) = True
                End If
            End If
            If Not String.IsNullOrEmpty(lstrDscto) Then
                If String.IsNullOrEmpty(lstrComentario) Then
                    lstrComentario = "Se efectuó " & lstrDscto
                Else
                    lstrComentario &= ", " & lstrDscto
                End If
            End If
        Next
        If Not String.IsNullOrEmpty(lstrComentario) AndAlso String.IsNullOrEmpty(txtComentario.Text) Then
            txtComentario.Text = lstrComentario
        End If
    End Sub
#End Region
#Region "Medios de Pago"
    Private Sub SDetermineMedPagoInicial()
        If MblnPrimeraVezMediosPago Then
            If dgrMediosPago.Items.Count > 0 Then
                Do While dgrMediosPagoNuevo.Items.Count > 0
                    dgrMediosPagoNuevo.SelectedIndex = 0
                    SElimineMedPago()
                Loop
            End If
            MblnPrimeraVezMediosPago = False
            SMedioPagoPorDefecto()
        End If
        SHabiliteCtlsMedPago()
    End Sub
    Private Sub SMedioPagoPorDefecto()
        SCreeNuevoMedioPago()
        With MobjMedioPago
            Dim lenuMedioPago As EnuTipoMedioPagoDef = FenuTipoMedPagoDefecto()
            cboTipoMedPagoNuevo.SelectedIndex = lenuMedioPago
            .ObjIdTipoMedPago_MedPagoByt.ObjValorPro = lenuMedioPago
            If .ObjIdTipoMedPago_MedPagoByt.ObjValorPro = EnuTipoMedioPagoDef.EnuEfectivo Then
                cboCuentaIngresoNuevo.SelectedIndex = 1
                .ObjIdCtaContabIngresoStr.ObjValorPro = GobjParametros.ObjIdCtaCajaStr.ObjValorPro
                txtCuentaIngreso.Content = .StrCuentaIngreso
            End If
            txtValorMedPagoNuevo.Text = txtValorRec.Text
            SRegValorMedioPago()
        End With
        SValideMedPago()
    End Sub
    Private Function FenuTipoMedPagoDefecto() As EnuTipoMedioPagoDef
        Dim lenuMedioPago = EnuTipoMedioPagoDef.None
        If Not IsNothing(MobjObjetoWin) Then
            Dim lstrPredAgru = MobjObjetoWin.ObjIdPredioAgrupador_RecStr.ObjValorPro
            If lstrPredAgru.Split(",").Length > 0 Then
                lenuMedioPago = MobjObjetoWin.ObjClienteRecibo.EnuMedioPagoDefecto
            End If
            If lenuMedioPago = EnuTipoMedioPagoDef.None Then
                lenuMedioPago = MobjObjetoWin.ObjClienteRecibo.ObjIdMedioPagoClienteByt.ObjValorPro
            End If
            If lenuMedioPago = EnuTipoMedioPagoDef.None Then
                lenuMedioPago = GobjParametros.ObjIdMedioPagoDefectoByt.ObjValorPro
                If lenuMedioPago = EnuTipoMedioPagoDef.None Then
                    lenuMedioPago = EnuTipoMedioPagoDef.EnuEfectivo
                End If
            End If
        End If
        Return lenuMedioPago
    End Function
    Private Sub SCreeNuevoMedioPago()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            MobjMedioPago = MobjObjetoWin.FobjNewMedioPago
            MenuEstadoMedPago = EnuEstadoMedPagoDef.enuCreandoMedPago
            SHabiliteCtlsMedPago()
            cboTipoMedPagoNuevo.SelectedIndex = 0
            cboCuentaIngresoNuevo.SelectedIndex = 0
            txtNroMedPagoNuevo.Text = String.Empty
            txtValorMedPagoNuevo.Text = String.Empty
            SValideMedPago()
            cboTipoMedPagoNuevo.Focus()
        End If
    End Sub
    Private Sub SAcepteMedPago()
        SRegMedPago()
        If MenuEstadoMedPago = EnuEstadoMedPagoDef.enuCreandoMedPago AndAlso
                            FblnEsValidoMedPago() Then
            MobjObjetoWin.SAdicioneMedioPago(MobjMedioPago)
            MenuEstadoMedPago = EnuEstadoMedPagoDef.enuConsultandoMedPago
            SValideMediosPago()
            SHabiliteCtlsMedPago()
            cnvMediosPagoNuevo.DataContext = MobjObjetoWin.DtbMediosPago
            If dgrMediosPagoNuevo.Items.Count > 0 Then
                dgrMediosPagoNuevo.SelectedIndex = 0
            End If
            txtComentario.Focus()
        End If
        MobjObjetoWin.SVerfiqueDsctoPP()
        SMuestreDatos()
    End Sub
    Private Sub SCanceleMedPago()
        If dgrMediosPagoNuevo.Items.Count > 0 Then
            dgrMediosPagoNuevo.SelectedIndex = 0
        Else
            cboTipoMedPagoNuevo.SelectedIndex = 0
            txtNroMedPagoNuevo.Text = String.Empty
            cboCuentaIngresoNuevo.SelectedIndex = 0
            txtCuentaIngreso.Content = String.Empty
            txtValorMedPagoNuevo.Text = Format(0, "c")
        End If
        MenuEstadoMedPago = EnuEstadoMedPagoDef.enuConsultandoMedPago
        SHabiliteCtlsMedPago()
    End Sub
    Private Sub SElimineMedPago()
        If dgrMediosPagoNuevo.Items.Count > 0 AndAlso Not IsNothing(dgrMediosPagoNuevo.SelectedItem) Then
            Dim ldrvMedioPago As DataRowView = dgrMediosPagoNuevo.SelectedItem
            MblnEliminandoMedPago = True
            Dim lshrOrdinal As Short = ldrvMedioPago("Ordinal")
            MobjObjetoWin.SElimineMedPago(lshrOrdinal)
            MblnEliminandoMedPago = False
            If dgrMediosPagoNuevo.Items.Count > 0 Then
                dgrMediosPagoNuevo.SelectedIndex = 0
            Else
                MobjMedioPago = Nothing
                cboTipoMedPagoNuevo.SelectedIndex = 0
                txtCuentaIngreso.Content = String.Empty
                txtNroMedPagoNuevo.Text = String.Empty
                cboCuentaIngresoNuevo.SelectedIndex = 0
                txtValorMedPagoNuevo.Text = Format(0, "c")
            End If
        End If
        SValideMedPago()
    End Sub
#End Region
#Region "Predios Agrupadores"
    Private Sub SPuebleLsbPredAgru()
        GobjPanDat.SControleProcesoObj(True)
        Dim lstrPrediosAgr As String() = Nothing
        MblnPoblandoCbo = True
        lsbPrediosAgru.Items.Clear()
        If Not IsNothing(MobjObjetoWin.ObjClienteRecibo) AndAlso
                MobjObjetoWin.ObjClienteRecibo.BlnExiste Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
                lstrPrediosAgr = MobjObjetoWin.ObjClienteRecibo.FstrPrediosAgruClienteTodos(True)
                If Not lstrPrediosAgr.Contains(String.Empty) Then
                    ReDim Preserve lstrPrediosAgr(lstrPrediosAgr.Count)
                    lstrPrediosAgr.Append(String.Empty)
                End If
            ElseIf EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                lstrPrediosAgr = MobjObjetoWin.ObjIdPredioAgrupador_RecStr.ToString.Split(",")
            End If
            If Not IsNothing(lstrPrediosAgr) AndAlso lstrPrediosAgr.Count > 0 Then
                For Each lstrPredAgr As String In lstrPrediosAgr
                    If String.IsNullOrEmpty(lstrPredAgr) Then
                        lsbPrediosAgru.Items.Add(GCSTRSINPA)
                    Else
                        lsbPrediosAgru.Items.Add(lstrPredAgr)
                    End If
                Next
            End If
        End If
        MblnPoblandoCbo = False
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SMuestrePredAgru()
        Dim lstrPrediosAgr As String()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            lsbPrediosAgru.Items.Clear()
            If MobjObjetoWin.BlnExiste Then
                lstrPrediosAgr = MobjObjetoWin.ObjIdPredioAgrupador_RecStr.ToString.Split(",")
                If Not IsNothing(lstrPrediosAgr) AndAlso lstrPrediosAgr.Count > 0 Then
                    For Each lstrPredAgr As String In lstrPrediosAgr
                        If String.IsNullOrEmpty(lstrPredAgr) Then
                            lsbPrediosAgru.Items.Add(GCSTRSINPA)
                        Else
                            lsbPrediosAgru.Items.Add(lstrPredAgr)
                        End If
                    Next
                End If
            End If
        Else
            lsbPrediosAgru.SelectedItems.Clear()
            lstrPrediosAgr = MobjObjetoWin.ObjIdPredioAgrupador_RecStr.ToString.Split(",")
            If Not IsNothing(lstrPrediosAgr) AndAlso lstrPrediosAgr.Count > 0 Then
                For Each lstrPredAgr As String In lstrPrediosAgr
                    If String.IsNullOrEmpty(lstrPredAgr) Then
                        lsbPrediosAgru.SelectedItems.Add(GCSTRSINPA)
                    Else
                        lsbPrediosAgru.SelectedItems.Add(lstrPredAgr)
                    End If
                Next
            End If
        End If
    End Sub
    Private Function FstrPrediosAgrSele() As String
        Dim lstrPreAgr As String = String.Empty
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            If lsbPrediosAgru.SelectedItems.Count > 0 Then
                For i = 0 To lsbPrediosAgru.SelectedItems.Count - 1
                    If lsbPrediosAgru.SelectedItems(i) = GCSTRSINPA Then
                        lstrPreAgr &= ","
                    Else
                        lstrPreAgr &= lsbPrediosAgru.SelectedItems(i) & ","
                    End If
                Next
            Else
                lstrPreAgr = "***"
            End If
        End If
        If lstrPreAgr IsNot Nothing AndAlso lstrPreAgr.EndsWith(",") Then
            lstrPreAgr = lstrPreAgr.Substring(0, lstrPreAgr.Length - 1)
        End If
        Return lstrPreAgr
    End Function
#End Region
#Region "Servicios"
    Private Sub SPuebleLsbServicios(ByRef astrMens As String)
        MblnPoblandoCbo = True
        lsbServicios.Items.Clear()
        MstrUltimoServiSel = String.Empty
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando AndAlso
                Not IsNothing(MobjCliente) Then
            If chkAnticipo.IsChecked Then
                SPuebleLsbServSinDeuda()
            Else
                SPuebleLsbServDeuda()
                If lsbServicios.Items.Count > 0 Then
                    lsbServicios.SelectedItem = lsbServicios.Items(0)
                End If
            End If
        End If
        SRegServicios()
        MblnPoblandoCbo = False
        SVerifiqueDeuda(astrMens)
    End Sub
    Private Sub SPuebleLsbServSinDeuda()
        If MobjObjetoWin.ObjIdCliente_RecDbl.BlnEsValido AndAlso
                MobjObjetoWin.ObjIdPredioAgrupador_RecStr.BlnEsValido Then
            Dim lstrPredsAgru As String() =
                    MobjObjetoWin.ObjIdPredioAgrupador_RecStr.ToString().Split(",")
            Dim ldtbServiciosConDeuda = MobjCliente.FdtbServiciosConDeuda(lstrPredsAgru)
            Dim lshrIdServicio As Short
            Dim lstrFiltro As String, lstrServicios As String() = {}
            Dim i = -1, lblnAdd As Boolean
            Dim lcolServiciosPerm = GobjParametros.ColServiciosPer
            If ldtbServiciosConDeuda.Rows.Count = 0 Then
                i += 1
                ReDim Preserve lstrServicios(i)
                lstrServicios(i) = My.Resources.Todos
            End If
            If GobjParametros.ObjPermiteAnticipoPorServicioBln.ObjValorPro Then
                lstrFiltro = ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd & " <> 0"
                If ldtbServiciosConDeuda.Select(lstrFiltro).Length = 0 Then
                    i += 1
                    ReDim Preserve lstrServicios(i)
                    lstrServicios(i) = My.Resources.CuotasAdmin
                End If
                For Each lobjServicio As ClsServicio In lcolServiciosPerm
                    If lobjServicio.ObjEstaActivoServicioBln.ObjValorPro Then
                        lshrIdServicio = lobjServicio.ObjIdServicioShr.ObjValorPro
                        lstrFiltro = ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd & " = 0 AND " &
                                ClsIdServicio_ItemFactShr.SstrNombreCampoBd & " = " & lshrIdServicio
                        If ldtbServiciosConDeuda.Select(lstrFiltro).Length = 0 Then
                            i += 1
                            ReDim Preserve lstrServicios(i)
                            lstrServicios(i) = lobjServicio.ObjConceptoServicioStr.ObjValorPro
                        End If
                    End If
                Next
            End If
            For Each lstrServicio As String In lstrServicios
                lblnAdd = lsbServicios.Items.Add(lstrServicio)
            Next
        End If
    End Sub
    Private Sub SPuebleLsbServDeuda()
        If MobjObjetoWin.ObjIdCliente_RecDbl.BlnEsValido AndAlso
                MobjObjetoWin.ObjIdPredioAgrupador_RecStr.BlnEsValido Then
            Dim lstrPredsAgru As String() =
                    MobjObjetoWin.ObjIdPredioAgrupador_RecStr.ToString().Split(",")
            Dim ldtbServiciosConDeuda = MobjCliente.FdtbServiciosConDeuda(lstrPredsAgru)
            Dim lshrIdAno As Short, lshrIdServicio As Short
            Dim lstrKeySer As String
            Dim lobjServicio As ClsServicio, lstrSer As String
            If ldtbServiciosConDeuda.Rows.Count > 0 Then
                lsbServicios.Items.Add(My.Resources.Todos)
            End If
            For Each ldrwAgrSerDeu As DataRow In ldtbServiciosConDeuda.Rows
                lshrIdAno = ClsPanorama.FobjValorCampo(ldrwAgrSerDeu(0), EnuTipoValor.enuShort)
                lshrIdServicio =
                        ClsPanorama.FobjValorCampo(ldrwAgrSerDeu(1), EnuTipoValor.enuShort)
                lstrKeySer = lshrIdAno.ToString() & "," & lshrIdServicio.ToString()
                If lshrIdAno > 0 Then
                    lobjServicio = GobjParametros.FobjServicio(lstrKeySer)
                    lstrSer = My.Resources.CuotasAdmin
                Else
                    lobjServicio = GobjParametros.ColServiciosPer(lstrKeySer)
                    lstrSer = lobjServicio.ObjConceptoServicioStr.ToString()
                End If
                If Not lsbServicios.Items.Contains(lstrSer) Then
                    lsbServicios.Items.Add(lstrSer)
                End If
            Next
        End If
    End Sub
    Private Function FstrServiSele() As String
        Dim lstrSerSel = String.Empty, lstrNomSer As String, lstrIdSer As String
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            If lsbServicios.SelectedItems.Count > 0 Then
                If String.IsNullOrEmpty(MstrUltimoServiSel) Then
                    If lsbServicios.SelectedItems(0) = My.Resources.Todos Then
                        MstrUltimoServiSel = My.Resources.Todos
                    End If
                End If
                If lsbServicios.SelectedItems.Count > 1 Then
                    If MstrUltimoServiSel = My.Resources.Todos Then
                        MstrUltimoServiSel = lsbServicios.SelectedItems(lsbServicios.SelectedItems.Count - 1)
                        lsbServicios.UnselectAll()
                        lsbServicios.SelectedItems.Add(MstrUltimoServiSel)
                    ElseIf lsbServicios.SelectedItems.Contains(My.Resources.Todos) Then
                        lsbServicios.UnselectAll()
                        lsbServicios.SelectedItems.Add(My.Resources.Todos)
                        MstrUltimoServiSel = My.Resources.Todos
                    End If
                End If
                For i = 0 To lsbServicios.SelectedItems.Count - 1
                    If lsbServicios.SelectedItems(i) = My.Resources.Todos Then
                        lstrSerSel &= "A"
                    ElseIf lsbServicios.SelectedItems(i) = My.Resources.CuotasAdmin Then
                        lstrSerSel &= "0,"
                    ElseIf lsbServicios.SelectedItems(i) = My.Resources.Todos Then
                        lstrSerSel &= " ,"
                    Else
                        lstrNomSer = lsbServicios.SelectedItems(i)
                        lstrIdSer = GobjParametros.FshrIdServicio(lstrNomSer, False).ToString()
                        lstrSerSel &= lstrIdSer & ","
                    End If
                Next
            Else
                lstrSerSel = String.Empty
            End If
        End If
        If Not String.IsNullOrEmpty(lstrSerSel) AndAlso lstrSerSel.EndsWith(",") Then
            lstrSerSel = lstrSerSel.Substring(0, lstrSerSel.Length - 1)
        End If
        Return lstrSerSel
    End Function
#End Region
#Region "CuentasIngresos"
    Private Sub SPuebleCboCtasIngresos()
        If cboCuentaIngresoNuevo.Items.Count = 0 Then
            Dim ldtbCuentasIng = ClsOrionCop.FdtbCuentasIngresos
            Dim lstrCuenta As String, lentIdBanco As Integer
            For Each ldrwCuenta As DataRow In ldtbCuentasIng.Rows
                lentIdBanco = ClsPanorama.FobjValorCampo(ldrwCuenta(ClsIdCuentaBancoShr.SstrNombreCampoBd),
                        EnuTipoValor.enuShort)
                If lentIdBanco <= 0 Then
                    lstrCuenta = ldrwCuenta(ClsNombreBancoStr.SstrNombreCampoBd)
                Else
                    lstrCuenta = ldrwCuenta(ClsNombreBancoStr.SstrNombreCampoBd) & " Cuenta número " &
                        ldrwCuenta(ClsNumeroCuentaStr.SstrNombreCampoBd)
                End If
                cboCuentaIngresoNuevo.Items.Add(lstrCuenta)
            Next
        End If
    End Sub
    Private Function FstrCtaContabilidadIngresos() As String
        Static lobjCtaBanco As New ClsCuentaBanco(EnuModoInstanciaObjDef.enuNavegable)
        Dim lstrCtaContIng = String.Empty, lblnProcesarCboItem = False
        If Not IsNothing(cboCuentaIngresoNuevo.SelectedItem) Then
            If cboCuentaIngresoNuevo.SelectedIndex = 1 Then
                lstrCtaContIng = GobjParametros.ObjIdCtaCajaStr.ObjValorPro
            ElseIf cboCuentaIngresoNuevo.SelectedIndex = cboCuentaIngresoNuevo.Items.Count - 1 Then
                If GobjParametros.ObjIdCtaIngPorIdentificarStr.ToString.Length > 0 Then
                    lstrCtaContIng = GobjParametros.ObjIdCtaIngPorIdentificarStr.ObjValorPro
                    lblnProcesarCboItem = IsNothing(lstrCtaContIng)
                Else
                    lblnProcesarCboItem = True
                End If
            ElseIf cboCuentaIngresoNuevo.SelectedIndex > 1 Then
                lblnProcesarCboItem = True
            End If
            If lblnProcesarCboItem Then
                Dim lstrCtaBco = cboCuentaIngresoNuevo.SelectedItem.ToString().ToUpper()
                lobjCtaBanco.SVayaAlPrimero()
                Do While lobjCtaBanco.BlnExiste
                    If lstrCtaBco.Contains(lobjCtaBanco.ObjNombreBancoStr.ToString.ToUpper()) AndAlso
                            lstrCtaBco.Contains(lobjCtaBanco.ObjNumeroCuentaStr.ToString.ToUpper()) Then
                        lstrCtaContIng = lobjCtaBanco.ObjIdCtaContabilidadStr.ObjValorPro
                        Exit Do
                    End If
                    lobjCtaBanco.SVayaAlSiguiente()
                Loop
                If String.IsNullOrEmpty(lstrCtaBco) Then
                    Throw New ErrorInesperadoPanLException("Cuenta Banco no encontrada!")
                End If
            End If
        End If
        Return lstrCtaContIng
    End Function
#End Region
#Region "Validaciones"
#Region "Validacion Descuentos"
    Private Sub SValideTipoDscto()
        Dim lblnEsValido As Boolean = cboTipoDsctoNuevo.SelectedIndex > 0
        StcValidValido(EnuValidEntradaDef.enuTipoDscto) = lblnEsValido
        If MenuEstadoDscto = EnuEstadoDsctoDef.enuCreandoDscto Then
            SValideDscto()
        End If
    End Sub
    Private Sub SValideFraDscto()
        StcValidValido(EnuValidEntradaDef.enuFraDscto) = cboNroFactNuevo.SelectedIndex > 0
        If MenuEstadoDscto = EnuEstadoDsctoDef.enuCreandoDscto Then
            SValideDscto()
        End If
    End Sub
    Private Sub SValideItemDscto()
        StcValidValido(EnuValidEntradaDef.enuItemDscto) = cboItemFactNuevo.SelectedIndex > 0
        If MenuEstadoDscto = EnuEstadoDsctoDef.enuCreandoDscto Then
            SValideDscto()
        End If
    End Sub
    Private Sub SValideDscto()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            Dim lblnEsValido As Boolean = False
            If MenuEstadoDscto = EnuEstadoDsctoDef.enuCreandoDscto Then
                If StcValidValido(EnuValidEntradaDef.enuFraDscto) AndAlso
                        StcValidValido(EnuValidEntradaDef.enuItemDscto) AndAlso
                        StcValidValido(EnuValidEntradaDef.enuTipoDscto) Then
                    If cboItemFactNuevo.SelectedItem IsNot Nothing Then
                        Dim lstrItemFac = cboItemFactNuevo.SelectedItem
                        Dim lshrIdItemFac = 0S
                        MobjObjetoWin.StrNroFacDscto = cboNroFactNuevo.SelectedItem
                        If lstrItemFac <> My.Resources.Ninguno Then
                            lshrIdItemFac = lstrItemFac.Substring(0, lstrItemFac.IndexOf("-"))
                        End If
                        MobjObjetoWin.ShrIdItemFacDscto = lshrIdItemFac
                        MobjObjetoWin.EnuTipoDscto = cboTipoDsctoNuevo.SelectedIndex
                        MobjObjetoWin.StrVlrDscto = txtValorDctoNuevo.Text
                        lblnEsValido = MobjObjetoWin.FblnEsValidoDscto
                    Else
                        lblnEsValido = False
                    End If
                End If
            Else
                If MdtbDescuentos Is Nothing OrElse MdtbDescuentos.Rows.Count = 0 Then
                    lblnEsValido = True
                Else
                    lblnEsValido = Not String.IsNullOrEmpty(txtItemFactura.Text)
                End If
            End If
            If lblnEsValido Then
                SLevanteEveOk()
            End If
            StcValidValido(EnuValidEntradaDef.enuDscto) = lblnEsValido
            FblnEstanTodosBien()
        Else
            StcValidValido(EnuValidEntradaDef.enuTipoDscto) = True
            StcValidValido(EnuValidEntradaDef.enuFraDscto) = True
            StcValidValido(EnuValidEntradaDef.enuItemDscto) = True
            StcValidValido(EnuValidEntradaDef.enuDscto) = True
        End If
    End Sub
#End Region
#Region "Validacion Pago"
    Private Sub SValideValorRecibo()
        If Not IsNothing(MobjObjetoWin) Then
            StcValidValido(EnuValidEntradaDef.enuValor) = MobjObjetoWin.ObjValor_RecDec.BlnEsValido
        End If
        FblnEstanTodosBien()
    End Sub
    Private Sub SValideMedPago()
        Dim ldecTotalMediosPago = 0D
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            If Not IsNothing(MobjMedioPago) Then
                With MobjMedioPago
                    StcValidValido(EnuValidEntradaDef.enuTipoMedPago) = .ObjIdTipoMedPago_MedPagoByt.BlnEsValido
                    StcValidValido(EnuValidEntradaDef.enuNroMedPago) = .ObjNumeroMedPagoStr.BlnEsValido
                    StcValidValido(EnuValidEntradaDef.enuNroCtaCon) = .ObjIdCtaContabIngresoStr.BlnEsValido
                    StcValidValido(EnuValidEntradaDef.enuValorMedPago) = .ObjValor_MedPagoDec.BlnEsValido
                End With
            Else
                StcValidValido(EnuValidEntradaDef.enuTipoMedPago) = False
                StcValidValido(EnuValidEntradaDef.enuNroCtaCon) = False
                StcValidValido(EnuValidEntradaDef.enuValorMedPago) = False
            End If
            If MenuEstadoMedPago = EnuEstadoMedPagoDef.enuCreandoMedPago Then
                SHabiliteBotonesWin()
            End If
            SValideMediosPago()
        Else
            ldecTotalMediosPago = MobjObjetoWin.FdecValorTotalMedPago
            txtTotalMedPagoNuevo.Content = Format(ldecTotalMediosPago, "c")
            StcValidValido(EnuValidEntradaDef.enuTipoMedPago) = True
            StcValidValido(EnuValidEntradaDef.enuNroMedPago) = True
            StcValidValido(EnuValidEntradaDef.enuNroCtaCon) = True
            StcValidValido(EnuValidEntradaDef.enuValorMedPago) = True
            StcValidValido(EnuValidEntradaDef.enuMediosPago) = True
        End If
        FblnEstanTodosBien()
    End Sub
    Private Sub SValideMediosPago()
        Dim ldecTotalMediosPago = MobjObjetoWin.FdecValorTotalMedPago
        Dim lblnEsValido = (ldecTotalMediosPago > 0)
        If lblnEsValido Then
            lblnEsValido = ldecTotalMediosPago = MobjObjetoWin.ObjValor_RecDec.ObjValorPro
        End If
        txtTotalMedPagoNuevo.Content = Format(ldecTotalMediosPago, "c")
        StcValidValido(EnuValidEntradaDef.enuMediosPago) = lblnEsValido
    End Sub
    Private Function FblnEsValidoMedPago() As Boolean
        Dim lblnEsValido = StcValidValido(EnuValidEntradaDef.enuTipoMedPago) AndAlso
            StcValidValido(EnuValidEntradaDef.enuNroMedPago) AndAlso
            StcValidValido(EnuValidEntradaDef.enuNroCtaCon) AndAlso
            StcValidValido(EnuValidEntradaDef.enuValorMedPago)
        Return lblnEsValido
    End Function
#End Region
#End Region
#Region "Deuda"
    Private Sub SVerifiqueDeuda(ByRef astrMens As String)
        If MobjObjetoWin.ObjIdCliente_RecDbl.BlnEsValido Then
            If lsbServicios.Items.Count = 0 AndAlso lsbPrediosAgru.SelectedItems.Count > 0 Then
                astrMens = If(chkAnticipo.IsChecked,
                    "No se puede hacer un anticipo. No hay Predios Agrupadores sin Deuda.",
                    "El Cliente y el Predio agrupador seleccionado, no tiene Cuentas por pagar!")
                lsbServicios.Items.Clear()
            End If
            SMuestreDatos()
            SProceseDeuda()
        End If
        If String.IsNullOrEmpty(astrMens) Then
            SValide()
        End If
    End Sub
    Private Sub SVerifiqueOtrasDeudas()
        Dim ldblIdTerceroDiffConDeuda As Double
        Dim lobjPredio As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
        Dim lstrIdPredio As String, lstrPreAgr As String()
        If MobjObjetoWin.ObjIdPredioAgrupador_RecStr.BlnEsValido Then
            lstrPreAgr = MobjObjetoWin.ObjIdPredioAgrupador_RecStr.ObjValorPro.Split(",")
            For i = 0 To lstrPreAgr.Length - 1
                lstrIdPredio = lstrPreAgr(i)
                If Not String.IsNullOrEmpty(lstrIdPredio) Then
                    lobjPredio.SAbra({GshrIdCarpeta, GshrIdCentroUtil, lstrIdPredio})
                    ldblIdTerceroDiffConDeuda = lobjPredio.FdblIdClienteDiferenteConDeuda(
                            MobjCliente.ObjIdClienteDbl.ObjValorPro)
                    If ldblIdTerceroDiffConDeuda <> 0 Then
                        Dim lstrMens = "El Predio Agrupador " & lstrIdPredio & " tiene deuda " &
                                    "con el Cliente " & ldblIdTerceroDiffConDeuda & "!"
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuAdvertencia)
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub
    Private Sub SProceseDeuda()
        cnvDeuda.DataContext = Nothing
        MdecIntPorCausar = MobjObjetoWin.DecInteresesPorCausar
        MdecDeuda = MobjObjetoWin.DecDeuda
        If MdecIntPorCausar = 0 Then
            SProceseDetalleDeuda()
            SMuestreSaldoAPagar()
        End If
        SProceseIntereses()
    End Sub
    Private Sub SProceseIntereses()
        MdecIntPorCausar = MobjObjetoWin.DecInteresesPorCausar
        If MdecIntPorCausar > 0 Then
            lblIntPorCausar.Visibility = Visibility.Visible
            txtIntPorCausar.Visibility = Visibility.Visible
            bttCausarIntereses.Visibility = Visibility.Visible
            txtIntPorCausar.Content = Format(MdecIntPorCausar, "c")
        Else
            lblIntPorCausar.Visibility = Visibility.Collapsed
            txtIntPorCausar.Visibility = Visibility.Collapsed
            bttCausarIntereses.Visibility = Visibility.Collapsed
        End If
        StcValidValido(EnuValidEntradaDef.enuCausarInt) = MdecIntPorCausar = 0
    End Sub
    Private Sub SProceseDetalleDeuda()
        If MdecDeuda > 0 Then
            Dim lstrPredAgru = MobjObjetoWin.ObjIdPredioAgrupador_RecStr.ToString().Split(",")
            Dim lstrServicios = MobjObjetoWin.ObjServicios_RecStr.ToString().Split(",")
            Dim ldtbDetDeuda = MobjCliente.FdtbDetalleDeuda(lstrPredAgru, lstrServicios)
            cnvDeuda.DataContext = ldtbDetDeuda
        Else
            cnvDeuda.DataContext = Nothing
        End If
    End Sub
    Private Sub SMuestreSaldoAPagar()
        txtSaldo.Content = Format(MobjObjetoWin.FdecSaldoRC, "c")
        txtDeudaRecibo.Content = txtSaldo.Content
    End Sub
    Private Sub SCauseMora()
        Dim lblnNoHayError = False, ldecValorCausado = 0D
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lenuSevNot As EnuSeveridadNot
        Try
            Mouse.OverrideCursor = Cursors.Wait
            GobjPanDat.SControleProcesoObj(True)
            ldecValorCausado = MobjObjetoWin.SCauseMoraClienteRC(lstrMens)
            If Not String.IsNullOrEmpty(lstrMens) Then
                lenuSevNot = EnuSeveridadNot.EnuFalta
            End If
            If ldecValorCausado > 0 Then
                If GobjParametros.BlnEFacAutorizado Then
                    MblnCausoMora = True
                End If
            End If
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
                StcValidValido(EnuValidEntradaDef.enuCausarInt) = True
                lstrMens = "Se causo Intereses de Mora por un valor de " &
                    Format(ldecValorCausado, "c") & "!"
                lenuSevNot = EnuSeveridadNot.EnuInformacion
                SProceseDeuda()
                GobjPanDat.SControleProcesoObj(False)
            Else
                lenuSevNot = EnuSeveridadNot.EnuExcep
                GobjPanDat.SControleProcesoObj(False, True)
            End If
            Mouse.OverrideCursor = Cursors.Arrow
            If Not String.IsNullOrEmpty(lstrMens) Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, lenuSevNot)
            End If
        End Try
    End Sub
#End Region
#Region "Generales"
    ''' <summary>
    ''' Si la fecha del Rc, el Cliente, y el Concepto de pago y los predios agrupadores 
    ''' seleccionados son válidos se ejecuta este método que es llamado despues de registrar
    ''' en la clase los datos citados
    ''' </summary>
    Private Sub SProceseGrales()
        SProceseDeuda()
        If Not FblnGeneralesOk() Then
            txtSaldo.Content = Format(0, "c")
            txtDeudaRecibo.Content = Format(0, "c")
        End If
    End Sub
    Private Function FblnGeneralesOk() As Boolean
        Dim lblnGenOk = StcValidValido(EnuValidEntradaDef.enuCliente) AndAlso
                StcValidValido(EnuValidEntradaDef.enuFecha) AndAlso
                StcValidValido(EnuValidEntradaDef.enuPreAgr) AndAlso
                StcValidValido(EnuValidEntradaDef.enuServicios)
        Return lblnGenOk
    End Function
    Private Sub SInicialiceRec()
        If MobjObjetoWin IsNot Nothing Then
            If cboPref.SelectedItem Is Nothing Then
                cboPref.SelectedIndex = 0
            End If
            Dim lstrPref As String = cboPref.SelectedItem
            If MobjObjetoWin.ObjPrefijo_RecStr.ToString() <> lstrPref Then
                ObjObjetoWin = New ClsReciboCaja(lstrPref)
                MobjObjetoWin = ObjObjetoWin
                If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                    MobjObjetoWin.SVayaAlUltimo()
                End If
            End If
        End If
    End Sub
    Private Sub SAbraRecCaja()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                If cboPref.SelectedItem <> MobjObjetoWin.ObjPrefijo_RecStr.ToString() OrElse
                        txtNroRecCaja.Text <> MobjObjetoWin.ObjIdRecCajaEnt.ToString() Then
                    Dim lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, cboPref.SelectedItem,
                            txtNroRecCaja.Text}
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
    Private Sub SPaseAlSiguiente()
        Select Case MenuEstadoCaptura
            Case EnuEstadoCapturaDef.enuGenerales
                If FblnGeneralesOk() Then
                    Dim lblnAbrirDsctos = Not chkAnticipo.IsChecked
                    If lblnAbrirDsctos Then
                        lblnAbrirDsctos = MobjObjetoWin.DtbDescuentos IsNot Nothing AndAlso
                                MobjObjetoWin.DtbDescuentos.Rows.Count > 0
                        If Not lblnAbrirDsctos Then
                            lblnAbrirDsctos = MsgBox("Debe entrar a hacer Descuentos?", vbYesNo,
                                    "Abrir Descuentos") = vbYes
                        End If
                        If lblnAbrirDsctos Then
                            MenuEstadoCaptura = EnuEstadoCapturaDef.enuCreditos
                            MenuEstadoDscto = EnuEstadoDsctoDef.enuConsultandoDscto
                            txtDeudaCap.Content = Format(0, "c")
                            txtDeudaMora.Content = Format(0, "c")
                            MblnPrimeraVezMediosPago = True
                            SCargueDescuentos()
                            SHabiliteControlesEstado()
                        Else
                            SPaseAValores()
                        End If
                    Else
                        SPaseAValores()
                    End If
                End If
            Case EnuEstadoCapturaDef.enuCreditos
                If StcValidValido(EnuValidEntradaDef.enuDscto) Then
                    SPaseAValores()
                End If
        End Select
    End Sub
    Private Sub SPaseAValores()
        MenuEstadoCaptura = EnuEstadoCapturaDef.enuValores
        MblnPrimeraVezMediosPago = True
        txtDeuda.Content = txtDeudaRecibo.Content
        MobjObjetoWin.ObjValorDeudaAlPagoDec.ObjValorPro = txtDeuda.Content
        MobjObjetoWin.ObjComentario_RecStr.ObjValorPro = txtComentario.Text
        SHabiliteControlesEstado()
        txtValorRec.Text = txtDeuda.Content
        txtRecibido.Content = txtValorRec.Text
        If Not chkAnticipo.IsChecked Then
            SRegValorRec()
            SValideValorRecibo()
            SDetermineMedPagoInicial()
            SValideMedPago()
        Else
            cboTipoMedPagoNuevo.SelectedIndex = 0
        End If
        txtValorRec.Focus()
        txtValorRec.SelectAll()
    End Sub
    Private Sub SPaseAlAnterior()
        MobjObjetoWin.ObjValorAnticipoDec.ObjValorPro = 0
        MobjObjetoWin.ObjValor_RecDec.ObjValorPro = 0
        MobjObjetoWin.SLimpieMediosPago()
        SMuestreSaldoAPagar()
        Select Case MenuEstadoCaptura
            Case EnuEstadoCapturaDef.enuCreditos
                MenuEstadoCaptura = EnuEstadoCapturaDef.enuGenerales
                MblnPrimeraVezGrales = True
                SHabiliteControlesEstado()
                If FblnGeneralesOk() Then
                    bttSiguiente.IsEnabled = StcValidValido(EnuValidEntradaDef.enuCausarInt)
                End If
            Case EnuEstadoCapturaDef.enuValores
                txtValorAnt.Content = Format(0, "c")
                txtCuentaIngreso.Content = String.Empty
                txtTotalMedPagoNuevo.Content = Format(0, "c")
                txtComentario.Text = String.Empty
                MenuEstadoCaptura = EnuEstadoCapturaDef.enuGenerales
                SHabiliteControlesEstado()
        End Select
    End Sub
    Private Sub SCargueDtbDescuentos()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando AndAlso MblnPrimeraVezGrales Then
            MdtbDescuentos = Nothing
            Dim lstrPredAgru = MobjObjetoWin.ObjIdPredioAgrupador_RecStr.ObjValorPro
            If Not IsNothing(lstrPredAgru) Then
                Dim lstrPresAgr As String() = lstrPredAgru.Split(",")
                If Not (IsNothing(MobjCliente) OrElse IsNothing(lstrPresAgr)) Then
                    If MobjObjetoWin.ObjIdPredioAgrupador_RecStr.BlnEsValido Then
                        Dim lstrServicios As String() = FstrServiSele().Split(",")
                        MdtbDescuentos = MobjCliente.FdtbDescuentos(lstrPresAgr,
                                lstrServicios, MobjObjetoWin.ObjFechaRecDtm.ObjValorPro)
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub SMuestreServicios()
        Dim lstrNombServ As String = String.Empty
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            lsbServicios.Items.Clear()
            If MobjObjetoWin.BlnExiste AndAlso
                    MobjObjetoWin.ObjServicios_RecStr.BlnEsValido Then
                Dim lstrSers As String() = MobjObjetoWin.ObjServicios_RecStr.ToString().Split(",")
                For Each lstrSer As String In lstrSers
                    If lstrSer = "0" Then
                        lstrNombServ = My.Resources.Admin
                    ElseIf lstrSer = "A" Then
                        lstrNombServ = "Pago aplicado a todos los Servicios!"
                    ElseIf Not String.IsNullOrEmpty(lstrSer) Then
                        lstrNombServ = GobjParametros.FstrNombreServicio(lstrSer)
                    Else
                        If MobjObjetoWin.ObjIdAnticipo_RecEnt.ObjValorPro = 0 Then
                            lstrNombServ = My.Resources.SinInfo
                        End If
                    End If
                    If Not String.IsNullOrEmpty(lstrNombServ) Then
                        lsbServicios.Items.Add(lstrNombServ)
                    End If
                Next
                If MobjObjetoWin.ObjIdAnticipo_RecEnt.ObjValorPro > 0 Then
                    lsbServicios.Items.Add(My.Resources.AntRec)
                End If
            End If
        End If
    End Sub
    Private Sub SPuebleFras()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            MblnPoblandoCbo = True
            cboNroFactNuevo.Items.Clear()
            cboNroFactNuevo.Items.Add(My.Resources.Ninguno)
            Dim lstrIdFacturas As New ArrayList
            If Not IsNothing(MobjCliente) Then
                Dim lstrPredAgru = MobjObjetoWin.ObjIdPredioAgrupador_RecStr.ObjValorPro
                Dim lstrPreAgr As String() = lstrPredAgru.Split(",")
                If cboTipoDsctoNuevo.SelectedIndex = EnuTipoDescuentoDef.EnuDsctoIntMora Then
                    lstrIdFacturas = MobjCliente.FstrIdFacturasConMora(lstrPreAgr)
                Else
                    If MobjObjetoWin.ObjIdPredioAgrupador_RecStr.BlnEsValido Then
                        Dim lstrServicios As String() = FstrServiSele().Split(",")
                        lstrIdFacturas = MobjCliente.FstrIdFacturasVivas(lstrPreAgr, lstrServicios)
                    End If
                End If
            End If
            If Not IsNothing(lstrIdFacturas) Then
                lstrIdFacturas.Sort()
                For Each lstrIdFact As String In lstrIdFacturas
                    cboNroFactNuevo.Items.Add(lstrIdFact)
                Next
            End If
            MblnPoblandoCbo = False
            cboNroFactNuevo.SelectedIndex = 0
            cboItemFactNuevo.Items.Clear()
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
                Dim lblnInserte = False
                Dim lstrServiciosRec = MobjObjetoWin.ObjServicios_RecStr.ToString()
                Dim lstrServicioItFac = String.Empty
                For Each lobjItemFra As ClsItemFactura In MobjFactura.ColItemsFactura
                    If cboTipoDsctoNuevo.SelectedIndex = EnuTipoDescuentoDef.EnuDsctoIntMora Then
                        lblnInserte = lobjItemFra.FdecDeudaIntMora > 0
                    Else
                        lblnInserte = lstrServiciosRec.Contains("A")
                        If Not lblnInserte Then
                            If lobjItemFra.ObjIdAno_ServicioItemFactShr.ObjValorPro > 0 Then
                                lblnInserte = lstrServiciosRec.Contains("0")
                            Else
                                lstrServicioItFac = lobjItemFra.ObjIdServicio_ItemFactShr.ToString()
                                lblnInserte =
                                    lobjItemFra.ObjIdAno_ServicioItemFactShr.ObjValorPro = 0 AndAlso
                                    lstrServiciosRec.Contains(lstrServicioItFac)
                            End If
                        End If
                    End If
                    If lblnInserte Then
                        lstrItem = lobjItemFra.ObjIdItemFacturaShr.ToString & "-" &
                            lobjItemFra.ObjDetalle_ItemFactStr.ObjValorPro
                        If Not String.IsNullOrEmpty(lstrItem) Then
                            cboItemFactNuevo.Items.Add(lstrItem)
                        End If
                    End If
                Next
                cboItemFactNuevo.SelectedIndex = 0
            End If
        End If
        MblnPoblandoCbo = False
    End Sub
    Private Sub SPuebleCboPref()
        MblnPoblandoCbo = True
        Dim ldrwConst = ClsOrionCop.FdrwPrefDoc(EnuTipoDocOri.EnuReciboCaja)
        SPuebleComboBox(ldrwConst, cboPref)
        Dim lstrPref = GobjParametros.FstrPrefijoDoc(EnuIdDocumentoDef.EnuReciboCaja)
        If Not cboPref.Items.Contains(lstrPref) Then
            cboPref.Items.Add(lstrPref)
        End If
        MblnPoblandoCbo = False
    End Sub
    Private Sub SPuebleComboBoxes()
        MblnPoblandoCbo = True
        Dim ldrwTposDcto = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuTipoDescuento)
        SPuebleComboBox(ldrwTposDcto, cboTipoDsctoNuevo)
        cboTipoDsctoNuevo.SelectedItem = cboTipoDsctoNuevo.Items(8)
        cboTipoDsctoNuevo.Items.Remove(cboTipoDsctoNuevo.SelectedItem)
        cboTipoDsctoNuevo.SelectedIndex = 0
        Dim ldrwMediosPago = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.enuMediosPago)
        SPuebleComboBox(ldrwMediosPago, cboTipoMedPagoNuevo)
        cboTipoMedPagoNuevo.SelectedIndex = 0
        MblnPoblandoCbo = False
    End Sub
    Private Sub SEstablezcaDataContext()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            GobjPanDat.SControleProcesoObj(True)
            Select Case True
                Case tbiDetalles.IsSelected
                    SEnlaceItems()
                Case tbiMediosPago.IsSelected
                    SEnlaceMedPago()
                Case tbiNovedades.IsSelected
                    SEnlaceNovedades()
            End Select
            GobjPanDat.SControleProcesoObj(False)
        End If
    End Sub
    Private Sub SEnlaceItems()
        If MobjObjetoWin.BlnExiste Then
            Dim ldtbItemsRC = MobjObjetoWin.FdtbInfItemsRC()
            tbiDetalles.DataContext = ldtbItemsRC
            SOrdeneDataGrid(dgrItems, dgrItems.Columns(0),
                        "Ordinal", ListSortDirection.Ascending)
        Else
            tbiDetalles.DataContext = Nothing
        End If
    End Sub
    Private Sub SEnlaceMedPago()
        If MobjObjetoWin.BlnExiste Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                tbiMediosPago.DataContext = MobjObjetoWin.DtbMediosPago
                If dgrMediosPago.Items.Count > 0 Then
                    Dim e As SelectedCellsChangedEventArgs = Nothing
                    DgrMediosPago_SelectedCellsChanged(dgrMediosPago, e)
                End If
                SMuestreMediosPago()
            End If
        Else
            tbiMediosPago.DataContext = Nothing
        End If
    End Sub
    Private Sub SEnlaceNovedades()
        If MobjObjetoWin.BlnExiste Then
            dgrNovedades.DataContext = MobjObjetoWin.FdtbNovedadesRC
        Else
            dgrNovedades.DataContext = Nothing
        End If
    End Sub
    ''' <summary>
    ''' En estado de creacion, informa si alguno de los predios agrupadores seleccionados 
    ''' tiene deuda con un cliente diferente
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SEnvieEmailRecCaja()
        Dim lstrMens = String.Empty
        If MobjObjetoWin.ObjClienteRecibo.ObjRecibeDocsPorEmailBln.ObjValorPro Then
            SEnvieCorreo(Me, EnuTipoCorreoE.EnuRC,
                    MobjObjetoWin.ObjClienteRecibo.ObjIdClienteDbl.ObjValorPro,
                    MobjObjetoWin.ObjIdPredioAgrupador_RecStr.ObjValorPro,
                    MobjObjetoWin.StrNumeroRecCaja)
        Else
            lstrMens = "El Cliente del Recibo de Caja no tiene habilitado el envio de " &
                                          "Documentos por Email!"
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SExporteRecibos()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        GobjPanDat.SControleProcesoObj(True)
        MobjReportesOrion = New ClsRepOrionCop(GCOBJREGISTRO)
        Try
            lstrMens = "Exportando Recibos de Caja!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            lstrMens = String.Empty
            ClsOrionCop.SCreeCarpetaRec()
            MobjReportesOrion.EnuReporte = EnuReporteDef.enuExpRecsCaja
            MobjReportesOrion.SGenereReporte()
            lblnNoHayError = True
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                GobjPanDat.SControleProcesoObj(False)
                lstrMens = "Proceso terminado exitosamente!"
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Else
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            MobjReportesOrion = Nothing
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub
#End Region
#Region "Procesos eFac"
    Private Sub SProceseNdb()
        Dim lstrMens = String.Empty
        SProceseEFac(lstrMens)
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SProceseNsCr()
        Dim lstrMens = String.Empty
        SProceseEFac(lstrMens)
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SProceseNsRcr()
        Dim lstrMens = String.Empty
        SProceseEFac(lstrMens)
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
#End Region
#End Region
#Region "Registro Entradas"
    ' Generales
    Private Sub SEstablezcaCtrlIni()
        With GobjParametros
            MobjObjetoWin.SEstablezcaFechaRec()
            SMuestreDatos()
            If .ObjExigeFechaHoyDocsBln.ObjValorPro Then
                dtpFechaRec.Style = FindResource("RecCtlNoHabilitado")
                bttEncontrarCliente.Focus()
            Else
                dtpFechaRec.Focus()
            End If
        End With
    End Sub
    Private Sub SRegCliente()
        With MobjObjetoWin
            If .ObjIdCliente_RecDbl.ToString().Trim <> txtIdCliente.Text.Trim Then
                .ObjIdCliente_RecDbl.ObjValorPro = txtIdCliente.Text
                If .ObjIdCliente_RecDbl.BlnEsValido Then
                    MobjCliente = .ObjClienteRecibo
                    txtNombreCliente.Content = MobjCliente.ObjNombreCompletoStr.ToString()
                    SPuebleLsbPredAgru()
                    lsbPrediosAgru.Focus()
                Else
                    lsbPrediosAgru.Items.Clear()
                End If
            End If
        End With
        SLevanteEveOk()
        SMuestreDatos()
        lsbPrediosAgru.Focus()
        MblnPrimeraVezGrales = True
    End Sub
    Private Sub SRegServicios()
        Dim lstrServSele = FstrServiSele()
        MobjObjetoWin.ObjServicios_RecStr.ObjValorPro = lstrServSele
        StcValidValido(EnuValidEntradaDef.enuServicios) =
                    MobjObjetoWin.ObjServicios_RecStr.BlnEsValido
        SCargueDescuentos()
        SProceseGrales()
        SHabiliteControlesEstado()
    End Sub
    Private Sub SRegPrediosAgr(ByRef astrMens As String)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            MobjObjetoWin.ObjIdPredioAgrupador_RecStr.ObjValorPro = FstrPrediosAgrSele()
            StcValidValido(EnuValidEntradaDef.enuPreAgr) =
                    MobjObjetoWin.ObjIdPredioAgrupador_RecStr.BlnEsValido
            SProceseGrales()
            SPuebleLsbServicios(astrMens)
            SHabiliteControlesEstado()
        End If
    End Sub
    ' Descuentos
    Private Sub SRegNroFactDscto()
        If MenuEstadoDscto = EnuEstadoDsctoDef.enuCreandoDscto Then
            StcValidValido(EnuValidEntradaDef.enuFraDscto) = False
            StcValidValido(EnuValidEntradaDef.enuItemDscto) = False
            If MobjFactura IsNot Nothing AndAlso cboItemFactNuevo.SelectedItem IsNot Nothing AndAlso
                    Not cboItemFactNuevo.SelectedIndex = 0 Then
                Dim lstrItemFac = cboItemFactNuevo.SelectedItem
                Dim lstrIdItemFac As String = lstrItemFac.Substring(0, lstrItemFac.IndexOf("-"))
                Dim lobjItemFac As ClsItemFactura = MobjFactura.ColItemsFactura(lstrIdItemFac)
                txtDeudaCap.Content = Format(lobjItemFac.FdecDeudaCapital -
                        lobjItemFac.FdecDeudaIvaCapital, "c")
                ' AVV Revisar
                'txtDeudaMora.Content = Format(lobjItemFac.FdecDeudaIntMora -
                '        lobjItemFac.FdecDeudaIvaInt, "c")
            End If
            SValideFraDscto()
            SValideItemDscto()
        End If
    End Sub
    Private Sub SRegTipoDscto()
        If MenuEstadoDscto = EnuEstadoDsctoDef.enuCreandoDscto Then
            SPuebleFras()
            SValideTipoDscto()
        End If
    End Sub
    Private Sub SRegValorDscto()
        If IsNumeric(txtValorDctoNuevo.Text) Then
            Dim ldecVlrDcto As Decimal = CType(txtValorDctoNuevo.Text, Decimal)
            txtValorDctoNuevo.Text = Format(ldecVlrDcto, "c")
        End If
        SValideDscto()
        MblnDejoUltimoControl = True
    End Sub
    Private Sub SRegDscto()
        SRegNroFactDscto()
        SRegValorDscto()
        SValideDscto()
    End Sub
    ' Medios de Pago
    Private Sub SRegValorRec()
        With MobjObjetoWin
            Dim ldecVlrReciboActual As Decimal = .ObjValor_RecDec.ObjValorPro
            .ObjValor_RecDec.ObjValorPro = txtValorRec.Text
            If .ObjValor_RecDec.BlnEsValido Then
                If ldecVlrReciboActual <> .ObjValor_RecDec.ObjValorPro Then
                    MblnPrimeraVezMediosPago = True
                End If
                If .ObjValor_RecDec.BlnEsValido AndAlso MblnPrimeraVezMediosPago Then
                    If chkAnticipo.IsChecked Then
                        .ObjValorAnticipoDec.ObjValorPro = .ObjValor_RecDec.ObjValorPro
                        SCreeNuevoMedioPago()
                    End If
                End If
                If MblnPrimeraVezMediosPago AndAlso Not chkAnticipo.IsChecked Then
                    .ObjValorAnticipoDec.ObjValorPro = .FdecAnticipo
                    If .ObjValorAnticipoDec.ObjValorPro > 0 Then
                        Dim lstrMens = "El Valor pagado de más se ha registrado como un Anticipo!"
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                End If
                txtValorAnt.Content = Format(.ObjValorAnticipoDec.ObjValorPro, "c")
                txtAnticipo.Content = txtValorAnt.Content
                SValideMedPago()
                If .ObjValor_RecDec.BlnEsValido Then
                    txtValorRec.Text = Format(MobjObjetoWin.ObjValor_RecDec.ObjValorPro, "c")
                    txtRecibido.Content = txtValorRec.Text
                    txtSaldo.Content = Format(.FdecSaldoRC, "c")
                    txtSaldoPen.Content = txtSaldo.Content
                    MobjObjetoWin.ObjSaldo_RecDec.ObjValorPro = txtSaldo.Content
                    SMuestreDescuentos()
                    txtValorMedPagoNuevo.Text = txtValorRec.Text
                    MobjObjetoWin.SVerfiqueDsctoPP()
                Else
                    SHabiliteCtlsMedPago()
                End If
                cboTipoMedPagoNuevo.Focus()
            End If
        End With
    End Sub
    Private Sub SRegTipoMedioPago()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando AndAlso
                MenuEstadoMedPago = EnuEstadoMedPagoDef.enuCreandoMedPago Then
            If Not IsNothing(MobjMedioPago) Then
                ObjHijoObjWin = MobjMedioPago
                MobjMedioPago.ObjIdTipoMedPago_MedPagoByt.ObjValorPro =
                        cboTipoMedPagoNuevo.SelectedIndex
            End If
            SValideMedPago()
            If cboTipoMedPagoNuevo.SelectedIndex = 0 AndAlso
                    MobjMedioPago.ObjIdTipoMedPago_MedPagoByt.ObjValorNuevo <> 0 Then
                SCanceleMedPago()
            Else
                If cboTipoMedPagoNuevo.SelectedIndex = 1 OrElse
                        cboTipoMedPagoNuevo.SelectedIndex = 2 Then
                    cboCuentaIngresoNuevo.SelectedIndex = 1
                ElseIf cboCuentaIngresoNuevo.SelectedIndex = 1 Then
                    cboCuentaIngresoNuevo.SelectedIndex = 0
                End If
            End If
        End If
    End Sub
    Private Sub SRegCtaIngreso()
        If Not IsNothing(MobjMedioPago) Then
            Dim lstrIdCtaCont = FstrCtaContabilidadIngresos()
            MobjMedioPago.ObjIdCtaContabIngresoStr.ObjValorPro = lstrIdCtaCont
        End If
        SValideMedPago()
    End Sub
    Private Sub SRegNroMedioPago()
        If Not IsNothing(MobjMedioPago) Then
            MobjMedioPago.ObjNumeroMedPagoStr.ObjValorPro = txtNroMedPagoNuevo.Text
        End If
        SValideMedPago()
    End Sub
    Private Sub SRegValorMedioPago()
        If Not IsNothing(MobjMedioPago) Then
            MobjMedioPago.ObjValor_MedPagoDec.ObjValorPro = txtValorMedPagoNuevo.Text
        End If
        If MobjMedioPago.ObjValor_MedPagoDec.BlnEsValido Then
            txtValorMedPagoNuevo.Text = Format(MobjMedioPago.ObjValor_MedPagoDec.ObjValorPro, "c")
        End If
        SValideMedPago()
        SMuestreMediosPago()
        MblnDejoUltimoControl = True
    End Sub
    Private Sub SRegMedPago()
        SRegTipoMedioPago()
        SRegCtaIngreso()
        SRegNroMedioPago()
        SRegValorMedioPago()
        SValideMedPago()
    End Sub
    ' Comentario
    Private Sub SRegComentario()
        MobjObjetoWin.ObjComentario_RecStr.ObjValorPro = txtComentario.Text
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Select Case True
                Case lelmElemento.Equals(bttCausarIntereses)
                    SCauseMora()
                    SHabiliteControlesEstado()
                Case lelmElemento.Equals(bttNuevoDscto)
                    SNuevoDscto()
                Case lelmElemento.Equals(bttCancelarDscto)
                    SCanceleDscto()
                Case lelmElemento.Equals(bttAceptarDscto)
                    SAcepteDscto()
                Case lelmElemento.Equals(bttEliminarDscto)
                    SElimineDscto()
                Case lelmElemento.Equals(bttNuevoMedPago)
                    SCreeNuevoMedioPago()
                Case lelmElemento.Equals(bttCancelarMedPago)
                    SCanceleMedPago()
                Case lelmElemento.Equals(bttAceptarMedPago)
                    SAcepteMedPago()
                Case lelmElemento.Equals(bttEliminarMedPago)
                    SElimineMedPago()
                Case lelmElemento.Equals(bttSiguiente)
                    SPaseAlSiguiente()
                Case lelmElemento.Equals(bttAnterior)
                    SPaseAlAnterior()
                Case lelmElemento.Equals(bttEncontrarCliente)
                    SBusqueCliente()
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
                    SEnvieEmailRecCaja()
                ElseIf lmnuOpcion.Name = "MnuExportarRecs" Then
                    SExporteRecibos()
                ElseIf lmnuOpcion.Name = "MnuMostrarLista" Then
                    If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                        SMuestreRecibos(True)
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
            Dim lctlControl As TextBox = lelmElemento
            lctlControl.SelectAll()
        ElseIf TypeOf lelmElemento Is Button Then
            If MblnDejoUltimoControl Then
                If lelmElemento.Equals(HbttCancelar) Then
                    HbttAceptar.Focus()
                ElseIf lelmElemento.Equals(bttCancelarDscto) Then
                    If bttAceptarDscto.IsEnabled Then bttAceptarDscto.Focus()
                ElseIf lelmElemento.Equals(bttCancelarMedPago) Then
                    If bttAceptarMedPago.IsEnabled Then bttAceptarMedPago.Focus()
                End If
            End If
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Dim lblnMostrarDatos = True
        If Not HblnSeEstaCerrando AndAlso Not HblnMostrandoDatos Then
            MblnDejoUltimoControl = False
            If TypeOf lelmElemento Is TextBox OrElse (TypeOf lelmElemento Is ComboBox AndAlso
                    lelmElemento.Name = "cboTipoMedPagoNuevo") Then
                Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                Try
                    GobjPanDat.SControleProcesoObj(True)
                    If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                        Select Case lelmElemento.Name
                            Case "txtIdCliente"
                                SRegCliente()
                                lblnMostrarDatos = False
                            Case "txtValorRec"
                                SRegValorRec()
                            Case "txtValorDctoNuevo"
                                SRegValorDscto()
                            Case "txtNroMedPagoNuevo"
                                SRegNroMedioPago()
                            Case "txtValorMedPagoNuevo"
                                SRegValorMedioPago()
                            Case "txtComentario"
                                SRegComentario()
                            Case "cboTipoMedPagoNuevo"
                                SRegTipoMedioPago()
                            Case Else
                                lblnMostrarDatos = False
                        End Select
                        If lblnMostrarDatos Then SMuestreDatos()
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
        If Not MblnPoblandoCbo Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If Not HblnSeEstaCerrando AndAlso Not HblnMostrandoDatos Then
                If TypeOf lelmElemento Is ComboBox Then
                    Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                    If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
                        Try
                            GobjPanDat.SControleProcesoObj(True)
                            Select Case lelmElemento.Name
                                Case "cboNroFactNuevo"
                                    SPuebleItemsFra()
                                    SRegNroFactDscto()
                                Case "cboItemFactNuevo"
                                    SRegNroFactDscto()
                                Case "cboTipoDsctoNuevo"
                                    SRegTipoDscto()
                                Case "cboTipoMedPagoNuevo"
                                    SRegTipoMedioPago()
                                Case "cboCuentaIngresoNuevo"
                                    SRegCtaIngreso()
                            End Select
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
                    Else
                        If lelmElemento.Name = "cboPref" Then
                            SInicialiceRec()
                            SMuestreDatos()
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub OnCambioFecha(sender As Object, e As RoutedEventArgs)
        If TypeOf sender Is DatePicker Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando AndAlso
                    Not HblnMostrandoDatos Then
                If MobjObjetoWin.ObjFechaRecDtm.ObjValorPro <> dtpFechaRec.SelectedDate AndAlso
                        dtpFechaRec.SelectedDate <> GCDTMFECHANULA Then
                    MdtmUltimaFechaIng = dtpFechaRec.SelectedDate
                    MobjObjetoWin.ObjFechaRecDtm.ObjValorPro = MdtmUltimaFechaIng
                End If
                SMuestreDatos()
                SProceseGrales()
                If MdecIntPorCausar > 0 Then
                    Dim lstrMens = "Aún no se han causado los Intereses de Mora a la Fecha!"
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            End If
        End If
    End Sub
    Private Sub ChkAnticipo_Click(sender As Object, e As RoutedEventArgs) Handles chkAnticipo.Click
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando AndAlso
                Not IsNothing(MobjCliente) Then
            Dim lstrMens = String.Empty
            MobjObjetoWin.BlnEsSoloAnticipo = chkAnticipo.IsChecked
            MobjObjetoWin.ObjIdPredioAgrupador_RecStr.SValide()
            If MobjObjetoWin.ObjIdCliente_RecDbl.BlnEsValido AndAlso
                    MobjObjetoWin.ObjIdPredioAgrupador_RecStr.BlnEsValido Then
                SPuebleLsbServicios(lstrMens)
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            End If
        End If
    End Sub
    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles _
            txtNroRecCaja.KeyDown
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If e.Key = Key.Return OrElse e.Key = Key.Tab Then
                SAbraRecCaja()
            End If
        End If
    End Sub
    Private Sub DgrDescuentos_SelectedCellsChanged(sender As Object,
            e As SelectedCellsChangedEventArgs) Handles dgrDescuentos.SelectedCellsChanged
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            If dgrDescuentos.Items.Count = 0 Then
                cboNroFactNuevo.SelectedIndex = 0
                cboTipoDsctoNuevo.SelectedIndex = 0
                txtValorDctoNuevo.Text = Format(0, "c")
            Else
                cboTipoDsctoNuevo.SelectedItem = txtTipoDcto.Text
                cboNroFactNuevo.SelectedItem = txtNroFactura.Text
                cboItemFactNuevo.SelectedItem = txtItemFactura.Text
                txtValorDctoNuevo.Text = txtValorDcto.Text
            End If
            SValideDscto()
        End If
    End Sub
    Private Sub DgrMediosPago_SelectedCellsChanged(sender As Object, e As SelectedCellsChangedEventArgs) Handles dgrMediosPago.SelectedCellsChanged
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If Not MblnEliminandoMedPago Then
                If Not IsNothing(dgrMediosPago.SelectedItem) Then
                    Dim ldrvMedioPago As DataRowView = dgrMediosPago.SelectedItem
                    Dim lentOrdinalMedPago As Integer = ldrvMedioPago("Ordinal")
                    Dim lcolMediosPago As Collection = MobjObjetoWin.ColMediosPago
                    MobjMedioPago = lcolMediosPago(lentOrdinalMedPago.ToString)
                    SMuestreMediosPago()
                End If
            End If
        End If
    End Sub
    Private Sub DgrMediosPagoNuevo_SelectedCellsChanged(sender As Object, e As SelectedCellsChangedEventArgs) _
            Handles dgrMediosPagoNuevo.SelectedCellsChanged
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            If Not MblnEliminandoMedPago Then
                If Not String.IsNullOrEmpty(txtOrdinal.Text) Then
                    Dim lentOrdinalMedPago As Integer = CType(txtOrdinal.Text, Integer)
                    Dim lcolMediosPago As Collection = MobjObjetoWin.ColMediosPago
                    If lcolMediosPago.Count > 0 Then
                        MobjMedioPago = lcolMediosPago(lentOrdinalMedPago.ToString)
                    End If
                End If
                    SMuestreMediosPago()
            End If
        End If
    End Sub
    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
                dgrNovedades.MouseRightButtonUp, dgrItems.MouseRightButtonUp,
                dgrDeuda.MouseRightButtonUp
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            Dim lstrNroFac As String, lstrPrefFac As String, lentIdFac As Integer
            ldrvFilaActual = ldgrActual.SelectedItem
            If ldrvFilaActual IsNot Nothing Then
                If ldgrActual.Name = "dgrNovedades" Then
                    Dim lenuTipoNovedad As EnuTipoNov = ldrvFilaActual(ClsIdTipoNovedadByt.SstrNombreCampoBd)
                    If lenuTipoNovedad = EnuTipoNov.EnuCrAntRec Then
                        SAbraAnticipo(MobjObjetoWin.ObjIdAnticipo_RecEnt.ObjValorPro)
                    Else
                        lstrPrefFac = ldrvFilaActual(ClsPrefijoFact_NovStr.SstrNombreCampoBd)
                        lentIdFac = ldrvFilaActual(ClsIdFactura_NovEnt.SstrNombreCampoBd)
                        SAbraFactura(lstrPrefFac, lentIdFac)
                    End If
                ElseIf ldgrActual.Name = "dgrItems" Then
                    lstrNroFac = ldrvFilaActual("NroFact")
                    If lstrNroFac = "0" Then
                        SAbraAnticipo(MobjObjetoWin.ObjIdAnticipo_RecEnt.ObjValorPro)
                    ElseIf lstrNroFac = "Todas" Then
                        '
                    Else
                        lstrPrefFac = ClsPanorama.FstrPrefijoDcto(lstrNroFac)
                        lentIdFac = ClsPanorama.FentIdDcto(lstrNroFac)
                        SAbraFactura(lstrPrefFac, lentIdFac)
                    End If
                ElseIf ldgrActual.Name = "dgrDeuda" Then
                    If Not IsNothing(ldrvFilaActual) Then
                        lstrNroFac = ldrvFilaActual("NroFactura")
                        If lstrNroFac <> "Todas" Then
                            lstrPrefFac = ClsPanorama.FstrPrefijoDcto(lstrNroFac)
                            lentIdFac = ClsPanorama.FentIdDcto(lstrNroFac)
                            SAbraFactura(lstrPrefFac, lentIdFac)
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub Ctl_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles _
            txtIdAnticipo.MouseDoubleClick, txtNroNotaRCr.MouseDoubleClick, txtNroNotasCr.MouseDoubleClick
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Dim lstrMens = String.Empty
        If lelmElemento.Name = "txtIdAnticipo" Then
            If Not String.IsNullOrEmpty(txtIdAnticipo.Content) Then
                SAbraAnticipo(MobjObjetoWin.ObjIdAnticipo_RecEnt.ObjValorPro)
            Else
                lstrMens = "Este Recibo de Caja no generó Anticipo!"
            End If
        ElseIf lelmElemento.Name = "txtNroNotaRCr" Then
            If Not String.IsNullOrEmpty(txtNroNotaRCr.Content) Then
                Dim lstrPrefNotaRRC As String = ClsPanorama.FstrPrefijoDcto(txtNroNotaRCr.Content)
                Dim lentIdNotaRRC As Integer = ClsPanorama.FentIdDcto(txtNroNotaRCr.Content)
                SAbraNotaRCr(lstrPrefNotaRRC, lentIdNotaRRC)
            Else
                lstrMens = "Este Recibo de Caja no ha sido reversado!"
            End If
        ElseIf lelmElemento.Name = "txtNroNotasCr" Then
            If Not String.IsNullOrEmpty(txtNroNotasCr.Content) Then
                Static i As Integer = -1
                i += 1
                Dim lstrNroNCr As String
                If i > MobjObjetoWin.ObjIdNotasCrStr.ToString.Split(",").GetUpperBound(0) Then
                    i = 0
                End If
                lstrNroNCr = MobjObjetoWin.ObjIdNotasCrStr.ToString.Split(",")(i)
                Dim lstrPrefNotaCr As String = ClsPanorama.FstrPrefijoDcto(lstrNroNCr)
                Dim lentIdNotaCr As Integer = ClsPanorama.FentIdDcto(lstrNroNCr)
                SAbraNotaCr(lstrPrefNotaCr, lentIdNotaCr)
            Else
                lstrMens = "Este Recibo de Caja no generó Notas Crédito!"
            End If
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub Lsb_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles _
            lsbPrediosAgru.SelectionChanged, lsbServicios.SelectionChanged
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is ListBox Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
                If Not HblnMostrandoDatos Then
                    Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                    Try
                        GobjPanDat.SControleProcesoObj(True)
                        If lelmElemento.Name = "lsbPrediosAgru" Then
                            SRegPrediosAgr(lstrMens)
                        Else
                            SRegServicios()
                        End If
                        SHabiliteControlesEstado()
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
                            If MdecIntPorCausar > 0 Then
                                lstrMens = "Aún no se han causado los Intereses de Mora a la Fecha!"
                            End If
                            GobjPanDat.SControleProcesoObj(False)
                            SValide()
                            If Not String.IsNullOrEmpty(lstrMens) Then
                                SLevanteEveNoti(lstrMens, String.Empty, 0,
                                        EnuSeveridadNot.EnuInformacion)
                            End If
                        Else
                            GobjPanDat.SControleProcesoObj(False, True)
                            SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                        End If
                    End Try
                End If
            End If
        End If
    End Sub
    Private Sub DgrRC_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles dgrRecsCaja.MouseDoubleClick
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            Dim lstrPrefijo As String
            Dim lentIdRecCaja As Integer
            ldrvFilaActual = ldgrActual.SelectedItem
            If Not IsNothing(ldrvFilaActual) AndAlso ldgrActual.SelectedIndex >= 0 Then
                lstrPrefijo = ldrvFilaActual("Prefijo")
                lentIdRecCaja = ldrvFilaActual("IdReciboCaja")
                cboPref.SelectedItem = lstrPrefijo
                txtNroRecCaja.Text = lentIdRecCaja
                SAbraRecCaja()
            End If
        End If
        SMuestreRecibos(False)
    End Sub
    Private Sub DgrRC_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles dgrRecsCaja.MouseRightButtonUp
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            Dim lstrPrefijo As String
            Dim lentIdRecCaja As Integer
            ldrvFilaActual = ldgrActual.SelectedItem
            If Not IsNothing(ldrvFilaActual) AndAlso ldgrActual.SelectedIndex >= 0 Then
                lstrPrefijo = ldrvFilaActual("Prefijo")
                lentIdRecCaja = ldrvFilaActual("IdReciboCaja")
                cboPref.SelectedItem = lstrPrefijo
                txtNroRecCaja.Text = lentIdRecCaja
                SAbraRecCaja()
            End If
        End If
        SMuestreRecibos(False)
    End Sub
#End Region
End Class