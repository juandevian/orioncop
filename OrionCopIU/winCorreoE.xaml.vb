Imports Microsoft.Win32
Imports System.Threading
Imports System.ComponentModel
Public Class WinCorreoE
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        EnuTipoCorreo = 0
        EnuArch
        EnuFacMes
        EnuCliente
        EnuPreAgr
        EnuDoc
        EnuMens
        EnuAsun
        EnuFecIni
        EnuFecFin
        EnuDiasCobPer
        EnuConexion
    End Enum
#End Region
    ' Segundo plano
    Private WithEvents MbgwCorreo As New BackgroundWorker
    'Variables
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomECorreo
    Private MblnPoblandoCombo As Boolean = False
    Private MobjCliente As ClsCliente = Nothing
    Private MblnEnviandoEmail As Boolean = False
    Private MentDiasCobro As Integer = GobjParametros.ObjDiasParaPersuasivoShr.ObjValorPro
    '
    Private MobjObjetoWin As ClsCorreoE = Nothing
    Private WithEvents MobjReportes As New ClsRepOrionCop(GCOBJREGISTRO)
    Private MblnCancelando As Boolean = False
    Private MdtbCorreos As DataTable = Nothing
    Private MblnHayDocNoEnviados As Boolean = False
    Private MblnHayConeccion As Boolean = False
    Private MblnEnvioOk As Boolean = False
    Private MblnTimeOut As Boolean = False
    Private MblnEnvioCorreo As Boolean = False
    Private MblnDesdeDoc As Boolean = False ' Indica si el correo se esta enviando desde la ventana
    Private ReadOnly MenuTamanoIcono As EnuTamanoIconos
    ' del documento
    Private MnuCuentaOrigenEmail As MenuItem = Nothing
    Private MblnBuscarHistoClie As Boolean = False
    Private ReadOnly MshrIdCarpeta As Short = GshrIdCarpeta
    Private ReadOnly MshrIdCentroUtil As Short = GshrIdCentroUtil
    '
    Private ReadOnly MobjCentroUtilActual As ClsCentroUtilidad =
            GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
    Private ReadOnly MobjTerceroCentrUtil As ClsTercero =
            MobjCentroUtilActual.ObjTerceroCentroUtilidad
    Private ReadOnly MstrServidor As String = MobjCentroUtilActual.ObjServidorSmtpStr.ObjValorPro
    Private ReadOnly MblnHabilitarSSL As Boolean =
            MobjCentroUtilActual.ObjHabilitarSslBln.ObjValorPro
    Private ReadOnly MentPuerto As Integer =
            MobjCentroUtilActual.ObjPuertoHostShr.ObjValorPro
    Private ReadOnly MblnRequiereAutent As Boolean =
            MobjCentroUtilActual.ObjRequiereAutenticacionBln.ObjValorPro
    Private MstrNotifica As String = String.Empty
#End Region

#Region "Constructor"
    Friend Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuECorreo
        MbgwCorreo.WorkerReportsProgress = True
        MbgwCorreo.WorkerSupportsCancellation = True
        MenuTamanoIcono = GenuTamanoIcono
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lstrMens = String.Empty
        SAdicioneControlRestringido(bttEncontrarCli)
        SAdicioneControlRestringido(bttEncontrarCliHist)
        SAdicioneControlRestringido(txtIdClienteHis)
        SCargueForma(EnuElementosAdicionalesDef.None, 12,
                     Nothing, Nothing, True)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        If FblnPuedeEnviarEmail(lstrMens) Then
            SCree()
            SRegistreDoc()
        Else
            Exit Sub
        End If
        SVisibiliceCtls()
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
        SValide()
        If MobjObjetoWin.EnuTipoCorreo = EnuTipoCorreoE.None Then
            SVerifiquePorEnviar()
        End If
        MblnDesdeDoc = False
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
        ObjObjetoWin = GobjParametros
        If IsNothing(MobjObjetoWin) Then
            MobjObjetoWin = New ClsCorreoE
            MobjObjetoWin.FblnEsValidoTipoCorreo(EnuTipoCorreoE.None)
        End If
    End Sub

    Protected Overrides Sub SInicialiceControles()
        lblCuentaCorreoOri.Content = "Cuenta origen de Correo : " &
                GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjEmailOrigenStr.ObjValorPro
        StcValidaControl(EnuValidEntrada.EnuTipoCorreo) = lblTipoDoc
        StcValidaControl(EnuValidEntrada.EnuArch) = lblSelArch
        StcValidaControl(EnuValidEntrada.EnuFacMes) = chkSeguro
        StcValidaControl(EnuValidEntrada.EnuCliente) = lblCliente
        StcValidaControl(EnuValidEntrada.EnuPreAgr) = lblPredioAgr
        StcValidaControl(EnuValidEntrada.EnuDoc) = lblDoc
        StcValidaControl(EnuValidEntrada.EnuMens) = lblMensaje
        StcValidaControl(EnuValidEntrada.EnuAsun) = lblAsunto
        StcValidaControl(EnuValidEntrada.EnuFecIni) = lblFecIni
        StcValidaControl(EnuValidEntrada.EnuFecFin) = lblFecFin
        StcValidaControl(EnuValidEntrada.EnuDiasCobPer) = lblComDias
        '
        SPuebleCombos()
        txtIdClienteHis.Text = My.Resources.Todos
        dtpDesde.SelectedDate =
                GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        dtpHasta.SelectedDate = Date.Today
        dtpFechaDesde.SelectedDate =
                GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        dtpFechaHasta.SelectedDate = Date.Today
        cboTipoCorreoH.SelectedIndex = 0
        '
        HbttAceptar.TabIndex = 30
        HbttCancelar.TabIndex = 31
    End Sub

    Private Sub SVacie()
        If Not HblnCargandoForma Then
            MblnPoblandoCombo = True
            txtArchivo.Text = String.Empty
            txtIdCliente.Text = String.Empty
            txtNombreCliente.Content = String.Empty
            cboPredioAgr.Items.Clear()
            cboNroDoc.Items.Clear()
            chkSeguro.IsChecked = False
            dtpFecIni.SelectedDate = GCDTMFECHANULA
            dtpFecFin.SelectedDate = GCDTMFECHANULA
            pgbCorreo.Value = 0
            txtProceso.Content = "0%"
            MblnCancelando = False
            MblnPoblandoCombo = False
            If Not MblnDesdeDoc Then
                MobjObjetoWin.SVacie()
            End If
        End If
    End Sub

    Protected Overrides Sub SMuestreDatos()
        If Not HblnCargandoForma Then
            With MobjObjetoWin
                cboTipoCorreo.SelectedIndex = .EnuTipoCorreo
                SVisibiliceCtls()
                Select Case .EnuTipoCorreo
                    Case EnuTipoCorreoE.EnuSoloMens
                        txtMensaje.Text = MobjObjetoWin.StrMensaje
                    Case EnuTipoCorreoE.EnuArchExt
                        txtArchivo.Text = .StrArchivoExterno
                    Case EnuTipoCorreoE.EnuCobroPers
                        txtDiasVen.Text = .EntDiasCobroPers
                    Case EnuTipoCorreoE.EnuFac, EnuTipoCorreoE.EnuRC, EnuTipoCorreoE.EnuNCR,
                            EnuTipoCorreoE.EnuNDB, EnuTipoCorreoE.EnuNAA
                        txtIdCliente.Text = .DblIdCliente
                        txtNombreCliente.Content = .StrNombreCliente
                        cboPredioAgr.SelectedItem = FstrPredAgru()
                        cboNroDoc.SelectedItem = .StrNroDocumento
                        txtAsunto.Text = .StrAsunto
                        txtMensaje.Text = .StrMensaje
                    Case EnuTipoCorreoE.EnuFactAuto
                        chkSeguro.IsChecked = MobjObjetoWin.BlnFactAuto
                    Case EnuTipoCorreoE.EnuRecibos
                        dtpFecIni.SelectedDate = .DtmFechaIni
                        dtpFecFin.SelectedDate = .DtmFechaFin
                End Select
                SValide()
            End With
        End If
    End Sub

    Protected Overrides Sub SValide()
        Dim lstrMens = String.Empty
        Dim lenuIdMens As EnuIdMens = EnuIdMens.None
        For i = 0 To 11
            StcValidValido(i) = True
        Next
        With MobjObjetoWin
            StcValidValido(EnuValidEntrada.EnuConexion) = .FblnHayConInternet(lstrMens)
            If String.IsNullOrEmpty(lstrMens) Then
                StcValidValido(EnuValidEntrada.EnuTipoCorreo) = .FblnEsValidoTipoCorreo(
                    MobjObjetoWin.EnuTipoCorreo)
                StcValidValido(EnuValidEntrada.EnuAsun) = .FblnEsValidoAsunto(txtAsunto.Text,
                        lstrMens)
            End If
            If String.IsNullOrEmpty(lstrMens) Then
                StcValidValido(EnuValidEntrada.EnuMens) = .FblnEsValidoMensaje(txtMensaje.Text,
                        lstrMens)
            End If
            If String.IsNullOrEmpty(lstrMens) Then
                Select Case cboTipoCorreo.SelectedIndex
                    Case EnuTipoCorreoE.EnuArchExt
                        StcValidValido(EnuValidEntrada.EnuArch) = .FblnEsValidoArchivoExt(
                                txtArchivo.Text, lstrMens)
                    Case EnuTipoCorreoE.EnuFactAuto
                        StcValidValido(EnuValidEntrada.EnuFacMes) =
                                .FblnEsValidoFactAuto(chkSeguro.IsChecked)
                    Case EnuTipoCorreoE.EnuFac, EnuTipoCorreoE.EnuRC, EnuTipoCorreoE.EnuNAA,
                            EnuTipoCorreoE.EnuNCR, EnuTipoCorreoE.EnuNDB
                        Dim lstrIdpredAgru = "***"
                        StcValidValido(EnuValidEntrada.EnuCliente) =
                                .FblnEsValidoIdCliente(txtIdCliente.Text)
                        If Not IsNothing(cboPredioAgr.SelectedItem) Then
                            lstrIdpredAgru = cboPredioAgr.SelectedItem
                        End If
                        StcValidValido(EnuValidEntrada.EnuPreAgr) =
                                .FblnEsValidoIdPreAgr(lstrIdpredAgru)
                        StcValidValido(EnuValidEntrada.EnuDoc) =
                                .FblnEsValidoNroDoc(cboNroDoc.SelectedItem, lstrMens)
                        If String.IsNullOrEmpty(lstrMens) Then
                            StcValidValido(EnuValidEntrada.EnuAsun) =
                                    .FblnEsValidoAsunto(txtAsunto.Text, lstrMens)
                        End If
                        If String.IsNullOrEmpty(lstrMens) Then
                            StcValidValido(EnuValidEntrada.EnuMens) =
                                    .FblnEsValidoMensaje(txtMensaje.Text, lstrMens)
                        End If
                    Case EnuTipoCorreoE.EnuCobroPers
                        StcValidValido(EnuValidEntrada.EnuMens) =
                                .FblnEsValidoMensajeCobroPer(txtMensaje.Text, lstrMens)
                        If String.IsNullOrEmpty(lstrMens) Then
                            StcValidValido(EnuValidEntrada.EnuDiasCobPer) =
                                    .FblnEsValidoDiasVen(txtDiasVen.Text, lstrMens)
                        End If
                End Select
            End If
        End With
        '
        SHabiliteBotonesTlb()
        If FblnEstanTodosBien() Then
            HbttAceptar.IsEnabled = True
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, lenuIdMens, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Protected Overrides Sub SRegistre()
        '
        If cboTipoCorreo.SelectedIndex = EnuTipoCorreoE.EnuCobroPers Then
            SEscribaMensCobroPresuIni()
        End If
    End Sub

    Protected Overrides Sub SFinaliceOperacion()
        SEstablezcaWinConsultando()
        If MblnEnviandoEmail Then
            HbttCancelar.Content = My.Resources.Cancelar
        Else
            HbttCancelar.Content = My.Resources.CerrarBtn
        End If
    End Sub

    Protected Overrides Sub SConfigureMenuesPropios()
        Dim lsepSeparador As New Separator
        MnuCuentaOrigenEmail = FmnuiMenuItem("MnuCuentaOrigenEmail", "_Cuenta Origen Correo",
                 "RecMnuItemSec", "")
        HmnuAcciones.Items.Add(lsepSeparador)
        HmnuAcciones.Items.Add(MnuCuentaOrigenEmail)
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    Protected Overrides Sub SCree()
        EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando
        SHabiliteControles(True)
        If Not IsNothing(HbttCancelar) Then
            HbttCancelar.Content = My.Resources.Cancelar
        End If
        SHabiliteWin(True)
    End Sub

    Protected Overrides Sub SGuarde()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        SRegistre()
        SValide()
        If FblnEstanTodosBien() Then
            Try
                GobjPanDat.SControleProcesoObj(True)
                HbttAceptar.IsEnabled = False
                If MobjObjetoWin.EnuTipoCorreo = EnuTipoCorreoE.EnuNAA OrElse
                    MobjObjetoWin.EnuTipoCorreo = EnuTipoCorreoE.EnuNCR OrElse
                    MobjObjetoWin.EnuTipoCorreo = EnuTipoCorreoE.EnuNDB OrElse
                    MobjObjetoWin.EnuTipoCorreo = EnuTipoCorreoE.EnuRC OrElse
                    MobjObjetoWin.EnuTipoCorreo = EnuTipoCorreoE.EnuFac Then
                    HbttCancelar.IsEnabled = False
                End If
                MbgwCorreo.RunWorkerAsync()
                MblnEnviandoEmail = True
                SHabiliteControles(False)
                lblnNoHayError = True
            Catch ex As ErrorInesperadoPanDatException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As ErrorInesperadoPanLException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As ArgumentoInvalidoPanException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As PanDatException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As PanLException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As DataException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As Exception
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Finally
                If lblnNoHayError Then
                    GobjPanDat.SControleProcesoObj(False)
                    If String.IsNullOrEmpty(lstrMens) Then
                        lstrMens = "Enviando Mensajes!"
                    End If
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Else
                    GobjPanDat.SControleProcesoObj(False, True)
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
                MdtbCorreos = Nothing
                SFinaliceOperacion()
            End Try
        End If
    End Sub

    Protected Overrides Sub SCancele()
        If MbgwCorreo.IsBusy Then
            If MsgBox("Desea detener el envio por Correo?", vbYesNo, "Cancelar acción") = vbYes Then
                MbgwCorreo.CancelAsync()
                MblnCancelando = True
            End If
        Else
            SCerrarClic()
        End If
    End Sub

    Friend WriteOnly Property ObjCorreoE As ClsCorreoE
        Set(value As ClsCorreoE)
            MobjObjetoWin = value
        End Set
    End Property

    Protected Overrides Function FblnNotificaOk(aenuIdMensNot As EnuIdMens) As Boolean
        Dim lblnOk = True, lstrMens = String.Empty
        Select Case aenuIdMensNot
            Case EnuIdMens.EnuArchivo
                lblnOk = MobjObjetoWin.FblnEsValidoArchivoExt(txtArchivo.Text, lstrMens)
            Case EnuIdMens.EnuCliente
                lblnOk = MobjObjetoWin.FblnEsValidoIdCliente(txtIdCliente.Text)
            Case EnuIdMens.EnuAsunto
                lblnOk = MobjObjetoWin.FblnEsValidoAsunto(txtAsunto.Text, lstrMens)
            Case EnuIdMens.EnuMensaje
                lblnOk = MobjObjetoWin.FblnEsValidoMensaje(txtMensaje.Text, lstrMens)
            Case EnuIdMens.EnuDoc
                lblnOk = MobjObjetoWin.FblnEsValidoNroDoc(cboNroDoc.SelectedItem, lstrMens)
            Case EnuIdMens.EnuFecIni
                lblnOk = MobjObjetoWin.FblnEsValFecIni(dtpFecIni.SelectedDate, lstrMens)
            Case EnuIdMens.EnuFecFin
                lblnOk = MobjObjetoWin.FblnEsValFecFin(dtpFecFin.SelectedDate, lstrMens)
        End Select
        Return lblnOk
    End Function
#End Region

#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            If BlnBusquedaOk AndAlso StrResultadoBusqueda.Length > 0 Then
                If MblnBuscarHistoClie Then
                    txtIdClienteHis.Text = StrResultadoBusqueda
                Else
                    txtIdCliente.Text = StrResultadoBusqueda
                    SRegistreCliente()
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
        SDefineBusquedaPredioAgr_Prop()
        SDefineBusquedaPredioAgr_Arren()
        SDefineBusquedaNombreCompleto()
        SDefineBusquedaPrimerApell()
        Return True
    End Function
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
        Dim lstrFiltro As String = "P." & OPT.OrionP.PanL.ClsIdCarpetaShr.SstrNombreCampoBd & " = " &
                MshrIdCarpeta & " AND P." & OPT.OrionP.PanL.ClsIdCentroUtilShr.SstrNombreCampoBd &
                " = " & MshrIdCentroUtil
        lstrFiltro &= " AND " & ClsIdClienteArrendatarioDbl.SstrNombreCampoBd & " > 0 AND " &
                lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador - Arrendatario", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, False)
    End Sub
    Private Sub SDefineBusquedaNombreCompleto()
        Dim lstrCamposMostrar As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                                 ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdClienteDbl.SstrNombreCampoBd
        Dim lstrTabla As String = ClsCliente.SstrNombreTabla
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND " & lstrCampoBusqueda &
                "<> ''"
        HwinBusqueda.SDefinaBusqueda("Nombre Completo", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
    Private Sub SDefineBusquedaPrimerApell()
        Dim lstrTablaPri As String = ClsCliente.SstrNombreTabla
        Dim lstrTablaSec As String = ClsTercero.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamposTabSec As String() = {"CONCAT(S." & ClsApellidoPrimeroStr.SstrNombreCampoBd & ", " &
                "' '" & ", S." & ClsNombrePrimeroStr.SstrNombreCampoBd & ")" & " AS ApellidoNombre"}
        Dim lstrCamTabPriRel As String() = {ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {ClsIdTerceroDbl.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = "ApellidoNombre"
        Dim lstrCampoRetornar As String = ClsIdClienteDbl.SstrNombreCampoBd
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion & " AND S." &
                ClsNombrePrimeroStr.SstrNombreCampoBd & " <> ''"
        HwinBusqueda.SDefinaBusqueda("Primer Apellido", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, True)
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SAsigneValidoTodo(ablnOk As Boolean)
        For i As Integer = 0 To 9
            StcValidValido(i) = ablnOk
        Next
    End Sub

    Private Function FblnPuedeEnviarEmail(ByRef astrMens As String) As Boolean
        Dim lblnPuede As Boolean
        lblnPuede = ClsPanorama.FblnEmailsHabilitado
        If Not lblnPuede Then
            astrMens = "No tiene habilitado el Módulo de Correo Electrónico!"
        End If
        If lblnPuede Then
            GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.FblnExisteCtaCorreoValida()
            If Not lblnPuede Then
                astrMens = "La Cuenta de Correo no esta debidamente parametrizada!"
            End If
        End If
        If lblnPuede Then
            lblnPuede = Not GblnPosteando
            If Not lblnPuede Then
                astrMens = "No es posible enviar Emails mientras haya un proceso " &
                            "de eFac activo!"
            End If
        End If
        Return lblnPuede
    End Function

    Private Sub SPuebleCombos()
        MblnPoblandoCombo = True
        With cboTipoCorreo
            .Items.Clear()
            .Items.Add(My.Resources.Ninguno)
            .Items.Add(My.Resources.SolComen)
            .Items.Add(My.Resources.ArchExt)
            .Items.Add(My.Resources.FactAuto)
            .Items.Add(My.Resources.Fact)
            .Items.Add(My.Resources.RecCaja)
            .Items.Add(My.Resources.NCR)
            .Items.Add(My.Resources.NDB)
            .Items.Add(My.Resources.NAA)
            .Items.Add(My.Resources.RecCajaFec)
            .Items.Add(My.Resources.CobPer)
        End With
        cboTipoCorreo.SelectedIndex = 0
        With cboTipoCorreoH
            .Items.Clear()
            .Items.Add(My.Resources.Ninguno)
            .Items.Add(My.Resources.SolComen)
            .Items.Add(My.Resources.ArchExt)
            .Items.Add(My.Resources.FactAuto)
            .Items.Add(My.Resources.Fact)
            .Items.Add(My.Resources.RecCaja)
            .Items.Add(My.Resources.NCR)
            .Items.Add(My.Resources.NDB)
            .Items.Add(My.Resources.NAA)
            .Items.Add(My.Resources.RecCajaFec)
            .Items.Add(My.Resources.CobPer)
        End With
        cboTipoCorreoH.SelectedIndex = 0

        MblnPoblandoCombo = False
    End Sub

    Private Sub SSeleccioneArchivo()
        Dim lofdOrigenDatos As New OpenFileDialog With {
            .DefaultExt = ".pdf",
            .Filter = "pdf|*.pdf|Documento Acces|*.doc;*.docx;*.xls;*.xlsx;*.mdb|Todos los Archivos|*.*",
            .InitialDirectory = GstrTrayReportes
        }
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Dim lblnOk As Boolean = lofdOrigenDatos.ShowDialog
            If lblnOk Then
                txtArchivo.Text = lofdOrigenDatos.FileName
                txtArchivo.Focus()
                txtArchivo.CaretIndex = txtArchivo.Text.ToString().Length
            End If
            lblnNoHayError = True
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As DataException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                StcValidValido(EnuValidEntrada.EnuArch) =
                        MobjObjetoWin.FblnEsValidoArchivoExt(txtArchivo.Text, lstrMens)
                SMuestreDatos()
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub

    Private Sub SVisibiliceCtls()
        cnvArch.Visibility = Visibility.Collapsed
        cnvDocs.Visibility = Visibility.Collapsed
        cnvFrasMes.Visibility = Visibility.Collapsed
        cnvRecs.Visibility = Visibility.Collapsed
        cnvCobroPer.Visibility = Visibility.Collapsed
        If MobjObjetoWin.FblnEsValidoTipoCorreo(cboTipoCorreo.SelectedIndex) Then
            Select Case MobjObjetoWin.EnuTipoCorreo
                Case EnuTipoCorreoE.EnuArchExt
                    cnvArch.Visibility = Visibility.Visible
                Case EnuTipoCorreoE.EnuFactAuto
                    cnvFrasMes.Visibility = Visibility.Visible
                    txtProceso.Visibility = Visibility.Visible
                Case EnuTipoCorreoE.EnuRecibos
                    cnvRecs.Visibility = Visibility.Visible
                    txtProceso.Visibility = Visibility.Visible
                Case EnuTipoCorreoE.EnuCobroPers
                    cnvCobroPer.Visibility = Visibility.Visible
                Case EnuTipoCorreoE.None, EnuTipoCorreoE.EnuSoloMens
                    '
                Case Else
                    cnvDocs.Visibility = Visibility.Visible
            End Select
        End If
    End Sub

    Private Sub SRegistreCliente()
        StcValidValido(EnuValidEntrada.EnuCliente) =
                MobjObjetoWin.FblnEsValidoIdCliente(txtIdCliente.Text)
        MobjCliente = MobjObjetoWin.ObjCliente
        If Not IsNothing(MobjCliente) AndAlso MobjCliente.BlnExiste Then
            txtNombreCliente.Content = MobjCliente.ObjNombreCompletoStr.ObjValorPro
        End If
        SPuebleCboPredAgru()
    End Sub

    Private Sub SRegistreDoc()
        ' Cuando se envia el correo desde la ventana de un documento el tipo de correo (EnuTipoCorreo)
        ' en la clase tiene un valor diferente a None, lo que no sucede cuando se abre la ventana de
        ' correo desde elmenú principal
        With MobjObjetoWin
            If .EnuTipoCorreo = EnuTipoCorreoE.EnuFac OrElse .EnuTipoCorreo = EnuTipoCorreoE.EnuRC OrElse
                    .EnuTipoCorreo = EnuTipoCorreoE.EnuNCR OrElse .EnuTipoCorreo = EnuTipoCorreoE.EnuNDB OrElse
                    .EnuTipoCorreo = EnuTipoCorreoE.EnuNAA Then
                MblnDesdeDoc = True
                MobjCliente = MobjObjetoWin.ObjCliente
                If Not IsNothing(MobjCliente) Then
                    txtIdCliente.Text = MobjCliente.ObjIdClienteDbl.ToString
                    SPuebleCboPredAgru()
                    SMuestreDatos()
                End If
            ElseIf .EnuTipoCorreo = EnuTipoCorreoE.EnuFactAuto Then
                MblnDesdeDoc = True
                chkSeguro.IsChecked = True
                SMuestreDatos()
                chkSeguro.Focus()
            ElseIf .EnuTipoCorreo = EnuTipoCorreoE.None Then
                SAsigneValidoTodo(False)
                cboTipoCorreo.Focus()
            End If
        End With
    End Sub

    Private Sub SPuebleCboPredAgru()
        MblnPoblandoCombo = True
        cboPredioAgr.Items.Clear()
        cboNroDoc.Items.Clear()
        Dim lstrPrediosAgr = FarlPrediosAgrup()
        cboPredioAgr.Items.Add(My.Resources.Ninguno)
        If Not IsNothing(lstrPrediosAgr) AndAlso lstrPrediosAgr.Count > 0 Then
            For Each lstrPredAgr As String In lstrPrediosAgr
                If String.IsNullOrEmpty(lstrPredAgr) Then
                    cboPredioAgr.Items.Add(GCSTRSINPA)
                Else
                    If lstrPredAgr.Contains(",") Then
                        Dim lstrIdPrediosAgr As String() = lstrPredAgr.Split(",")
                        For Each lstrIdpreAgr As String In lstrIdPrediosAgr
                            If Not cboPredioAgr.Items.Contains(lstrIdpreAgr) Then
                                cboPredioAgr.Items.Add(lstrIdpreAgr)
                            End If
                        Next
                    Else
                        If Not cboPredioAgr.Items.Contains(lstrPredAgr) Then
                            cboPredioAgr.Items.Add(lstrPredAgr)
                        End If
                    End If
                End If
            Next
        End If
        cboPredioAgr.SelectedIndex = -1
        MblnPoblandoCombo = False
        If cboPredioAgr.Items.Count > 1 Then
            Dim lstrIdpreAgr = FstrPredAgru()
            If Not String.IsNullOrEmpty(lstrIdpreAgr) AndAlso
                    cboPredioAgr.Items.Contains(lstrIdpreAgr) Then
                cboPredioAgr.SelectedItem = lstrIdpreAgr
            Else
                cboPredioAgr.SelectedIndex = 0
            End If
        Else
            cboPredioAgr.SelectedIndex = 0
        End If
        If cboPredioAgr.Items.Count > 2 Then
            Dim lstrMens = "El presente Cliente tiene cuentas con más de un Predio!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Private Sub SPuebleDocs()
        MblnPoblandoCombo = True
        GobjPanDat.SControleProcesoObj(True)
        cboNroDoc.Items.Clear()
        cboNroDoc.Items.Add(My.Resources.Ninguno)
        SPuebleCboDoc()
        GobjPanDat.SControleProcesoObj(False)
        MblnPoblandoCombo = False
        If cboNroDoc.Items.Count > 1 Then
            If Not String.IsNullOrEmpty(MobjObjetoWin.StrNroDocumento) AndAlso
                    cboNroDoc.Items.Contains(MobjObjetoWin.StrNroDocumento) Then
                cboNroDoc.SelectedItem = MobjObjetoWin.StrNroDocumento
            Else
                cboNroDoc.SelectedIndex = 0
            End If
        Else
            cboNroDoc.SelectedIndex = 0
        End If
    End Sub

    Private Sub SPuebleCboDoc()
        If MobjObjetoWin.EnuTipoCorreo > EnuTipoCorreoE.None Then
            Dim ldtbDocs As DataTable
            Select Case MobjObjetoWin.EnuTipoCorreo
                Case EnuTipoCorreoE.EnuFac
                    ldtbDocs = FdtbFras()
                Case EnuTipoCorreoE.EnuRC
                    ldtbDocs = FdtbRsC()
                Case EnuTipoCorreoE.EnuNCR
                    ldtbDocs = FdtbNCR()
                Case EnuTipoCorreoE.EnuNDB
                    ldtbDocs = FdtbND()
                Case EnuTipoCorreoE.EnuNAA
                    ldtbDocs = FdtbNAA()
                Case Else
                    Throw New ErrorInesperadoPanLException("Tipo doc no corresponde!")
            End Select
            Dim lstrNroDoc As String, lstrPref As String, lentIdDoc As Integer, i = 0
            For Each ldrwDoc As DataRow In ldtbDocs.Rows
                i += 1
                lstrNroDoc = String.Empty
                lstrPref = ClsPanorama.FobjValorCampo(ldrwDoc(0), EnuTipoValor.EnuString)
                lentIdDoc = ClsPanorama.FobjValorCampo(ldrwDoc(1), EnuTipoValor.EnuInteger)
                lstrNroDoc = ClsPanorama.FstrNumeroDcto(lstrPref, lentIdDoc)
                cboNroDoc.Items.Add(lstrNroDoc)
                If i = 30 Then Exit For
            Next
        End If
    End Sub

    Private Function FdtbFras() As DataTable
        Dim lstrTabla = ClsFactura.SstrNombreTabla
        Dim lstrCamposSelect = {ClsPrefijo_FactStr.SstrNombreCampoBd,
                ClsIdFacturaEnt.SstrNombreCampoBd, ClsFechaFacturaDtm.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsFechaFacturaDtm.SstrNombreCampoBd, "DESC"},
                {ClsPrefijo_FactStr.SstrNombreCampoBd, "ASC"},
                {ClsIdFacturaEnt.SstrNombreCampoBd, "DESC"}}
        Dim lstrPreAgr As String = cboPredioAgr.SelectedItem
        If lstrPreAgr = GCSTRSINPA Then
            lstrPreAgr = String.Empty
        End If
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdCliente_FactDbl.SstrNombreCampoBd &
                " = " & txtIdCliente.Text & " AND " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd &
                " = '" & lstrPreAgr & "'"
        Dim ldtbFras As DataTable = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrIndice,
                lstrFiltro)
        Return ldtbFras
    End Function

    Private Function FdtbRsC() As DataTable
        Dim lstrTabla = ClsReciboCaja.SstrNombreTabla
        Dim lstrCamposSelect = {ClsPrefijo_RecStr.SstrNombreCampoBd,
                ClsIdRecCajaEnt.SstrNombreCampoBd, ClsFechaRecDtm.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsFechaRecDtm.SstrNombreCampoBd, "DESC"},
                {ClsPrefijo_RecStr.SstrNombreCampoBd, "ASC"},
                {ClsIdRecCajaEnt.SstrNombreCampoBd, "DESC"}}
        Dim lstrPreAgr As String = cboPredioAgr.SelectedItem
        If lstrPreAgr = GCSTRSINPA Then
            lstrPreAgr = String.Empty
        End If
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdCliente_RecDbl.SstrNombreCampoBd &
                " = " & txtIdCliente.Text & " AND " & ClsIdPredioAgrupador_RecStr.SstrNombreCampoBd
        If String.IsNullOrEmpty(lstrPreAgr) Then
            lstrFiltro &= " = ''"
        Else
            lstrFiltro &= " LIKE '%" & lstrPreAgr & "%'"
        End If
        Dim ldtbRsC As DataTable = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrIndice,
                lstrFiltro)
        Return ldtbRsC
    End Function

    Private Function FdtbNCR() As DataTable
        Dim lstrTabla = ClsNotaCr.SstrNombreTabla
        Dim lstrCamposSelect = {ClsPrefijo_NotaCrStr.SstrNombreCampoBd,
                ClsIdNotaCrEnt.SstrNombreCampoBd, ClsFecha_NotaCrDtm.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsFecha_NotaCrDtm.SstrNombreCampoBd, "DESC"},
                {ClsPrefijo_NotaCrStr.SstrNombreCampoBd, "ASC"},
                {ClsIdNotaCrEnt.SstrNombreCampoBd, "DESC"}}
        Dim lstrPreAgr = cboPredioAgr.SelectedItem
        If lstrPreAgr = GCSTRSINPA Then
            lstrPreAgr = String.Empty
        End If
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_NotaCrDbl.SstrNombreCampoBd &
                " = " & txtIdCliente.Text & " AND " &
                ClsIdPredioAgrupador_NotaCrStr.SstrNombreCampoBd &
                " = '" & lstrPreAgr & "'"
        Dim ldtbNCR As DataTable = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrIndice,
                lstrFiltro)
        Return ldtbNCR
    End Function

    Private Function FdtbND() As DataTable
        Dim lstrTabla = ClsNotaDb.SstrNombreTabla
        Dim lstrCamposSelect = {ClsPrefijo_NotaDbStr.SstrNombreCampoBd,
                ClsIdNotaDbEnt.SstrNombreCampoBd, ClsFecha_NotaDbDtm.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsFecha_NotaDbDtm.SstrNombreCampoBd, "DESC"},
                {ClsPrefijo_NotaDbStr.SstrNombreCampoBd, "ASC"},
                {ClsIdNotaDbEnt.SstrNombreCampoBd, "DESC"}}
        Dim lstrPreAgr = cboPredioAgr.SelectedItem
        If lstrPreAgr = GCSTRSINPA Then
            lstrPreAgr = String.Empty
        End If
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " &
                ClsIdCliente_NotaDbDbl.SstrNombreCampoBd &
                " = " & txtIdCliente.Text & " AND " &
                ClsIdPredioAgrupador_NotaDbStr.SstrNombreCampoBd &
                " = '" & lstrPreAgr & "'"
        Dim ldtbNDB As DataTable = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrIndice,
                lstrFiltro)
        Return ldtbNDB
    End Function

    Private Function FdtbNAA() As DataTable
        Dim lstrTabla = ClsNotaCon.SstrNombreTabla
        Dim lstrCamposSelect = {ClsPrefijo_NotaConStr.SstrNombreCampoBd,
                ClsIdNotaConEnt.SstrNombreCampoBd, ClsFecha_NotaConDtm.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsFecha_NotaConDtm.SstrNombreCampoBd, "DESC"},
                {ClsPrefijo_NotaConStr.SstrNombreCampoBd, "ASC"},
                {ClsIdNotaConEnt.SstrNombreCampoBd, "DESC"}}
        Dim lstrPreAgr = cboPredioAgr.SelectedItem
        If lstrPreAgr = GCSTRSINPA Then
            lstrPreAgr = String.Empty
        End If
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion & " AND " & ClsIdCliente_NotaConDbl.SstrNombreCampoBd &
                " = " & txtIdCliente.Text & " AND " & ClsIdPredioAgrupador_NotaConStr.SstrNombreCampoBd &
                " = '" & lstrPreAgr & "'"
        Dim ldtbNAA As DataTable = ClsPanorama.FdtbDataTable(lstrTabla, lstrCamposSelect, lstrIndice,
                lstrFiltro)
        Return ldtbNAA
    End Function

    Private Function FdtbRecibos() As DataTable
        Dim lstrFecIni As String = "'" &
                ClsPanoramaDat.FstrFechaHoraNormalizada(MobjObjetoWin.DtmFechaIni) & "'"
        Dim lstrFecFin As String = "'" &
                ClsPanoramaDat.FstrFechaHoraNormalizada(MobjObjetoWin.DtmFechaFin) & "'"
        Dim lstrTablaPri = ClsReciboCaja.SstrNombreTabla
        Dim lstrCamposPri As String() = {ClsPrefijo_RecStr.SstrNombreCampoBd,
                ClsIdRecCajaEnt.SstrNombreCampoBd}
        Dim lstrTablaSec = ClsCliente.SstrNombreTabla
        Dim lstrCamposSec As String() = Array.Empty(Of String)()
        Dim LstrCamposPriRel As String() = {StrCampoCarpeta,
                StrCampoCentroutil,
                ClsIdCliente_RecDbl.SstrNombreCampoBd}
        Dim lstrCamposSecRel As String() = {StrCampoCarpeta,
                StrCampoCentroutil, ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrIndice = {{ClsIdRecCajaEnt.SstrNombreCampoBd, "ASC"}}
        Dim lstrFiltro = "P." & StrCampoCarpeta & " = " & GshrIdCarpeta &
                " AND P." & StrCampoCentroutil & " = " & GshrIdCentroUtil &
                " AND " & ClsRecibeDocsPorEmailBln.SstrNombreCampoBd & " = TRUE " &
                " AND " & ClsFechaRecDtm.SstrNombreCampoBd &
                " BETWEEN " & lstrFecIni & " AND " & lstrFecFin
        Dim ldtbRecs = ClsPanorama.FdtbDataTable(lstrTablaPri, lstrCamposPri, lstrTablaSec, lstrCamposSec,
                LstrCamposPriRel, lstrCamposSecRel, lstrIndice, lstrFiltro,
                Array.Empty(Of String)(), True)
        Return ldtbRecs
    End Function

    Private Function FarlPrediosAgrup() As ArrayList
        Dim larlPreAgr As ArrayList = Nothing
        If Not IsNothing(MobjCliente) AndAlso MobjCliente.BlnExiste Then
            Select Case MobjObjetoWin.EnuTipoCorreo
                Case EnuTipoCorreoE.EnuFac
                    larlPreAgr = MobjCliente.FarlPrediosAgrupEnFras
                Case EnuTipoCorreoE.EnuRC
                    larlPreAgr = MobjCliente.FarlPrediosAgrupEnRC
                Case EnuTipoCorreoE.EnuNCR
                    larlPreAgr = MobjCliente.FarlPrediosAgrupEnNCR
                Case EnuTipoCorreoE.EnuNDB
                    larlPreAgr = MobjCliente.FarlPrediosAgrupEnNDB
                Case EnuTipoCorreoE.EnuNAA
                    larlPreAgr = MobjCliente.FarlPrediosAgrupEnNAA
                Case Else
                    larlPreAgr = Nothing
            End Select
        End If
        Return larlPreAgr
    End Function

    Private Sub STextosDefecto()
        Dim lstrAsunto = String.Empty, lstrIdpredAgru As String
        Dim lstrMens = String.Empty, lstrDoc As String
        lstrIdpredAgru = FstrPredAgru()
        If GobjParametros.BlnEFacAutorizado Then
            lstrDoc = "Factura"
        Else
            lstrDoc = "Cuenta de Cobro"
        End If
        Select Case MobjObjetoWin.EnuTipoCorreo
            Case EnuTipoCorreoE.EnuFactAuto
                lstrAsunto = lstrDoc & " del Mes"
                lstrMens = "Adjunto encuentra la " & lstrDoc & " correspondiente al presente Mes." &
                        vbCrLf & "Cordial saludo."
            Case EnuTipoCorreoE.EnuFac
                lstrAsunto = lstrDoc
                lstrMens = "Adjunto encuentra la " & lstrDoc & " número " & MobjObjetoWin.StrNroDocumento &
                        " correspondiente al predio " & lstrIdpredAgru & vbCrLf &
                        "Cordial saludo."
            Case EnuTipoCorreoE.EnuRC
                lstrAsunto = "Recibo de Caja"
                lstrMens = "Adjunto encuentra el Recibo de Caja número " & MobjObjetoWin.StrNroDocumento &
                        " correspondiente al Predio " & lstrIdpredAgru & vbCrLf &
                        "Cordial saludo."
            Case EnuTipoCorreoE.EnuNAA
                lstrAsunto = "Nota Legalización Anticipo"
                lstrMens = "Adjunto encuentra la Nota Legalización Anticipo número " &
                        MobjObjetoWin.StrNroDocumento & " correspondiente al Predio " &
                        lstrIdpredAgru & vbCrLf & "Cordial saludo."
            Case EnuTipoCorreoE.EnuNCR
                lstrAsunto = "Nota Crédito"
                lstrMens = "Adjunto encuentra la Nota Crédito número " & MobjObjetoWin.StrNroDocumento &
                        " correspondiente al Predio " & lstrIdpredAgru & vbCrLf &
                        "Cordial saludo."
            Case EnuTipoCorreoE.EnuNDB
                lstrAsunto = "Nota Débito"
                lstrMens = "Adjunto encuentra la Nota Débito por Intereses de Mora número " &
                        MobjObjetoWin.StrNroDocumento & " correspondiente al Predio " &
                        lstrIdpredAgru & vbCrLf &
                        "Cordial saludo."
            Case EnuTipoCorreoE.EnuRecibos
                lstrAsunto = "Recibo de Caja"
                dtpFecIni.SelectedDate = Date.Today.AddDays(-7)
                dtpFecFin.SelectedDate = Date.Today
                lstrMens = "Adjunto encuentra el Recibo de Caja correspondiente al pago efectuado entre " &
                         vbCrLf & " el " & dtpFecIni.SelectedDate & " y el " & dtpFecFin.SelectedDate &
                         vbCrLf & "Cordial saludo."
            Case EnuTipoCorreoE.EnuCobroPers
                lstrMens = FstrLeaMensCobroPresuIni()
                lstrAsunto = "Cobro persuasivo"
                If String.IsNullOrEmpty(lstrMens) Then
                    lstrMens = "Apreciado(a) copropietario:" & vbCrLf & "Adjunto " &
                        "le estoy enviando el estado de cuenta a la fecha, " &
                        "en el cual se informa el saldo por pagar en mora. De una " &
                        "manera muy cordial me permito solicitarle hacer el pago " &
                        "respectivo lo antes posible. " & vbCrLf & "Si tiene algún " &
                        "reparo a la información enviada, o si bien ya hizo el pago, " &
                        "cordialmente le solicito ponerse en contacto con esta " &
                        "administración, con el fin de aclarar dicho estado de cuenta." &
                        vbCrLf & "Cordial saludo y muchas garcias por su atención."
                End If
                txtDiasVen.Text = GobjParametros.ObjDiasParaPersuasivoShr.ObjValorPro
            Case EnuTipoCorreoE.None
                lstrAsunto = String.Empty
                lstrMens = String.Empty
        End Select
        txtAsunto.Text = lstrAsunto
        txtMensaje.Text = lstrMens
        StcValidValido(EnuValidEntrada.EnuAsun) =
                MobjObjetoWin.FblnEsValidoAsunto(txtAsunto.Text, lstrMens)
        If MobjObjetoWin.EnuTipoCorreo = EnuTipoCorreoE.EnuCobroPers Then
            StcValidValido(EnuValidEntrada.EnuMens) =
                MobjObjetoWin.FblnEsValidoMensajeCobroPer(txtMensaje.Text, lstrMens)
        Else
            StcValidValido(EnuValidEntrada.EnuMens) =
                MobjObjetoWin.FblnEsValidoMensaje(txtMensaje.Text, lstrMens)
        End If
    End Sub

    Private Sub SHabiliteControles(ablnHabilte As Boolean)
        PanelControl.IsEnabled = ablnHabilte
    End Sub

    Private Sub SEspereIntervalo(aentMinutosEspera As Double, aentProgreso As Integer)
        If aentMinutosEspera > 0 Then
            Dim llngEspera As Long = aentMinutosEspera * 60 * 1000
            Dim lentFrecuencia As Integer = 3000
            Dim llngCiclos As Long = Int(llngEspera / lentFrecuencia), i As Long = 0
            Dim llngTiempoEsperado As Long, lentCanMinEspera As Integer, lentCantSegEspera As Integer
            Do While i <= llngCiclos
                llngTiempoEsperado = i * lentFrecuencia / 1000
                lentCanMinEspera = Int((llngTiempoEsperado) / 60)
                lentCantSegEspera = llngTiempoEsperado - (lentCanMinEspera * 60)
                lentCanMinEspera = aentMinutosEspera - lentCanMinEspera
                If lentCantSegEspera > 0 Then
                    lentCantSegEspera = 60 - lentCantSegEspera
                    lentCanMinEspera -= 1
                End If
                MstrNotifica = "Faltan " & lentCanMinEspera.ToString &
                        " Minutos y " & lentCantSegEspera.ToString &
                        " Segundos para reanudar el Envio de los Mensajes!"
                MbgwCorreo.ReportProgress(aentProgreso)
                Thread.Sleep(lentFrecuencia)
                i += 1
            Loop
            MstrNotifica = String.Empty
            MbgwCorreo.ReportProgress(aentProgreso)
        End If
    End Sub

    Private Function FstrPredAgru() As String
        Dim lstrpreAgr As String
        If String.IsNullOrEmpty(MobjObjetoWin.StrIdPredioAgrupador) Then
            lstrpreAgr = GCSTRSINPA
        Else
            lstrpreAgr = MobjObjetoWin.StrIdPredioAgrupador
        End If
        Return lstrpreAgr
    End Function

    Private Sub SMuestreCorreos()
        If tbiHistorico.IsSelected Then
            Dim ldtbCorreos As DataTable
            Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.EnuUnico)
            If Not String.IsNullOrEmpty(txtIdClienteHis.Text) Then
                If txtIdClienteHis.Text = My.Resources.Todos Then
                    bttMostrarTodos.Visibility = Visibility.Hidden
                    ldtbCorreos = MobjObjetoWin.FdtbCorreosEnviados(0, dtpDesde.SelectedDate,
                            dtpHasta.SelectedDate)
                    txtNombreClienteHist.Content = "Todos los clientes"
                Else
                    bttMostrarTodos.Visibility = Visibility.Visible
                    lobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, txtIdClienteHis.Text})
                    If lobjCliente.BlnExiste Then
                        txtNombreClienteHist.Content = lobjCliente.ObjNombreCompletoStr.ObjValorPro
                        ldtbCorreos = MobjObjetoWin.FdtbCorreosEnviados(
                                lobjCliente.ObjIdClienteDbl.ObjValorPro, dtpDesde.SelectedDate,
                            dtpHasta.SelectedDate)
                    Else
                        ldtbCorreos = Nothing
                        Dim lstrMens = "La Id. del Cliente ingresada, no existe"
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                End If
                dgrHistorico.DataContext = ldtbCorreos
            End If
        End If
    End Sub

    Private Sub SMuestreHistoricoPorAsunto()
        If tbiHistoricoAsunto.IsSelected Then
            Dim lbytIdTipo = cboTipoCorreoH.SelectedIndex
            Dim ldtbHistAsunto = MobjObjetoWin.FdtbHistoricoTipo(lbytIdTipo, dtpFechaDesde.SelectedDate,
                    dtpFechaHasta.SelectedDate)
            dgrHistoricoAsunto.DataContext = ldtbHistAsunto
        End If
    End Sub

    Private Sub SEscribaMensCobroPresuIni()
        Dim lstrMensCobPer = txtMensaje.Text
        Dim lstrArchivo = GstrTrayDatPrg & "MensCobPer.ini"
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            If My.Computer.FileSystem.FileExists(lstrArchivo) Then
                My.Computer.FileSystem.DeleteFile(lstrArchivo)
            End If
            Using lswArchivoMens = File.AppendText(lstrArchivo)
                lswArchivoMens.WriteLine(lstrMensCobPer)
                lswArchivoMens.Flush()
            End Using
            lblnNoHayError = True
        Catch ex As ArgumentException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PathTooLongException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As DirectoryNotFoundException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As NotSupportedException
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
    End Sub

    Private Function FstrLeaMensCobroPresuIni() As String
        Dim lsrArchivoIni As StreamReader
        Dim lstrLinea As String, lstrMensaje = String.Empty
        Dim lstrArchivoIni As String = GstrTrayDatPrg & "MensCobPer.ini"
        If My.Computer.FileSystem.FileExists(lstrArchivoIni) Then
            lsrArchivoIni = ClsPanorama.FsrStreamReader(lstrArchivoIni)
            If Not IsNothing(lsrArchivoIni) Then
                lstrLinea = lsrArchivoIni.ReadLine
                Do While Not IsNothing(lstrLinea)
                    lstrMensaje &= lstrLinea
                    lstrLinea = lsrArchivoIni.ReadLine
                Loop
                lsrArchivoIni.Close()
            End If
        End If
        lstrMensaje = lstrMensaje.Replace(".", "." & vbCrLf)
        Return lstrMensaje
    End Function

    Private Sub SVerifiquePorEnviar()
        If MobjObjetoWin.FblnHayCorreoPorEnviar Then
            Dim lstrMens = String.Empty
            If MobjObjetoWin.FblnHayConInternet(lstrMens) Then
                lstrMens = "Hay correo pendiente por ser enviado. Desea enviarlo?"
                Dim lblnEnviar = MsgBox(lstrMens, vbYesNo, "Correo por enviar") = MsgBoxResult.Yes
                If lblnEnviar Then
                    MblnPoblandoCombo = True
                    txtAsunto.Text = MobjObjetoWin.StrAsunto
                    txtMensaje.Text = MobjObjetoWin.StrMensaje
                    If MobjObjetoWin.EnuTipoCorreo = EnuTipoCorreoE.EnuArchExt Then
                        txtArchivo.Text = MobjObjetoWin.StrArchivoExterno
                    Else
                        txtArchivo.Text = String.Empty
                    End If
                    If MobjObjetoWin.EnuTipoCorreo = EnuTipoCorreoE.EnuCobroPers Then
                        txtDiasVen.Text = MobjObjetoWin.EntDiasCobroPers
                    Else
                        txtDiasVen.Text = String.Empty
                    End If
                    cboTipoCorreo.SelectedIndex = MobjObjetoWin.EnuTipoCorreo
                    SMuestreDatos()
                    If MobjObjetoWin.EnuTipoCorreo = EnuTipoCorreoE.EnuFactAuto Then
                        chkSeguro.IsChecked = True
                    End If
                    MblnPoblandoCombo = False
                    SGuarde()
                Else
                    MobjObjetoWin.SRegistreUltimo()
                    SFinaliceOperacion()
                    SCree()
                    SMuestreDatos()
                End If
            Else
                lstrMens = "Hay correo pendiente por ser enviado, pero no hay conexión a internet!"
                MsgBox(lstrMens, vbOKOnly, "Correo por enviar")
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            End If
        End If
    End Sub
#End Region

#Region "Exportacion"
    Private Sub SExporteDoc(aobjDoc As ClsCBObjetoPan)
        Select Case MobjObjetoWin.EnuTipoCorreo
            Case EnuTipoCorreoE.EnuFac
                Dim lobjFactura As ClsFactura = aobjDoc
                MobjReportes.SExporteFactura(lobjFactura, MobjTerceroCentrUtil, True)
            Case EnuTipoCorreoE.EnuRC, EnuTipoCorreoE.EnuRecibos
                Dim lobjRecCaja As ClsReciboCaja = aobjDoc
                MobjReportes.SExporteUnReciboCaja(lobjRecCaja, MobjTerceroCentrUtil,
                        GobjParametros.ObjFirmaRCeMail.ObjValorPro)
            Case EnuTipoCorreoE.EnuNCR
                Dim lobjNotaCr As ClsNotaCr = aobjDoc
                MobjReportes.SExporteNotaCr(lobjNotaCr, MobjTerceroCentrUtil)
            Case EnuTipoCorreoE.EnuNDB
                Dim lobjNotaDb As ClsNotaDb = aobjDoc
                MobjReportes.SExporteNotaDb(lobjNotaDb, MobjTerceroCentrUtil)
            Case EnuTipoCorreoE.EnuNAA
                Dim lobjNotaAA As ClsNotaCon = aobjDoc
                MobjReportes.SExporteNotaCon(lobjNotaAA, MobjTerceroCentrUtil)
        End Select
    End Sub
#End Region

#Region "Envio correos"
    Private Function FblnEnvioFacsAutoMes(ByRef ablnTimeOut As Boolean) As Boolean
        Dim ldtbFrasEnviar = ClsOrionCop.FdtbFacsMesEnviarEmail
        Dim lobjFactura As New ClsFactura()
        Dim lobjValorLlave As Object() = Array.Empty(Of Object)()
        Dim lstrEmailDestino As ArrayList, lblnEnvioFact = True
        Dim lentCantFras = ldtbFrasEnviar.Rows.Count
        Dim lstrPref = String.Empty, lentIdFact = 0, ldblIdCliente As Double
        Dim lstrArchFac = String.Empty, j = 0, lentMensPorTanda = 0
        Dim lentIntervaloEntreTandas = 0, lentProgreso = 0, i = 0
        With GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
            lentMensPorTanda = .ObjMensajesPorTandaShr.ObjValorPro
            lentIntervaloEntreTandas = .ObjIntervaloEntreTandasShr.ObjValorPro
        End With
        MblnEnvioOk = False
        For Each ldrwFac As DataRow In ldtbFrasEnviar.Rows
            i += 1
            If i Mod 50 = 0 Then
                MobjReportes = Nothing
                MobjReportes = New ClsRepOrionCop(GCOBJREGISTRO)
            End If
            lstrPref = ClsPanorama.FobjValorCampo(ldrwFac(ClsPrefijo_FactStr.SstrNombreCampoBd),
                    EnuTipoValor.EnuString)
            lentIdFact = ClsPanorama.FobjValorCampo(ldrwFac(ClsIdFacturaEnt.SstrNombreCampoBd),
                    EnuTipoValor.EnuInteger)
            lobjValorLlave = {MshrIdCarpeta, MshrIdCentroUtil, lstrPref, lentIdFact}
            lobjFactura.SAbra(lobjValorLlave)
            MblnEnvioOk = True
            If Not lobjFactura.ObjEnviadaMailBln.ObjValorPro AndAlso
                    lobjFactura.BlnEnviarPorCorreo Then
                ldblIdCliente = lobjFactura.ObjIdCliente_FactDbl.ObjValorPro
                lstrEmailDestino = lobjFactura.FarlCorreosFac
                If lstrEmailDestino.Count > 0 Then
                    MobjReportes.SExporteFactura(lobjFactura, MobjTerceroCentrUtil, True)
                    lstrArchFac = ClsOrionCop.FstrArchivoPdfDcto(lstrPref, lentIdFact,
                            EnuTipoDocOri.EnuFactura)
                    Dim lstrNroFac = ClsPanorama.FstrNumeroDcto(lstrPref, lentIdFact)
                    If My.Computer.FileSystem.FileExists(lstrArchFac) Then
                        lblnEnvioFact = FblnEnvioEmail(lstrEmailDestino, lstrArchFac, lstrNroFac,
                                ldblIdCliente, ablnTimeOut)
                        If lblnEnvioFact AndAlso Not ablnTimeOut Then
                            SRegistreFraPublicada(lobjFactura)
                            j += 1
                            lentProgreso = Int((i / lentCantFras) * 100)
                            MbgwCorreo.ReportProgress(lentProgreso)
                            If j = lentMensPorTanda Then
                                SEspereIntervalo(lentIntervaloEntreTandas, lentProgreso)
                                j = 0
                            End If
                        Else
                            Exit For
                        End If
                    Else
                        SEscribaReporteEmails(False, lobjFactura.StrNumeroFactura,
                                "No se encontró archivo.")
                    End If
                End If
            Else
                Dim lstrMens As String = "Cuenta de cobro"
                If GobjParametros.BlnEFacAutorizado Then
                    lstrMens = "Factura"
                End If
                SEscribaReporteEmails(False, lobjFactura.StrNumeroFactura, "Factura ya publicada.")
            End If
            If MbgwCorreo.CancellationPending Then
                MblnCancelando = True
                Exit For
            End If
        Next
        If Not MblnCancelando AndAlso lblnEnvioFact Then
            MbgwCorreo.ReportProgress(100)
        End If
        Return lblnEnvioFact
    End Function

    Private Shared Sub SRegistreFraPublicada(aobjFactura As ClsFactura)
        If aobjFactura.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            aobjFactura.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
        aobjFactura.ObjEnviadaMailBln.ObjValorPro = True
        aobjFactura.SActualice(True)
    End Sub

    Private Sub SRegistreEnvioFac(astrPrefFac As String, aentIdFac As Integer)
        Dim lobjFac As New ClsFactura()
        Dim lobjValorLlave As Object = {MshrIdCarpeta, MshrIdCentroUtil, astrPrefFac, aentIdFac}
        lobjFac.SAbra(lobjValorLlave)
        SRegistreFraPublicada(lobjFac)
    End Sub

    Private Shared Sub SRegistreRecCajaPublicado(aobjRec As ClsReciboCaja)
        If aobjRec.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            aobjRec.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando
        End If
        aobjRec.ObjRCEnviadoMailBln.ObjValorPro = True
        aobjRec.SActualice(True)
    End Sub

    Private Function FblnEnvioRecibos(ByRef ablnTimeOut As Boolean) As Boolean
        Const LCENTPRO = 1
        Dim lstrPref = String.Empty, lentIdRec = 0, lentProgreso = 0, lentCantMen = 0
        Dim lblnEnvio = False, lentMensPorTanda = 0
        Dim lentIntervaloEntreTandas = 0, i = 0, j = 0
        With GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
            lentMensPorTanda = .ObjMensajesPorTandaShr.ObjValorPro
            lentIntervaloEntreTandas = .ObjIntervaloEntreTandasShr.ObjValorPro
        End With
        Dim lobjRec As New ClsReciboCaja()
        Dim lobjValorLlave As Object() = Array.Empty(Of Object)()
        Dim ldtbRecs As DataTable = FdtbRecibos()
        Dim lstrEmail As ArrayList, lstrNomArchivo = String.Empty, ldblIdCliente As Double
        lentCantMen = ldtbRecs.Rows.Count
        MblnEnvioOk = False
        If lentCantMen > 0 Then
            For Each ldrwRec As DataRow In ldtbRecs.Rows
                lstrPref = ClsPanorama.FobjValorCampo(ldrwRec(ClsPrefijo_RecStr.SstrNombreCampoBd),
                        EnuTipoValor.EnuString)
                lentIdRec = ClsPanorama.FobjValorCampo(ldrwRec(ClsIdRecCajaEnt.SstrNombreCampoBd),
                        EnuTipoValor.enuInteger)
                lobjValorLlave = {MshrIdCarpeta, MshrIdCentroUtil, lstrPref, lentIdRec}
                lobjRec.SAbra(lobjValorLlave)
                lstrEmail = lobjRec.FarlCorreosRec
                MblnEnvioOk = True
                lblnEnvio = True
                If lstrEmail.Count > 0 AndAlso Not lobjRec.ObjRCEnviadoMailBln.ObjValorPro Then
                    For Each lstrMail As String In lstrEmail
                        ldblIdCliente = lobjRec.ObjIdCliente_RecDbl.ObjValorPro
                    Next
                    lstrNomArchivo = ClsOrionCop.FstrNombreArchRecibo(lstrPref, lentIdRec)
                    SEscribaReporteEmails(True, String.Empty, String.Empty)
                    SExporteDoc(lobjRec)
                    i += LCENTPRO
                    If My.Computer.FileSystem.FileExists(lstrNomArchivo) Then
                        lblnEnvio = FblnEnvioEmail(lstrEmail, lstrNomArchivo, lobjRec.StrIdObjeto,
                                ldblIdCliente, ablnTimeOut)
                        If lblnEnvio AndAlso Not ablnTimeOut Then
                            SRegistreRecCajaPublicado(lobjRec)
                            j += 1
                            lentProgreso = Int((i / lentCantMen) * 100)
                            MbgwCorreo.ReportProgress(lentProgreso)
                            If j = lentMensPorTanda Then
                                SEspereIntervalo(lentIntervaloEntreTandas, lentProgreso)
                                j = 0
                            End If
                        Else
                            Exit For
                        End If
                    Else
                        SEscribaReporteEmails(False, MobjObjetoWin.StrNroDocumento, String.Empty)
                    End If
                    lentProgreso = Int((i / lentCantMen) * 100)
                    MbgwCorreo.ReportProgress(lentProgreso)
                End If
                If MbgwCorreo.CancellationPending Then
                    MblnCancelando = True
                    Exit For
                End If
            Next
        Else
            lblnEnvio = True
            MblnEnvioOk = True
        End If
        Return lblnEnvio
    End Function

    Private Function FblnEnvioArchivo(ByRef ablnTimeOut As Boolean) As Boolean
        Dim lblnOk = String.IsNullOrEmpty(MobjObjetoWin.StrArchivoExterno) AndAlso
            MobjObjetoWin.EnuTipoCorreo = EnuTipoCorreoE.EnuSoloMens
        If Not lblnOk Then
            lblnOk = MobjObjetoWin.StrArchivoExterno.Length > 0 AndAlso
                MobjObjetoWin.EnuTipoCorreo <> EnuTipoCorreoE.EnuSoloMens
        End If
        Dim lblnEnvio As Boolean = False
        MblnEnvioOk = False
        If lblnOk Then
            Dim lentMensPorTanda = 0, lentIntervaloEntreTandas = 0, lentProgreso = 0
            Dim lentCantAProcesar = 0, i = 0, k = 0, ldblIDCliente = 0.0
            With GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
                lentMensPorTanda = .ObjMensajesPorTandaShr.ObjValorPro
                lentIntervaloEntreTandas = .ObjIntervaloEntreTandasShr.ObjValorPro
            End With
            Dim lobjValorLlave As Object() = Array.Empty(Of Object)()
            Dim ldtbClientesConCorreo = MobjObjetoWin.FdtbClientesAEnviar(False)
            lentCantAProcesar = ldtbClientesConCorreo.Rows.Count
            Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
            Dim lstrEmail As New ArrayList
            If lentCantAProcesar > 0 Then
                SEscribaReporteEmails(True, String.Empty, String.Empty)
            End If
            For Each ldrwIdCliente As DataRow In ldtbClientesConCorreo.Rows
                MblnHayConeccion = FblnHayInternet()
                If Not MblnHayConeccion Then
                    lblnEnvio = False
                    Exit For
                End If
                MblnEnvioOk = True
                k += 1
                ldblIDCliente = ClsPanorama.FobjValorCampo(ldrwIdCliente(0),
                        EnuTipoValor.EnuDouble)
                lobjValorLlave = {MshrIdCarpeta, MshrIdCentroUtil, ldblIDCliente}
                lobjCliente.SAbra(lobjValorLlave)
                If lobjCliente.ObjRecibeDocsPorEmailBln.ObjValorPro AndAlso
                        lobjCliente.ObjEmailStr.ObjValorPro.Length > 0 Then
                    lstrEmail.Add(lobjCliente.ObjEmailStr.ObjValorPro)
                    lblnEnvio = FblnEnvioEmail(lstrEmail, MobjObjetoWin.StrArchivoExterno,
                        String.Empty, ldblIDCliente, ablnTimeOut)
                    If lblnEnvio Then
                        i += 1
                        lentProgreso = Int((k / lentCantAProcesar) * 100)
                        MbgwCorreo.ReportProgress(lentProgreso)
                        If i >= lentMensPorTanda Then
                            SEspereIntervalo(lentIntervaloEntreTandas, lentProgreso)
                            i = 0
                        End If
                        lstrEmail.Clear()
                    End If
                End If
                If MbgwCorreo.CancellationPending Then
                    MblnCancelando = True
                    Exit For
                End If
            Next
            If Not MblnCancelando AndAlso lblnEnvio Then
                If lstrEmail.Count > 0 Then
                    lblnEnvio = FblnEnvioEmail(lstrEmail, MobjObjetoWin.StrArchivoExterno,
                        String.Empty, ldblIDCliente, ablnTimeOut)
                    lstrEmail.Clear()
                End If
                If lblnEnvio Then
                    MbgwCorreo.ReportProgress(100)
                End If
            End If
        Else
            Throw New ErrorInesperadoPanLException("No esta definido el archivo a ser enviado!")
        End If
        Return lblnEnvio
    End Function

    Private Function FblnEnvioDoc(ByRef ablnTimeOut As Boolean) As Boolean
        Const lentProgreso As Integer = 100
        Dim lstrPref = ClsPanorama.FstrPrefijoDcto(MobjObjetoWin.StrNroDocumento)
        Dim lentIdDoc = ClsPanorama.FentIdDcto(MobjObjetoWin.StrNroDocumento)
        Dim lobjValorLlave As Object() = {MshrIdCarpeta, MshrIdCentroUtil, lstrPref, lentIdDoc}
        Dim lenuTipoDoc As EnuTipoDocOri = MobjObjetoWin.FenuTipoDocOrigen
        Dim lstrNomArchivo As String, lobjDocum As ClsCBObjetoPan = Nothing,
                lstrNroDoc = String.Empty
        Dim lstrEmail As ArrayList = Nothing, lblnEnvio = False
        If lenuTipoDoc = EnuTipoDocOri.EnuReciboCaja Then
            lstrNomArchivo = ClsOrionCop.FstrNombreArchRecibo(lstrPref, lentIdDoc)
        Else
            lstrNomArchivo = ClsOrionCop.FstrArchivoPdfDcto(lstrPref, lentIdDoc, lenuTipoDoc)
        End If
        Dim ldblIdCliente As Double
        Select Case lenuTipoDoc
            Case EnuTipoDocOri.EnuFactura
                Dim lobjDoc As ClsFactura = FobjFactura(lobjValorLlave, lstrEmail)
                lobjDocum = lobjDoc
                ldblIdCliente = lobjDoc.ObjIdCliente_FactDbl.ObjValorPro
                lstrNroDoc = lobjDoc.StrIdObjeto
            Case EnuTipoDocOri.EnuReciboCaja
                Dim lobjDoc As ClsReciboCaja = FobjRecCaja(lobjValorLlave, lstrEmail)
                lobjDocum = lobjDoc
                ldblIdCliente = lobjDoc.ObjIdCliente_RecDbl.ObjValorPro
                lstrNroDoc = lobjDoc.StrIdObjeto
            Case EnuTipoDocOri.EnuNotaCr
                Dim lobjDoc As ClsNotaCr = FobjNotaCr(lobjValorLlave, lstrEmail)
                lobjDocum = lobjDoc
                ldblIdCliente = lobjDoc.ObjIdCliente_NotaCrDbl.ObjValorPro
                lstrNroDoc = lobjDoc.StrIdObjeto
            Case EnuTipoDocOri.EnuNotaDb
                Dim lobjDoc As ClsNotaDb = FobjNotaDb(lobjValorLlave, lstrEmail)
                lobjDocum = lobjDoc
                ldblIdCliente = lobjDoc.ObjIdCliente_NotaDbDbl.ObjValorPro
                lstrNroDoc = lobjDoc.StrIdObjeto
            Case EnuTipoDocOri.EnuNotaCon
                Dim lobjDoc As ClsNotaCon = FobjNotaCon(lobjValorLlave, lstrEmail)
                lobjDocum = lobjDoc
                ldblIdCliente = lobjDoc.ObjIdCliente_NotaConDbl.ObjValorPro
                lstrNroDoc = lobjDoc.StrIdObjeto
        End Select
        MblnEnvioOk = False
        SEscribaReporteEmails(True, String.Empty, String.Empty)
        If lstrEmail.Count > 0 Then
            MblnEnvioOk = True
            SExporteDoc(lobjDocum)
            lblnEnvio = FblnEnvioEmail(lstrEmail, lstrNomArchivo, lstrNroDoc, ldblIdCliente,
                    ablnTimeOut)
            If lblnEnvio Then
                If lenuTipoDoc = EnuTipoDocOri.EnuFactura Then
                    SRegistreEnvioFac(lstrPref, lentIdDoc)
                End If
                MbgwCorreo.ReportProgress(lentProgreso)
            End If
        End If
        Return lblnEnvio
    End Function

    Private Function FblnEnvioCuentasCobro(ByRef ablnTimeOut As Boolean) As Boolean
        Dim lentProgreso As Integer, i = 0
        Dim j = 0, lentMensPorTanda = 0, lentIntervaloEntreTandas = 0
        Dim lobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
        Dim lobjValorLlave As Object(), lstrNombreArch As String
        Dim ldtbDeudoresEnMora = ClsOrionCop.FdtbPrediosAgrMorosos(MentDiasCobro, 0)
        Dim lentCantAProcesar = ldtbDeudoresEnMora.Rows.Count
        Dim lstrEmail As New ArrayList()
        Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
            .EnuReporte = EnuReporteDef.enuEstadoCtaCli
            }
        With GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
            lentMensPorTanda = .ObjMensajesPorTandaShr.ObjValorPro
            lentIntervaloEntreTandas = .ObjIntervaloEntreTandasShr.ObjValorPro
        End With
        MblnEnvioOk = False
        Dim lblnEnvio = False
        If ldtbDeudoresEnMora.Rows.Count > 0 Then
            For Each ldrwDeuMor As DataRow In ldtbDeudoresEnMora.Rows
                i += 1 : lstrEmail.Clear
                lstrNombreArch = String.Empty
                MblnEnvioOk = True
                Dim lstrIdPredAgr As String = ClsPanorama.FobjValorCampo(ldrwDeuMor(
                    ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd), EnuTipoValor.EnuString)
                Dim ldblIdCliente As Double = ClsPanorama.FobjValorCampo(ldrwDeuMor(
                    ClsIdCliente_FactDbl.SstrNombreCampoBd), EnuTipoValor.EnuDouble)
                lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, ldblIdCliente}
                lobjCliente.SAbra(lobjValorLlave)
                If lobjCliente.ObjRecibeDocsPorEmailBln.ObjValorPro Then
                    MblnEnvioOk = True
                    lstrEmail.Add(lobjCliente.ObjEmailStr.ObjValorPro)
                    SGenereEstadoCuenta(lobjRep, lobjCliente, lstrIdPredAgr, lstrNombreArch)
                    lblnEnvio = FblnEnvioEmail(lstrEmail, lstrNombreArch, String.Empty,
                        ldblIdCliente, ablnTimeOut)
                    If lblnEnvio Then
                        If i Mod 50 = 0 Then
                            MobjReportes = Nothing
                            MobjReportes = New ClsRepOrionCop(GCOBJREGISTRO)
                        End If
                        lentProgreso = Int((i / lentCantAProcesar) * 100)
                        MbgwCorreo.ReportProgress(lentProgreso)
                        j += 1
                        If j = lentMensPorTanda Then
                            SEspereIntervalo(lentIntervaloEntreTandas, lentProgreso)
                            j = 0
                        End If
                    End If
                End If
            Next
        Else
            lblnEnvio = True
            MblnEnvioOk = True
        End If
        Return lblnEnvio
    End Function

    Private Sub SGenereEstadoCuenta(aobjRep As ClsRepOrionCop, aobjCliente As ClsCliente,
            astrIdPredioAgr As String, ByRef astrNombreArchivo As String)
        Dim ldecIntPorCausar = aobjCliente.FdecIntMoraPorCausar({astrIdPredioAgr}, Date.Today)
        With aobjRep
            .DblIdCliente = aobjCliente.ObjIdClienteDbl.ObjValorPro
            If String.IsNullOrEmpty(astrIdPredioAgr) Then astrIdPredioAgr = GCSTRSINPA
            .StrIdPredioAgru = astrIdPredioAgr
            .DecIntPortCausar = ldecIntPorCausar
            .SExporteEstadoCta(astrNombreArchivo)
        End With
    End Sub

    Private Shared Function FobjFactura(aobjValorLlave As Object(),
        ByRef aarlCorreo As ArrayList) As ClsFactura
        Dim lobjFac As New ClsFactura()
        lobjFac.SAbra(aobjValorLlave)
        aarlCorreo = lobjFac.FarlCorreosFac
        Return lobjFac
    End Function

    Private Shared Function FobjRecCaja(aobjValorLlave As Object(),
            ByRef aarlCorreo As ArrayList) As ClsReciboCaja
        Dim lobjRecCaja As New ClsReciboCaja()
        lobjRecCaja.SAbra(aobjValorLlave)
        aarlCorreo = lobjRecCaja.FarlCorreosRec
        Return lobjRecCaja
    End Function

    Private Shared Function FobjNotaCr(aobjValorLlave As Object(),
        ByRef aarlCorreo As ArrayList) As ClsNotaCr
        Dim lobjNotaCr As New ClsNotaCr()
        lobjNotaCr.SAbra(aobjValorLlave)
        aarlCorreo = lobjNotaCr.FarlCorreosNCr
        Return lobjNotaCr
    End Function

    Private Shared Function FobjNotaDb(aobjValorLlave As Object(),
        ByRef aarlCorreo As ArrayList) As ClsNotaDb
        Dim lobjNotaDb As New ClsNotaDb()
        lobjNotaDb.SAbra(aobjValorLlave)
        aarlCorreo = lobjNotaDb.FarlCorreosNDb
        Return lobjNotaDb
    End Function

    Private Shared Function FobjNotaCon(aobjValorLlave As Object(),
        ByRef aarlCorreo As ArrayList) As ClsNotaCon
        Dim lobjNotaCon As New ClsNotaCon()
        lobjNotaCon.SAbra(aobjValorLlave)
        aarlCorreo = lobjNotaCon.FarlCorreosNCon
        Return lobjNotaCon
    End Function

    Private Function FblnEnvioEmail(astrEmailDestino As ArrayList, astrArchivoAEnviar As String,
            astrNroDoc As String, adblIdCliente As Double,
            ByRef ablnTimeOut As Boolean) As Boolean
        Dim lblnFueEnviadoMens = False
        If Not String.IsNullOrEmpty(astrArchivoAEnviar) Then
            If Not My.Computer.FileSystem.FileExists(astrArchivoAEnviar) Then
                Throw New ErrorInesperadoPanLException("Archivo " & astrArchivoAEnviar &
                        " no Encotrado!")
            End If
        End If
        Dim lstrMailOrigen = MobjCentroUtilActual.ObjEmailOrigenStr.ObjValorPro
        Dim lstrContrasena = MobjCentroUtilActual.FstrContrasena
        Dim lstrMensaje = MobjObjetoWin.StrMensaje
        lblnFueEnviadoMens = ClsPanorama.FblnEnvioMensaje(astrEmailDestino,
                MobjObjetoWin.StrAsunto, lstrMensaje, astrArchivoAEnviar,
                MstrServidor, MblnHabilitarSSL, MentPuerto, lstrMailOrigen,
                MblnRequiereAutent, lstrContrasena, ablnTimeOut)
        If lblnFueEnviadoMens Then
            If astrNroDoc = My.Resources.MenCtaCobro Then
                MobjObjetoWin.StrMensaje = My.Resources.MenCtaCobro
            End If
            MobjObjetoWin.SRegistreMens(astrArchivoAEnviar, adblIdCliente, astrNroDoc)
            If astrArchivoAEnviar <> MobjObjetoWin.StrArchivoExterno Then
                My.Computer.FileSystem.DeleteFile(astrArchivoAEnviar)
            End If
        End If
        Return lblnFueEnviadoMens
    End Function
#End Region

#Region "Segundo plano"
    Private Sub Bgw_DoWork(sender As System.Object, e As DoWorkEventArgs) Handles MbgwCorreo.DoWork
        Dim lblnTimeOut = False
        MblnHayConeccion = FblnHayInternet()
        MblnTimeOut = False
        If MblnHayConeccion Then
            GblnEnviandoEmail = True
            If My.Computer.FileSystem.DirectoryExists(GstrTrayEmails) Then
                SLimpieCarpeta(GstrTrayEmails, "*.pdf")
                SLimpieCarpeta(GstrTrayEmails, "*.txt")
            End If
            GobjPanDat.SControleProcesoObj(True)
            Select Case MobjObjetoWin.EnuTipoCorreo
                Case EnuTipoCorreoE.EnuFactAuto
                    MblnEnvioCorreo = FblnEnvioFacsAutoMes(MblnTimeOut)
                Case EnuTipoCorreoE.EnuSoloMens, EnuTipoCorreoE.EnuArchExt
                    MblnEnvioCorreo = FblnEnvioArchivo(MblnTimeOut)
                Case EnuTipoCorreoE.EnuRecibos
                    MblnEnvioCorreo = FblnEnvioRecibos(MblnTimeOut)
                Case EnuTipoCorreoE.EnuCobroPers
                    MblnEnvioCorreo = FblnEnvioCuentasCobro(MblnTimeOut)
                Case Else
                    MblnEnvioCorreo = FblnEnvioDoc(MblnTimeOut)
            End Select
            If Not MblnEnvioCorreo Then
                MblnHayConeccion = FblnEstaConectado()
            End If
            GobjPanDat.SControleProcesoObj(False)
        End If
    End Sub

    Private Sub Bgw_RunWorkerCompleted(sender As System.Object,
            e As RunWorkerCompletedEventArgs) Handles MbgwCorreo.RunWorkerCompleted
        Dim lstrMens As String
        If MblnCancelando Then
            lstrMens = "El Proceso fue cancelado por el Usuario!"
            If MobjObjetoWin.EnuTipoCorreo <> EnuTipoCorreoE.EnuFactAuto Then
                MobjObjetoWin.SRegistreUltimo()
            End If
        ElseIf Not MblnHayConeccion Then
            lstrMens = "No está conectado a Internet en este momento. " &
                "Intente de nuevo cuando se haya restablecido!"
            MsgBox(lstrMens, vbOKOnly, "Sin conexión as internet")
        ElseIf MblnTimeOut Then
            lstrMens = "No es posible enviar mensajes en este momento. " &
                " Intentelo de nuevo más tarde!"
        ElseIf Not MblnEnvioCorreo Then
            lstrMens = "Se presentó un error. Reinicie el programa. Si se vuelve a " &
                "presentar, comuniquese con soporte"
        ElseIf Not IsNothing(e.Error) Then
            lstrMens = e.Error.ToString
        Else
            If MblnEnvioOk Then
                lstrMens = "El Proceso terminó exitosamente!"
            Else
                If MobjObjetoWin.EnuTipoCorreo = EnuTipoCorreoE.EnuFactAuto AndAlso
                        ClsOrionCop.FdtbFacsMesEnviarEmail.Rows.Count = 0 Then
                    lstrMens = "Las facturas ya fueron enviadas!!"
                Else
                    lstrMens = "No hay Documentos para ser enviados!"
                End If
            End If
        End If
        txtNotifica.Content = String.Empty
        GblnEnviandoEmail = False
        MblnEnviandoEmail = False
        SEstablezcaWinConsultando()
        HbttAceptar.IsEnabled = True
        HbttCancelar.IsEnabled = True
        cboTipoCorreo.SelectedIndex = EnuTipoCorreoE.None
        If MblnHayDocNoEnviados Then
            SMuestreCorreosNoEnviados()
            lstrMens = "No fue posible enviar algún Documento!"
            MblnHayDocNoEnviados = False
        End If
        SCree()
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Private Sub Bgw_ProgressChanged(sender As Object,
            e As ProgressChangedEventArgs) Handles MbgwCorreo.ProgressChanged
        pgbCorreo.Value = (e.ProgressPercentage)
        txtProceso.Content = e.ProgressPercentage.ToString() & "%"
        txtNotifica.Content = MstrNotifica
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError As Boolean
        If TypeOf lelmElemento Is MenuItem Then
            Try
                If lelmElemento.Name = "MnuCuentaOrigenEmail" Then
                    Dim lwinVentana As New WinCuentaCorreoOrigen
                    lwinVentana.ShowDialog()
                    lblCuentaCorreoOri.Content = "Cuenta origen de Correo : " &
                            GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.
                            ObjEmailOrigenStr.ObjValorPro
                    MsgBox("Si ha realizado algún cambio en la cuenta de correo origen, reinicie el programa para que los cambios tengan efecto.",
                            vbOKOnly, "Información")
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
            If Not String.IsNullOrEmpty(lstrMens) Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuInformacion)
            End If
        Else
            SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
        End If
        End Try
        End If
    End Sub

    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            If lelmElemento.Name = "bttExaminar" Then
                SSeleccioneArchivo()
            ElseIf lelmElemento.Name = "bttEncontrarCliHist" Then
                MblnBuscarHistoClie = True
                SBuscar()
            ElseIf lelmElemento.Name = "bttEncontrarCli" Then
                MblnBuscarHistoClie = False
                SBuscar()
            ElseIf lelmElemento.Name = "bttMostrarTodos" Then
                txtIdClienteHis.Text = My.Resources.Todos
            End If
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
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If (TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is PasswordBox) AndAlso
                Not HblnSeEstaCerrando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                GobjPanDat.SControleProcesoObj(True)
                With MobjObjetoWin
                    Select Case lelmElemento.Name
                        Case "txtIdCliente"
                            SRegistreCliente()
                        Case "txtArchivo"
                            StcValidValido(EnuValidEntrada.enuArch) =
                                    MobjObjetoWin.FblnEsValidoArchivoExt(txtArchivo.Text, lstrMens)
                        Case "txtMensaje"
                            StcValidValido(EnuValidEntrada.enuMens) =
                                    MobjObjetoWin.FblnEsValidoMensaje(txtMensaje.Text, lstrMens)
                        Case "txtAsunto"
                            StcValidValido(EnuValidEntrada.enuAsun) =
                                    MobjObjetoWin.FblnEsValidoAsunto(txtAsunto.Text, lstrMens)
                    End Select
                End With
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
                If lblnNoHayError Then
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                    GobjPanDat.SControleProcesoObj(False)
                Else
                    GobjPanDat.SControleProcesoObj(False, True)
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub

    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is ComboBox Then
            If Not (MblnPoblandoCombo OrElse HblnCargandoForma OrElse HblnMostrandoDatos) Then
                Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                Try
                    GobjPanDat.SControleProcesoObj(True)
                    Dim lblnValido As Boolean
                    With MobjObjetoWin
                        Select Case lelmElemento.Name
                            Case "cboTipoCorreo"
                                If Not MblnDesdeDoc Then
                                    SVacie()
                                End If
                                lblnValido = MobjObjetoWin.FblnEsValidoTipoCorreo(
                                        cboTipoCorreo.SelectedIndex)
                                SAsigneValidoTodo(lblnValido)
                                STextosDefecto()
                                SVisibiliceCtls()
                            Case "cboTipoCorreoH"
                                SMuestreHistoricoPorAsunto()
                            Case "cboPredioAgr"
                                lblnValido = MobjObjetoWin.FblnEsValidoIdPreAgr(
                                        cboPredioAgr.SelectedItem)
                                StcValidValido(EnuValidEntrada.enuPreAgr) = lblnValido
                                SPuebleDocs()
                                STextosDefecto()
                            Case "cboNroDoc"
                                lblnValido = MobjObjetoWin.FblnEsValidoNroDoc(cboNroDoc.SelectedItem,
                                        lstrMens)
                                StcValidValido(EnuValidEntrada.enuDoc) = lblnValido
                                STextosDefecto()
                        End Select
                    End With
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
                    If lblnNoHayError Then
                        If Not String.IsNullOrEmpty(lstrMens) Then
                            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                        End If
                        GobjPanDat.SControleProcesoObj(False)
                    Else
                        GobjPanDat.SControleProcesoObj(False, True)
                        SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                    End If
                End Try
            End If
        End If
    End Sub

    Private Sub Dtp_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles _
            dtpFecIni.SelectedDateChanged, dtpFecFin.SelectedDateChanged, dtpDesde.SelectedDateChanged,
            dtpHasta.SelectedDateChanged, dtpFechaDesde.SelectedDateChanged,
            dtpFechaHasta.SelectedDateChanged
        Dim lstrMens = String.Empty
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is DatePicker Then
            If Not IsNothing(MobjObjetoWin) Then
                If lelmElemento.Name = "dtpFecIni" Then
                    StcValidValido(EnuValidEntrada.enuFecIni) =
                            MobjObjetoWin.FblnEsValFecIni(dtpFecIni.SelectedDate, lstrMens)
                ElseIf lelmElemento.Name = "dtpFecFin" Then
                    StcValidValido(EnuValidEntrada.enuFecFin) =
                        MobjObjetoWin.FblnEsValFecFin(dtpFecFin.SelectedDate, lstrMens)
                ElseIf lelmElemento.Name = "dtpFechaDesde" OrElse lelmElemento.Name =
                        "dtpFechaHasta" Then
                    SMuestreHistoricoPorAsunto()
                End If
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
                SMuestreDatos()
            End If
            If lelmElemento.Name = "dtpDesde" OrElse lelmElemento.Name = "dtpHasta" Then
                SMuestreCorreos()
            End If
        End If
    End Sub

    Private Sub ChkSeguro_Click(sender As Object, e As RoutedEventArgs) Handles chkSeguro.Click
        If TypeOf sender Is CheckBox Then
            StcValidValido(EnuValidEntrada.EnuFacMes) =
                    MobjObjetoWin.FblnEsValidoFactAuto(chkSeguro.IsChecked)
            SMuestreDatos()
        End If
    End Sub

    Private Sub Txt_TextChanged(sender As Object, e As TextChangedEventArgs) Handles _
            txtIdClienteHis.TextChanged, txtDiasVen.TextChanged
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            If lelmElemento.Name = "txtIdClienteHis" Then
                SMuestreCorreos()
            ElseIf lelmElemento.Name = "txtDiasVen" Then
                Dim lstrMens = String.Empty
                Dim lblnEsValido = MobjObjetoWin.FblnEsValidoDiasVen(txtDiasVen.Text,
                        lstrMens)
                If lblnEsValido Then
                    MentDiasCobro = CInt(txtDiasVen.Text)
                End If
                StcValidValido(EnuValidEntrada.EnuDiasCobPer) = lblnEsValido
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0,
                            EnuSeveridadNot.EnuInformacion)
                End If
            Else
                '
            End If
        End If
    End Sub

    Private Sub TabControl_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        If TypeOf e.Source Is TabControl Then
            If tbiHistorico.IsSelected Then
                SMuestreCorreos()
            ElseIf tbiHistoricoAsunto.IsSelected Then
                SMuestreHistoricoPorAsunto()
            End If
        End If
    End Sub

    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
                dgrHistorico.MouseRightButtonUp
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            Dim ldblIdcliente As Double
            ldrvFilaActual = ldgrActual.SelectedItem
            If Not IsNothing(ldrvFilaActual) AndAlso ldgrActual.SelectedIndex >= 0 Then
                ldblIdcliente = ldrvFilaActual("IdTerceroCliente")
                txtIdClienteHis.Text = ldblIdcliente
            End If
        End If
    End Sub

    Private Sub ClsFormInterface_Closing(sender As Object, e As CancelEventArgs)
        If MbgwCorreo.IsBusy Then
            Dim lstrMens = "Aún se están enviando Mensajes. No es posible cerrar la ventana!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            e.Cancel = True
        Else
            GblnCorreoOn = False
        End If
    End Sub

    Protected Overrides Sub EwinClosed(sender As Object, e As EventArgs)
        GenuTamanoIcono = MenuTamanoIcono
        If WinPadre IsNot Nothing Then
            If WinPadre.Visibility <> Visibility.Visible Then
                WinPadre.Visibility = Visibility.Visible
                WinPadre.SRefresqueWin()
            End If
        End If
    End Sub
#End Region
End Class