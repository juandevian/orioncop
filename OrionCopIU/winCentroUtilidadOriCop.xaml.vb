Imports Microsoft.Win32
Public Class WinCentroUtilidadOriCop
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        EnuDiasPersu = 0
        EnuDiasPreju
        EnuDiasJurid
        EnuDiasPerdi
        EnuRedondeCP
        EnuRedondeIM
        EnuRedondeGr
        EnuPlazoFacM
        EnuTarReteIva
        EnuCtaCaja
        EnuCtaIntMDb
        EnuCtaRetFte
        EnuCtaRetIva
        EnuCtaRetIca
        EnuCtaAntRec
        EnuCtaDctoPP
        EnuCtaIngSinId
        EnuCtaImpAsu
        EnuMedPagDef
        EnuIdAppCont
        EnuTipoInterfaz
        EnuCodigoEmp
        EnuTipoTerCaja
        EnuResFac
        EnuFecRes
        EnuFecFinRes
        EnuRanIni
        EnuRanFin
        EnuContrasAPI
        EnuIdUsuAPI
        EnuIdProvEFac
        EnuURL
        EnuSubirFac
        EnuResCon
        EnuFecResCon
        EnuPrefCon
        EnuRanIniCon
        EnuRanFinCon
        EnuExiFechaHoyDocs
        EnuPieFacUno
        EnuPieFacDos
        EnuPieFacTres
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsCentroUtilOriCop = Nothing
    Private MblnPoblandoComboBox As Boolean = False
    Private MnuAbrirCtasBanco As MenuItem = Nothing
    Private MnuAutorizarEFac As MenuItem = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomCenUtilOri
#End Region

#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuCentroUtilOrionCop
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneCtlsRestringidos()
        SCargueForma(EnuElementosAdicionalesDef.None, 42, Nothing, txtPersuasivo, False)
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
            ObjObjetoWin = GobjParametros
        End If
        MobjObjetoWin = ObjObjetoWin
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando Then
            MobjObjetoWin.SInicialicePorDefecto()
        End If
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub

    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.EnuRedondeCP) = lblRedondeoCP
        StcValidaControl(EnuValidEntrada.EnuRedondeIM) = lblRedondeoIM
        StcValidaControl(EnuValidEntrada.EnuRedondeGr) = lblRedondeoGral
        StcValidaControl(EnuValidEntrada.EnuDiasPersu) = lblPersuasivo
        StcValidaControl(EnuValidEntrada.EnuDiasPreju) = lblPrejuridico
        StcValidaControl(EnuValidEntrada.EnuDiasJurid) = lblJuridico
        StcValidaControl(EnuValidEntrada.EnuDiasPerdi) = lblPerdida
        StcValidaControl(EnuValidEntrada.EnuPlazoFacM) = lblPlazoFacturaManual
        StcValidaControl(EnuValidEntrada.EnuTarReteIva) = lblTarifaReteIva
        StcValidaControl(EnuValidEntrada.EnuCtaCaja) = lblCtaCaja
        StcValidaControl(EnuValidEntrada.EnuCtaIntMDb) = lblCtaIntMoraDb
        StcValidaControl(EnuValidEntrada.EnuCtaRetFte) = lblCtaRetefuente
        StcValidaControl(EnuValidEntrada.EnuCtaRetIva) = lblCtaReteIva
        StcValidaControl(EnuValidEntrada.EnuCtaRetIca) = lblCtaReteIca
        StcValidaControl(EnuValidEntrada.EnuCtaAntRec) = lblCtaAnticipos
        StcValidaControl(EnuValidEntrada.EnuCtaDctoPP) = lblCtaDctosPP
        StcValidaControl(EnuValidEntrada.EnuTipoInterfaz) = lblTipoInterfaz
        StcValidaControl(EnuValidEntrada.EnuIdAppCont) = lblAppContable
        StcValidaControl(EnuValidEntrada.EnuCodigoEmp) = lblCodioEmp
        StcValidaControl(EnuValidEntrada.EnuCtaImpAsu) = lblCtaImptosAsum
        StcValidaControl(EnuValidEntrada.EnuCtaIngSinId) = lblCtaIngNoIdentificados
        StcValidaControl(EnuValidEntrada.EnuMedPagDef) = lblMedioPagoDefecto
        StcValidaControl(EnuValidEntrada.EnuCodigoEmp) = lblCodioEmp
        StcValidaControl(EnuValidEntrada.EnuTipoTerCaja) = lblTerceroCaja
        StcValidaControl(EnuValidEntrada.EnuResFac) = lblNroResolucionFra
        StcValidaControl(EnuValidEntrada.EnuFecRes) = lblFechaResolucionFra
        StcValidaControl(EnuValidEntrada.EnuFecFinRes) = lblFechaFinRes
        StcValidaControl(EnuValidEntrada.EnuRanIni) = lblNroIni
        StcValidaControl(EnuValidEntrada.EnuRanFin) = lblNroFin
        StcValidaControl(EnuValidEntrada.EnuIdProvEFac) = lblProvEFac
        StcValidaControl(EnuValidEntrada.EnuIdUsuAPI) = lblIdUsuarioProvEFac
        StcValidaControl(EnuValidEntrada.EnuContrasAPI) = lblContrasena
        StcValidaControl(EnuValidEntrada.EnuSubirFac) = chkSubirFact
        StcValidaControl(EnuValidEntrada.EnuURL) = lblUrl
        StcValidaControl(EnuValidEntrada.EnuResCon) = lblNroResolucionCon
        StcValidaControl(EnuValidEntrada.EnuFecResCon) = lblFechaResolucionCon
        StcValidaControl(EnuValidEntrada.EnuPrefCon) = lblPrefCon
        StcValidaControl(EnuValidEntrada.EnuRanIniCon) = lblNroIniCon
        StcValidaControl(EnuValidEntrada.EnuRanFinCon) = lblNroFinCon
        StcValidaControl(EnuValidEntrada.EnuExiFechaHoyDocs) = chkExigeFechaHoyDocs
        StcValidaControl(EnuValidEntrada.EnuPieFacUno) = lblPieFra1
        StcValidaControl(EnuValidEntrada.EnuPieFacDos) = lblPieFra2
        StcValidaControl(EnuValidEntrada.EnuPieFacTres) = lblPieFra3
        '
        SVisibiliceBttEncontrar(EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando)
        '
        If Not GobjParametros.ObjAutorizaEFacBln.ObjValorPro Then
            cnvEFac.Visibility = Visibility.Collapsed
        End If
        SEstablezcaToolTipGral()
        SPuebleComboBoxes()
        HbttAceptar.TabIndex = 100
        HbttCancelar.TabIndex = 101
    End Sub

    Protected Overrides Sub SMuestreDatos()
        SMuestreInfTab()
        txtFechaUltCausacion.Content = Format(MobjObjetoWin.ObjFechaUltCausacionGralDtm.ObjValorPro,
                GCSTRFMTFECHASIMPLE)
        SVisibiliceCodEmp()
        SEstablezcaToolTipLeido()
        SValide()
    End Sub

    Protected Overrides Sub SRefresqueWin()
        MyBase.SRefresqueWin()
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        Dim lcolCuentasBanco = GobjParametros.FcolCuentasBanco
        If FblnEstanTodosBien() AndAlso lcolCuentasBanco.Count > 0 Then
            GobjParametros.SVerifiqueApp(True, True)
        End If
        HmnuModificar.IsEnabled = True
        HbttModificar.IsEnabled = True
    End Sub

    Protected Overrides Sub SValide()
        With MobjObjetoWin
            StcValidValido(EnuValidEntrada.EnuRedondeCP) = .ObjBaseRedondeoCPByt.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuRedondeIM) = .ObjBaseRedondeoIntMoraDbl.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuRedondeGr) = .ObjBaseRedondeoGeneralDbl.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuDiasPersu) = .ObjDiasParaPersuasivoShr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuDiasPreju) = .ObjDiasParaPrejuridicoShr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuDiasJurid) = .ObjDiasParaJuridicoShr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuDiasPerdi) = .ObjDiasParaPerdidaShr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuPlazoFacM) = .ObjPlazoDefectoFacManualShr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuTarReteIva) = .ObjTarifaReteIvaDbl.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuCtaCaja) = .ObjIdCtaCajaStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuCtaImpAsu) = .ObjIdCtaImptosAsumidosStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuCtaIntMDb) = .ObjIdCtaIntMoraDbStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuCtaRetFte) = .ObjIdCtaReteFuenteStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuCtaRetIva) = .ObjIdCtaReteIvaStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuCtaRetIca) = .ObjIdCtaReteIcaStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuCtaAntRec) = .ObjIdCtaAnticiposRecibidosStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuCtaDctoPP) = .ObjIdCtaDescuentosPPStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuTipoInterfaz) = .ObjTipoInterfazByt.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuIdAppCont) = .ObjIdAppContableByt.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuCodigoEmp) = .ObjCodigoEmpShr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuMedPagDef) = .ObjIdMedioPagoDefectoByt.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuCtaIngSinId) = .ObjIdCtaIngPorIdentificarStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuTipoTerCaja) = .ObjTipoTerceroCajaByt.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuResFac) = .ObjNumeroResolFacturaStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuFecRes) = .ObjFechaResolucionFactDtm.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuFecFinRes) = .ObjFechaVenceResolFactDtm.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuRanIni) = .ObjRangoFraIniEnt.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuRanFin) = .ObjRangoFraFinEnt.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuIdProvEFac) = .FblnEsValidoIdProvEFac
            StcValidValido(EnuValidEntrada.EnuIdUsuAPI) = .FblnEsValidoIdUsuarioProvEFac
            StcValidValido(EnuValidEntrada.EnuContrasAPI) = .FblnEsValidoContAPIEFac
            StcValidValido(EnuValidEntrada.EnuSubirFac) = .FblnEsValidoSubirFac
            StcValidValido(EnuValidEntrada.EnuURL) = .FblnEsValidoUrl
            StcValidValido(EnuValidEntrada.EnuResCon) = .ObjNumeroResolContiStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuFecResCon) = .ObjFechaResolucionContDtm.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuPrefCon) = .ObjPrefijoFactContStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuRanIniCon) = .ObjRangoFraConIniEnt.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuRanFinCon) = .ObjRangoFraConFinEnt.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuExiFechaHoyDocs) = .ObjExigeFechaHoyDocsBln.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuPieFacUno) = .ObjPieFacturaUnoStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuPieFacDos) = .ObjPieFacturaDosStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuPieFacTres) = .ObjPieFacturaTresStr.BlnEsValido
        End With
        '
        FblnEstanTodosBien()
        SHabiliteBotonesTlb()
    End Sub

    Protected Overrides Sub SRegistre()
        If Not HblnCargandoForma Then
            With MobjObjetoWin
                .ObjDiasParaPersuasivoShr.ObjValorPro = txtPersuasivo.Text
                .ObjDiasParaPrejuridicoShr.ObjValorPro = txtPrejuridico.Text
                .ObjDiasParaJuridicoShr.ObjValorPro = txtJuridico.Text
                .ObjDiasParaPerdidaShr.ObjValorPro = txtPerdida.Text
                .ObjBaseRedondeoCPByt.ObjValorPro = txtRedondeoCP.Text
                .ObjBaseRedondeoIntMoraDbl.ObjValorPro = txtRedondeoIM.Text
                .ObjBaseRedondeoGeneralDbl.ObjValorPro = txtRedondeoGral.Text
                .ObjPlazoDefectoFacManualShr.ObjValorPro = txtPlazoFacManual.Text
                .ObjIdMedioPagoDefectoByt.ObjValorPro = cboMedioPagoDef.SelectedIndex
                .ObjTarifaReteIvaDbl.ObjValorPro = FdblTasa(txtTarifaReteIva.Text)
                .ObjIdCtaAnticiposRecibidosStr.ObjValorPro = txtCtaAnticipos.Text
                .ObjIdCtaCajaStr.ObjValorPro = txtCtaCaja.Text
                .ObjIdCtaIntMoraDbStr.ObjValorPro = txtCtaIntMoraDb.Text
                .ObjIdCtaImptosAsumidosStr.ObjValorPro = txtCtaImptosAsum.Text
                .ObjIdCtaIngPorIdentificarStr.ObjValorPro = txtCtaIngNoIdentificados.Text
                .ObjIdCtaReteFuenteStr.ObjValorPro = txtCtaRetefuente.Text
                .ObjIdCtaReteIcaStr.ObjValorPro = txtCtaReteIca.Text
                .ObjIdCtaReteIvaStr.ObjValorPro = txtCtaReteIva.Text
                .ObjIdCtaDescuentosPPStr.ObjValorPro = txtCtaDctosPP.Text
                .ObjTipoInterfazByt.ObjValorPro = cboTipoInterfaz.SelectedIndex
                .ObjIdAppContableByt.ObjValorPro = cboAppContable.SelectedIndex
                .ObjInformaSaldoTotalDespuesRCBln.ObjValorPro = chkInformarSaldoTotal.IsChecked
                .ObjConsolidaItemsFacBln.ObjValorPro = chkConsItemsFac.IsChecked
                .ObjExigeFechaHoyCajaBln.ObjValorPro = chkExigeFechaHoyCaja.IsChecked
                .ObjExigeFechaHoyDocsBln.ObjValorPro = chkExigeFechaHoyDocs.IsChecked
                .ObjPermiteAnticipoPorServicioBln.ObjValorPro = chkPermiteAnticipoPorSer.IsChecked
                .ObjNotificacionesSonorasBln.ObjValorPro = chkPermiteNotificacionSonora.IsChecked
                .ObjNoMostrarAyudaBln.ObjValorPro = chkNoMostrarAyuda.IsChecked
                .ObjServicioIdActivoBln.ObjValorPro = chkActivarServicioId.IsChecked
                .ObjFirmaRCeMail.ObjValorPro = chkFirmaRCeMail.IsChecked
                .ObjTipoTerceroCajaByt.ObjValorPro = cboTerceroCaja.SelectedIndex
                .ObjIdProvEFacByt.ObjValorPro = cboProvEFac.SelectedIndex
                .ObjNumeroResolFacturaStr.ObjValorPro = txtNroResolucionFra.Text
                .ObjFechaResolucionFactDtm.ObjValorPro = dtpFechaRes.SelectedDate
                .ObjFechaVenceResolFactDtm.ObjValorPro = dtpFechaFinRes.SelectedDate
                .ObjRangoFraFinEnt.ObjValorPro = txtNroFraFin.Text
                .ObjRangoFraIniEnt.ObjValorPro = txtNroFraIni.Text
                .ObjNumeroResolContiStr.ObjValorPro = txtNroResolucionCon.Text
                .ObjFechaResolucionContDtm.ObjValorPro = dtpFechaCon.SelectedDate
                .ObjPrefijoFactContStr.ObjValorPro = txtPrefCon.Text
                .ObjRangoFraConFinEnt.ObjValorPro = txtNroFraFinCon.Text
                .ObjRangoFraConIniEnt.ObjValorPro = txtNroFraIniCon.Text
                .ObjPieFacturaDosStr.ObjValorPro = txtPieFac2.Text
                .ObjPieFacturaUnoStr.ObjValorPro = txtPieFac1.Text
                .ObjPieFacturaTresStr.ObjValorPro = txtPieFac3.Text
                SRegistre_ProvEFac()
            End With
        End If
        SValide()
    End Sub

    Protected Overrides Sub SConfigureMenuesPropios()
        MnuAbrirCtasBanco = FmnuiMenuItemPan("MnuAbrirCtasBanco", "Abrir Cuentas de _Bancos", 1, "")
        MnuAutorizarEFac = FmnuiMenuItem("MnuAutorizarEFac", "Autorizar Interfaz EFactura", "RecMnuItemSec")
        Dim lsep = New Separator
        HmnuAcciones.Items.Insert(7, MnuAbrirCtasBanco)
        HmnuAcciones.Items.Insert(8, MnuAutorizarEFac)
        HmnuAcciones.Items.Insert(9, lsep)
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    Protected Overrides Sub SCree()
        MyBase.SCree()
        '
    End Sub
    Protected Overrides Sub SHabiliteMenues()
        MyBase.SHabiliteMenues()
        With GobjPanorama.ObjUsuarioActual
            Dim lblnTienePermiso = .FblnTienePermiso(EnuIdClasesPanDef.EnuCuentaBanco,
                    EnuIdAccionDef.enuConsultar)
            SHabiliteMenuItem(lblnTienePermiso, MnuAbrirCtasBanco)
            lblnTienePermiso = (GstrIdUsuario = GCSTRUSUARIOU)
            SHabiliteMenuItem(lblnTienePermiso, MnuAutorizarEFac)
        End With
        bttAbrirCtasBancos.IsEnabled = MnuAbrirCtasBanco.IsEnabled
        If MobjObjetoWin.ObjAutorizaEFacBln.ObjValorPro Then
            MnuAutorizarEFac.Visibility = Visibility.Collapsed
        End If
    End Sub
    Protected Overrides Sub SModifique()
        MyBase.SModifique()
        SVisibiliceBttEncontrar(EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando)
        If MobjObjetoWin.ObjAutorizaEFacBln.ObjValorPro Then
            txtPieFac1.Style = FindResource("RecCtlNoHabilitado")
        End If
    End Sub
    Protected Overrides Sub SGuarde()
        Dim lblnCreando = EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando
        MyBase.SGuarde()
        If lblnCreando AndAlso EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            SValide()
            If FblnEstanTodosBien() Then
                Dim lobjCarpeta As ClsCarpeta = GobjPanorama.ObjCarpetaActual
                Dim lobjCentroUtil As ClsCentroUtilidad = lobjCarpeta.ObjCentroUtilidadActual
                Dim lstrNomCop = lobjCentroUtil.ObjNombreCentroUtilStr.ObjValorPro
                Dim lstrMens = "La opciones de la Copropiedad " & lstrNomCop &
                        " se definieron exitosamente!"
                SLevanteEveNoti(lstrMens, "", 0, EnuSeveridadNot.EnuInformacion)
            End If
        End If
    End Sub
    Protected Overrides Sub SFinaliceOperacion()
        MyBase.SFinaliceOperacion()
        SVisibiliceBttEncontrar(EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando)
    End Sub
#End Region

#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            Me.Cursor = Cursors.Wait
            If IsNothing(HwinBusqueda) Then
                HwinBusqueda = New WinBusqueda With {
                    .WinPadre = Me
                }
            End If
            If FblnDefinioBusqueda() Then
                HwinBusqueda.ShowDialog()
            End If
            HwinBusqueda = Nothing
            Me.Cursor = Cursors.Arrow
        End If
    End Sub
    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        SDefineCuentaCont()
        Return True
    End Function
    Private Sub SDefineCuentaCont()
        Dim lstrCamposMostrar As String() = {ClsIdCuentaContStr.SstrNombreCampoBd,
                                                 ClsNombreCuentaStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCuentaStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdCuentaContStr.SstrNombreCampoBd
        Dim lstrTabla As String = ClsCuentaContabilidad.SstrNombreTabla
        Dim lstrFiltro As String = ClsIdCarpetaCuentaShr.SstrNombreCampoBd & " = " & GshrIdCarpeta &
                " AND " & lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Nombre Cuenta", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
    Private Sub SBusqueCuenta(abttBoton As Button)
        SBuscar()
        If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
            Dim ltxtCuentaCon As TextBox = Nothing
            Select Case abttBoton.Name
                Case "bttEncontrarCtaAnt"
                    ltxtCuentaCon = txtCtaAnticipos
                Case "bttEncontrarCtaCaja"
                    ltxtCuentaCon = txtCtaCaja
                Case "bttEncontrarCtaMoraDb"
                    ltxtCuentaCon = txtCtaIntMoraDb
                Case "bttEncontrarCtaRetFte"
                    ltxtCuentaCon = txtCtaRetefuente
                Case "bttEncontrarCtaRetIco"
                    ltxtCuentaCon = txtCtaReteIca
                Case "bttEncontrarCtaRetIva"
                    ltxtCuentaCon = txtCtaReteIva
                Case "bttEncontrarCtaDsctoPP"
                    ltxtCuentaCon = txtCtaDctosPP
                Case "bttEncontrarCtaIngNoIden"
                    ltxtCuentaCon = txtCtaIngNoIdentificados
                Case "bttEncontrarCtaImptosAsum"
                    ltxtCuentaCon = txtCtaImptosAsum
            End Select
            ltxtCuentaCon.Text = StrResultadoBusqueda
            SRegistreEntrada(ltxtCuentaCon)
        End If
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SAdicioneCtlsRestringidos()
        SAdicioneControlRestringido(tbcCentUtil)
        SAdicioneControlRestringido(bttAbrirCtasBancos)
        SAdicioneControlRestringido(bttEncontrarCtaAnt)
        SAdicioneControlRestringido(bttEncontrarCtaCaja)
        SAdicioneControlRestringido(bttEncontrarCtaDsctoPP)
        SAdicioneControlRestringido(bttEncontrarCtaMoraDb)
        SAdicioneControlRestringido(bttEncontrarCtaRetFte)
        SAdicioneControlRestringido(bttEncontrarCtaRetIco)
        SAdicioneControlRestringido(bttEncontrarCtaIngNoIden)
        SAdicioneControlRestringido(bttEncontrarCtaImptosAsum)
        SAdicioneControlRestringido(bttEncontrarCtaRetIva)
    End Sub

    Private Sub SPuebleComboBoxes()
        MblnPoblandoComboBox = True
        Dim ldrwDataRow = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuAppContable)
        SPuebleComboBox(ldrwDataRow, cboAppContable)
        ldrwDataRow = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuTipoInterfaz)
        SPuebleComboBox(ldrwDataRow, cboTipoInterfaz)
        ' Poblar combo Medios de Pago
        ldrwDataRow = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuMediosPago)
        SPuebleComboBox(ldrwDataRow, cboMedioPagoDef)
        ldrwDataRow = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuTipoTercero)
        SPuebleComboBox(ldrwDataRow, cboTerceroCaja)
        ldrwDataRow = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuProvEFac)
        SPuebleComboBox(ldrwDataRow, cboProvEFac)
        MblnPoblandoComboBox = False
    End Sub

    Private Sub SVisibiliceBttEncontrar(ablnVisible As Boolean)
        Dim lwvVisibilidad As Visibility = Visibility.Collapsed
        If ablnVisible Then lwvVisibilidad = Visibility.Visible
        bttEncontrarCtaAnt.Visibility = lwvVisibilidad
        bttEncontrarCtaCaja.Visibility = lwvVisibilidad
        bttEncontrarCtaDsctoPP.Visibility = lwvVisibilidad
        bttEncontrarCtaMoraDb.Visibility = lwvVisibilidad
        bttEncontrarCtaRetFte.Visibility = lwvVisibilidad
        bttEncontrarCtaRetIco.Visibility = lwvVisibilidad
        bttEncontrarCtaIngNoIden.Visibility = lwvVisibilidad
        bttEncontrarCtaImptosAsum.Visibility = lwvVisibilidad
        bttEncontrarCtaRetIva.Visibility = lwvVisibilidad
    End Sub
    Private Sub SVisibiliceCodEmp()
        If MobjObjetoWin.ObjIdAppContableByt.ObjValorPro = EnuAppConta.EnuContaPyme Then
            lblCodioEmp.Visibility = Visibility.Visible
            txtCodigoEmp.Visibility = Visibility.Visible
        Else
            lblCodioEmp.Visibility = Visibility.Collapsed
            txtCodigoEmp.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub SMuestreInfTab()
        SMuestreGrales()
        SMuestrePieFactura()
        SMuestreContabilidad()
        If MobjObjetoWin.ObjAutorizaEFacBln.ObjValorPro Then
            tbiAspectosTrib.Visibility = Visibility.Visible
            SMuestreTribut()
        Else
            tbiAspectosTrib.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub SMuestreGrales()
        With MobjObjetoWin
            txtPersuasivo.Text = .ObjDiasParaPersuasivoShr.ObjValorPro
            txtPrejuridico.Text = .ObjDiasParaPrejuridicoShr.ObjValorPro
            txtJuridico.Text = .ObjDiasParaJuridicoShr.ObjValorPro
            txtPerdida.Text = .ObjDiasParaPerdidaShr.ObjValorPro
            txtRedondeoCP.Text = .ObjBaseRedondeoCPByt.ObjValorPro
            txtRedondeoIM.Text = .ObjBaseRedondeoIntMoraDbl.ObjValorPro
            txtRedondeoGral.Text = .ObjBaseRedondeoGeneralDbl.ObjValorPro
            txtTotalArea.Content = .ObjTotalAreaCopropDec.ObjValorPro
            txtBaseCalculoCP.Content = .ObjTotalAreaPondDec.ObjValorPro
            txtPlazoFacManual.Text = .ObjPlazoDefectoFacManualShr.ObjValorPro
            cboMedioPagoDef.SelectedIndex = .ObjIdMedioPagoDefectoByt.ObjValorPro
            chkInformarSaldoTotal.IsChecked = .ObjInformaSaldoTotalDespuesRCBln.ObjValorPro
            chkConsItemsFac.IsChecked = .ObjConsolidaItemsFacBln.ObjValorPro
            chkExigeFechaHoyCaja.IsChecked = .ObjExigeFechaHoyCajaBln.ObjValorPro
            chkExigeFechaHoyDocs.IsChecked = .ObjExigeFechaHoyDocsBln.ObjValorPro
            chkPermiteAnticipoPorSer.IsChecked = .ObjPermiteAnticipoPorServicioBln.ObjValorPro
            chkPermiteNotificacionSonora.IsChecked = .ObjNotificacionesSonorasBln.ObjValorPro
            chkNoMostrarAyuda.IsChecked = .ObjNoMostrarAyudaBln.ObjValorPro
            chkActivarServicioId.IsChecked = .ObjServicioIdActivoBln.ObjValorPro
            chkFirmaRCeMail.IsChecked = .ObjFirmaRCeMail.ObjValorPro
        End With
    End Sub

    Private Sub SMuestrePieFactura()
        txtPieFac1.Text = MobjObjetoWin.ObjPieFacturaUnoStr.ObjValorPro
        txtPieFac2.Text = MobjObjetoWin.ObjPieFacturaDosStr.ObjValorPro
        txtPieFac3.Text = MobjObjetoWin.ObjPieFacturaTresStr.ObjValorPro
    End Sub

    Private Sub SMuestreContabilidad()
        With MobjObjetoWin
            txtCtaAnticipos.Text = .ObjIdCtaAnticiposRecibidosStr.ObjValorPro
            txtNomCtaAnticipos.Content = .ObjIdCtaAnticiposRecibidosStr.StrNombreCuenta
            txtCtaCaja.Text = .ObjIdCtaCajaStr.ObjValorPro
            txtNomCtaCaja.Content = .ObjIdCtaCajaStr.StrNombreCuenta
            txtCtaDctosPP.Text = .ObjIdCtaDescuentosPPStr.ObjValorPro
            txtNomCtaDctosPP.Content = .ObjIdCtaDescuentosPPStr.StrNombreCuenta
            txtCtaImptosAsum.Text = .ObjIdCtaImptosAsumidosStr.ObjValorPro
            txtNomCtaImptosAsum.Content = .ObjIdCtaImptosAsumidosStr.StrNombreCuenta
            txtCtaIntMoraDb.Text = .ObjIdCtaIntMoraDbStr.ObjValorPro
            txtNomCtaIntMoraDb.Content = .ObjIdCtaIntMoraDbStr.StrNombreCuenta
            txtCtaIngNoIdentificados.Text = .ObjIdCtaIngPorIdentificarStr.ObjValorPro
            txtNomCtaIngNoIdent.Content = .ObjIdCtaIngPorIdentificarStr.StrNombreCuenta
            txtCtaRetefuente.Text = .ObjIdCtaReteFuenteStr.ObjValorPro
            txtNomCtaRetefuente.Content = .ObjIdCtaReteFuenteStr.StrNombreCuenta
            txtCtaReteIca.Text = .ObjIdCtaReteIcaStr.ObjValorPro
            txtNomCtaReteIca.Content = .ObjIdCtaReteIcaStr.StrNombreCuenta
            txtCtaReteIva.Text = .ObjIdCtaReteIvaStr.ObjValorPro
            txtNomCtaReteIva.Content = .ObjIdCtaReteIvaStr.StrNombreCuenta
            cboTipoInterfaz.SelectedIndex = .ObjTipoInterfazByt.ObjValorPro
            cboAppContable.SelectedIndex = .ObjIdAppContableByt.ObjValorPro
            txtCodigoEmp.Text = .ObjCodigoEmpShr.ObjValorPro
            cboTerceroCaja.SelectedIndex = .ObjTipoTerceroCajaByt.ObjValorPro
        End With
    End Sub

    Private Sub SMuestreTribut()
        With MobjObjetoWin
            txtTarifaReteIva.Text = .ObjTarifaReteIvaDbl.ToString
            cboProvEFac.SelectedIndex = .ObjIdProvEFacByt.ObjValorPro
            txtNroResolucionFra.Text = .ObjNumeroResolFacturaStr.ObjValorPro
            dtpFechaRes.SelectedDate = .ObjFechaResolucionFactDtm.ObjValorPro
            dtpFechaFinRes.SelectedDate = .ObjFechaVenceResolFactDtm.ObjValorPro
            txtPrefRes.Content = .StrPrefijoResStr
            txtNroFraFin.Text = .ObjRangoFraFinEnt.ObjValorPro
            txtNroFraIni.Text = .ObjRangoFraIniEnt.ObjValorPro
            txtNroResolucionCon.Text = .ObjNumeroResolContiStr.ObjValorPro
            dtpFechaCon.SelectedDate = .ObjFechaResolucionContDtm.ObjValorPro
            txtPrefCon.Text = .ObjPrefijoFactContStr.ObjValorPro
            txtNroFraFinCon.Text = .ObjRangoFraConFinEnt.ObjValorPro
            txtNroFraIniCon.Text = .ObjRangoFraConIniEnt.ObjValorPro
            chkSubirFact.IsChecked = .ObjSubirFacBln.ObjValorPro
            txtUsuarioProvEFac.Text = .ObjIdUsuarioProvEFacStr.ObjValorPro
            pwbContrasenaAPI.Password = ClsPanorama.FstrContrasena(.ObjContrasenaAPIEFacStr)
            txtUrl.Text = .ObjURLStr.ObjValorPro
        End With
    End Sub

    Private Sub SRegistreEntrada(aobjControl As Control)
        Dim lblnRegistroIn = FblnRegistroEntrada(aobjControl)
        If Not lblnRegistroIn Then
            If TypeOf aobjControl Is TextBox Then
                lblnRegistroIn = True
                With MobjObjetoWin
                    Select Case aobjControl.Name
                        Case "txtPlazoFacManual"
                            .ObjPlazoDefectoFacManualShr.ObjValorPro = txtPlazoFacManual.Text
                        Case "txtTarifaReteIva"
                            .ObjTarifaReteIvaDbl.ObjValorPro = FdblTasa(txtTarifaReteIva.Text)
                        Case "txtRedondeoCP"
                            .ObjBaseRedondeoCPByt.ObjValorPro = txtRedondeoCP.Text
                        Case "txtRedondeoGral"
                            .ObjBaseRedondeoGeneralDbl.ObjValorPro = txtRedondeoGral.Text
                        Case "txtRedondeoIM"
                            .ObjBaseRedondeoIntMoraDbl.ObjValorPro = txtRedondeoIM.Text
                            .ObjBaseRedondeoIntMoraDbl.ObjValorPro = txtRedondeoIM.Text
                        Case "txtCodigoEmp"
                            .ObjCodigoEmpShr.ObjValorPro = txtCodigoEmp.Text
                        Case "txtNroResolucionFra"
                            .ObjNumeroResolFacturaStr.ObjValorPro = txtNroResolucionFra.Text
                        Case "txtNroFraFin"
                            .ObjRangoFraFinEnt.ObjValorPro = txtNroFraFin.Text
                        Case "txtNroFraIni"
                            .ObjRangoFraIniEnt.ObjValorPro = txtNroFraIni.Text
                        Case "txtNroResolucionCon"
                            .ObjNumeroResolContiStr.ObjValorPro = txtNroResolucionCon.Text
                        Case "txtPrefCon"
                            .ObjPrefijoFactContStr.ObjValorPro = txtPrefCon.Text
                        Case "txtNroFraFinCon"
                            .ObjRangoFraConFinEnt.ObjValorPro = txtNroFraFinCon.Text
                        Case "txtNroFraIniCon"
                            .ObjRangoFraConIniEnt.ObjValorPro = txtNroFraIniCon.Text
                        Case "txtPieFac1"
                            .ObjPieFacturaUnoStr.ObjValorPro = txtPieFac1.Text
                        Case "txtPieFac2"
                            .ObjPieFacturaDosStr.ObjValorPro = txtPieFac2.Text
                        Case "txtPieFac3"
                            .ObjPieFacturaTresStr.ObjValorPro = txtPieFac3.Text
                        Case Else
                            lblnRegistroIn = False
                    End Select
                    If Not lblnRegistroIn Then
                        If MobjObjetoWin.ObjIdProvEFacByt.ObjValorPro > 0 Then
                            SRegistreEntradaEFac(aobjControl)
                        End If
                    End If
                End With
            ElseIf TypeOf aobjControl Is PasswordBox Then
                SRegistreEntradaEFac(aobjControl)
            ElseIf TypeOf aobjControl Is DatePicker Then
                If aobjControl.Name = "dtpFechaRes" Then
                    MobjObjetoWin.ObjFechaResolucionFactDtm.ObjValorPro = dtpFechaRes.SelectedDate
                ElseIf aobjControl.Name = "dtpFechaFinRes" Then
                    MobjObjetoWin.ObjFechaVenceResolFactDtm.ObjValorPro = dtpFechaFinRes.SelectedDate
                ElseIf aobjControl.Name = "dtpFechaCon" Then
                    MobjObjetoWin.ObjFechaResolucionContDtm.ObjValorPro = dtpFechaCon.SelectedDate
                End If
            End If
        End If
        SMuestreDatos()
    End Sub

    Private Function FblnRegistroEntrada(aobjControl As Control) As Boolean
        Dim lblnRegistro = True
        If TypeOf aobjControl Is TextBox Then
            With MobjObjetoWin
                If aobjControl.Name.ToString.StartsWith("txtCta") Then
                    SRegistreCtaContable(aobjControl)
                Else
                    Select Case aobjControl.Name
                        Case "txtJuridico"
                            .ObjDiasParaJuridicoShr.ObjValorPro = txtJuridico.Text
                        Case "txtPerdida"
                            .ObjDiasParaPerdidaShr.ObjValorPro = txtPerdida.Text
                        Case "txtPersuasivo"
                            .ObjDiasParaPersuasivoShr.ObjValorPro = txtPersuasivo.Text
                        Case "txtPrejuridico"
                            .ObjDiasParaPrejuridicoShr.ObjValorPro = txtPrejuridico.Text
                        Case Else
                            lblnRegistro = False
                    End Select
                End If
            End With
        Else
            lblnRegistro = False
        End If
        Return lblnRegistro
    End Function

    Private Sub SRegistreCtaContable(actrlControl As Control)
        Dim lstrMens = String.Empty
        Dim ctrlTextBox As TextBox = TryCast(actrlControl, TextBox)
        Dim lstrIdCtaContable As String = ctrlTextBox.Text
        If ClsPanorama.FblnEsValidoNumero(lstrIdCtaContable, 1, Double.MaxValue, True,
                        EnuTipoValor.EnuDouble) Then
            Dim lobjCtaCont As New ClsCuentaContabilidad(GobjPanorama.ObjCarpetaActual,
                    EnuModoInstanciaObjDef.EnuNavegable)
            Dim lobjValorLlave As Object() = {GshrIdCarpeta, lstrIdCtaContable}
            lobjCtaCont.SAbra(lobjValorLlave)
            If Not lobjCtaCont.BlnExiste Then
                Dim lstrMensaje = "La Cuenta Contable ingresada no existe. " &
                        "Desea crearla ahora?"
                If MsgBox(lstrMensaje, MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                        "Crear Cuenta Contable ?") = MsgBoxResult.Yes Then
                    Dim lstrNombreCuenta As String = FstrNomCtaCont(actrlControl)
                    lobjCtaCont.SCreeObj({GshrIdCarpeta, lstrIdCtaContable})
                    lobjCtaCont.ObjIdCarpetaCuentaShr.ObjValorPro = GshrIdCarpeta
                    lobjCtaCont.ObjIdCuentaContStr.ObjValorPro = lstrIdCtaContable
                    lobjCtaCont.ObjNombreCuentaStr.ObjValorPro = lstrNombreCuenta
                    lobjCtaCont.SActualice(True)
                End If
            End If
            With MobjObjetoWin
                Select Case actrlControl.Name
                    Case "txtCtaAnticipos"
                        .ObjIdCtaAnticiposRecibidosStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaAnticipos.Content = lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
                    Case "txtCtaCaja"
                        .ObjIdCtaCajaStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaCaja.Content = lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
                    Case "txtCtaIntMoraDb"
                        .ObjIdCtaIntMoraDbStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaIntMoraDb.Content = lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
                    Case "txtCtaRetefuente"
                        .ObjIdCtaReteFuenteStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaRetefuente.Content = lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
                    Case "txtCtaReteIca"
                        .ObjIdCtaReteIcaStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaReteIca.Content = lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
                    Case "txtCtaReteIva"
                        .ObjIdCtaReteIvaStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaReteIva.Content = lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
                    Case "txtCtaDctosPP"
                        .ObjIdCtaDescuentosPPStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaDctosPP.Content = lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
                    Case "txtCtaIngNoIdentificados"
                        .ObjIdCtaIngPorIdentificarStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaIngNoIdent.Content =
                                    lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
                    Case "txtCtaImptosAsum"
                        .ObjIdCtaImptosAsumidosStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaImptosAsum.Content =
                                    lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
                End Select
            End With
        Else
            lstrMens = "El Valor ingresado no es válido!"
        End If
        SMuestreDatos()
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Private Function FstrNomCtaCont(actlTextBox As TextBox) As String
        Dim lstrNombreCuenta = String.Empty
        Select Case actlTextBox.Name
            Case "txtCtaAnticipos"
                lstrNombreCuenta = "Anticipos Recibidos"
            Case "txtCtaCaja"
                lstrNombreCuenta = "Caja"
            Case "txtCtaIntMoraDb"
                lstrNombreCuenta = "CxC Intereses de Mora"
            Case "txtCtaRetefuente"
                lstrNombreCuenta = "Retención en la Fuente"
            Case "txtCtaReteIca"
                lstrNombreCuenta = "Retención ICA"
            Case "txtCtaReteIva"
                lstrNombreCuenta = "Retención IVA"
            Case "txtCtaDctosPP"
                lstrNombreCuenta = "Descuentos por Pronto Pago"
            Case "txtCtaIngNoIdentificados"
                lstrNombreCuenta = "Ingresos por Identificar"
            Case "txtCtaImptosAsum"
                lstrNombreCuenta = "Impuestos Asumidos"
        End Select
        Return lstrNombreCuenta
    End Function

    Private Sub SRegistreEntradaEFac(aobjControl As Control)
        Select Case aobjControl.Name
            Case "txtUsuarioProvEFac"
                MobjObjetoWin.ObjIdUsuarioProvEFacStr.ObjValorPro = txtUsuarioProvEFac.Text
            Case "pwbContrasenaAPI"
                MobjObjetoWin.ObjContrasenaAPIEFacStr.ObjValorPro =
                        ClsPanorama.FstrContrasena(True, pwbContrasenaAPI.Password.Trim)
            Case "txtUrl"
                MobjObjetoWin.ObjURLStr.ObjValorPro = txtUrl.Text
        End Select
    End Sub

    ''' <summary>
    ''' Establece el ToolTip de los Controles a un valor constante
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SEstablezcaToolTipGral()
        txtRedondeoCP.ToolTip = My.Resources.TTRedoCP
        lblRedondeoCP.ToolTip = My.Resources.TTRedoCP
        txtRedondeoGral.ToolTip = My.Resources.TTRedoOpeMonet
        lblRedondeoGral.ToolTip = My.Resources.TTRedoOpeMonet
        txtRedondeoIM.ToolTip = My.Resources.TTRedoIM & vbCrLf & My.Resources.TTRedoOpeMonet
        lblRedondeoIM.ToolTip = My.Resources.TTRedoIM & vbCrLf & My.Resources.TTRedoOpeMonet
    End Sub

    ''' <summary>
    ''' Establece el ToolTip al contenido del control despues de leido
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SEstablezcaToolTipLeido()
        txtNomCtaAnticipos.ToolTip = txtNomCtaAnticipos.Content
        txtNomCtaCaja.ToolTip = txtNomCtaCaja.Content
        txtNomCtaIntMoraDb.ToolTip = txtNomCtaIntMoraDb.Content
        txtNomCtaRetefuente.ToolTip = txtNomCtaRetefuente.Content
        txtNomCtaReteIca.ToolTip = txtNomCtaReteIca.Content
        txtNomCtaReteIva.ToolTip = txtNomCtaReteIva.Content
        txtNomCtaDctosPP.ToolTip = txtNomCtaDctosPP.Content
        txtNomCtaIngNoIdent.ToolTip = txtNomCtaIngNoIdent.Content
    End Sub

    Private Sub SAbraAutoEFac()
        Dim lwinVentana As New WinAutorizaDscto With {
            .Title = "Autoriza Interfaz EFactura",
            .BlnEfact = True
        }
        lwinVentana.ShowDialog()
        If GblnOK Then
            MobjObjetoWin.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
            MobjObjetoWin.ObjAutorizaEFacBln.ObjValorPro = GblnOK
            MobjObjetoWin.SActualice(True)
            MnuAutorizarEFac.Visibility = Visibility.Collapsed
            cnvEFac.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub SAbraCtasBanco()
        Dim lwinVentana As New WinCuentasBanco With {
            .WinPadre = Me}
        lwinVentana.ShowDialog()
    End Sub

    Private Sub SRegistre_ProvEFac()
        With MobjObjetoWin
            .ObjIdCarpetaEFacShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtilEFacShr.ObjValorPro = GshrIdCentroUtil
            .ObjIdProveedorEFacEnt.ObjValorPro = cboProvEFac.SelectedIndex
            .ObjSubirFacBln.ObjValorPro = chkSubirFact.IsChecked
            .ObjIdUsuarioProvEFacStr.ObjValorPro = txtUsuarioProvEFac.Text
            .ObjContrasenaAPIEFacStr.ObjValorPro = ClsPanorama.FstrContrasena(True, pwbContrasenaAPI.Password.Trim)
            .ObjURLStr.ObjValorPro = txtUrl.Text
        End With
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Dim lbttBoton As Button = lelmElemento
            Select Case lelmElemento.Name
                Case "bttAbrirCtasBancos"
                    SAbraCtasBanco()
                Case "bttEncontrarCtaAnt", "bttEncontrarCtaCaja", "bttEncontrarCtaMoraDb",
                            "bttEncontrarCtaRetFte", "bttEncontrarCtaRetIco", "bttEncontrarCtaRetIva",
                            "bttEncontrarCtaDsctoPP", "bttEncontrarCtaIngNoIden",
                            "bttEncontrarCtaImptosAsum"
                    SBusqueCuenta(lbttBoton)
            End Select
        End If
    End Sub

    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            Select Case lelmElemento.Name
                Case "MnuAbrirCtasBanco"
                    SAbraCtasBanco()
                Case "MnuAutorizarEFac"
                    SAbraAutoEFac()
            End Select
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
        If Not HblnSeEstaCerrando Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is Control Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    SRegistreEntrada(lelmElemento)
                End If
            End If
        End If
    End Sub

    Private Sub OnTxtCambio(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            If ltxtTextBox.Text.StartsWith("?") Then
                StrResultadoBusqueda = String.Empty
                If ltxtTextBox.Equals(txtCtaAnticipos) OrElse ltxtTextBox.Equals(txtCtaCaja) OrElse
                        ltxtTextBox.Equals(txtCtaIntMoraDb) OrElse ltxtTextBox.Equals(txtCtaRetefuente) OrElse
                        ltxtTextBox.Equals(txtCtaReteIca) OrElse ltxtTextBox.Equals(txtCtaReteIva) OrElse
                        ltxtTextBox.Equals(txtCtaDctosPP) Then
                    SBuscar()
                End If
                If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                    ltxtTextBox.Text = StrResultadoBusqueda
                    SRegistreEntrada(ltxtTextBox)
                End If
            End If
        End If
    End Sub

    Private Sub Chk_Click(sender As Object, e As RoutedEventArgs) Handles _
            chkInformarSaldoTotal.Click, chkPermiteAnticipoPorSer.Click,
            chkPermiteNotificacionSonora.Click, chkNoMostrarAyuda.Click,
            chkActivarServicioId.Click, chkExigeFechaHoyCaja.Click, chkExigeFechaHoyDocs.Click,
            chkSubirFact.Click, chkConsItemsFac.Click, chkFirmaRCeMail.Click
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is CheckBox Then
                With MobjObjetoWin
                    Select Case lelmElemento.Name
                        Case "chkInformarSaldoTotal"
                            .ObjInformaSaldoTotalDespuesRCBln.ObjValorPro = chkInformarSaldoTotal.IsChecked
                        Case "chkPermiteAnticipoPorSer"
                            .ObjPermiteAnticipoPorServicioBln.ObjValorPro = chkPermiteAnticipoPorSer.IsChecked
                        Case "chkConsItemsFac"
                            .ObjConsolidaItemsFacBln.ObjValorPro = chkConsItemsFac.IsChecked
                        Case "chkExigeFechaHoyCaja"
                            .ObjExigeFechaHoyCajaBln.ObjValorPro = chkExigeFechaHoyCaja.IsChecked
                        Case "chkExigeFechaHoyDocs"
                            .ObjExigeFechaHoyDocsBln.ObjValorPro = chkExigeFechaHoyDocs.IsChecked
                        Case "chkPermiteNotificacionSonora"
                            .ObjNotificacionesSonorasBln.ObjValorPro = chkPermiteNotificacionSonora.IsChecked
                        Case "chkNoMostrarAyuda"
                            .ObjNoMostrarAyudaBln.ObjValorPro = chkNoMostrarAyuda.IsChecked
                        Case "chkActivarServicioId"
                            .ObjServicioIdActivoBln.ObjValorPro = chkActivarServicioId.IsChecked
                        Case "chkFirmaRCeMail"
                            .ObjFirmaRCeMail.ObjValorPro = chkFirmaRCeMail.IsChecked
                        Case "chkSubirFact"
                            .ObjSubirFacBln.ObjValorPro = chkSubirFact.IsChecked
                    End Select
                End With
            End If
            SMuestreDatos()
        End If
    End Sub

    Private Sub Cbo_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles _
            cboAppContable.SelectionChanged, cboTipoInterfaz.SelectionChanged,
            cboMedioPagoDef.SelectionChanged, cboTerceroCaja.SelectionChanged,
            cboProvEFac.SelectionChanged, cboProvEFac.SelectionChanged
        If Not MblnPoblandoComboBox Then
            If TypeOf sender Is ComboBox Then
                Dim lcboSender As ComboBox = sender
                With MobjObjetoWin
                    Select Case lcboSender.Name
                        Case "cboAppContable"
                            .ObjIdAppContableByt.ObjValorPro = cboAppContable.SelectedIndex
                            SVisibiliceCodEmp()
                        Case "cboTipoInterfaz"
                            .ObjTipoInterfazByt.ObjValorPro = cboTipoInterfaz.SelectedIndex
                        Case "cboMedioPagoDef"
                            .ObjIdMedioPagoDefectoByt.ObjValorPro = cboMedioPagoDef.SelectedIndex
                        Case "cboTerceroCaja"
                            .ObjTipoTerceroCajaByt.ObjValorPro = cboTerceroCaja.SelectedIndex
                        Case "cboProvEFac"
                            .ObjIdProvEFacByt.ObjValorPro = cboProvEFac.SelectedIndex
                            .ObjProveedorEFac.ObjIdProveedorEFacEnt.ObjValorPro = cboProvEFac.SelectedIndex
                    End Select
                End With
            End If
            SValide()
        End If
    End Sub

    Private Sub Tbc_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) _
            Handles tbcCentUtil.SelectionChanged
        If MobjObjetoWin IsNot Nothing Then
            SMuestreInfTab()
        End If
    End Sub
#End Region
End Class