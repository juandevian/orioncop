Imports System.Windows.Controls
Public Class WinServicios
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        EnuBaseRetFte
        EnuBaseRetIca
        EnuCtaCr
        EnuCtaDb
        EnuCtaDev
        EnuCtaIva
        EnuCtaMora
        EnuCantiPerio
        EnuPeriodoIni
        EnuTipoBase
        EnuTipoSer
        EnuNombre
        EnuTarIva
        EnuTarRetFte
        EnuTarRetIca
        EnuEsFactProgram
        EnuGeneraProgram
        EnuEsServId
        EnuTipoTerCtaCr
        EnuIdTerCtaCe
        EnuDiaFra
        EnuDiasVen
        EnuDiasGra
        enuGraciaFinMes
        EnuFactProp
        EnuConcepto
        EnuModuloSer
        EnuModoCausaMora
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsServicio = Nothing
    '
    Private ReadOnly MobjAno As ClsAno = Nothing
    Private MblnPoblandoComboBox As Boolean = False
    Private MnuConsultarHistorico As MenuItem = Nothing
    Private MnuReporte As MenuItem = Nothing
    Private MnuCalcular As MenuItemPan = Nothing
    Private MnuLimpiar As MenuItemPan = Nothing
    Private MblnSectoresAsignados As Boolean = False
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomSer
    Private ReadOnly MenuTamanoIcono As EnuTamanoIconos
    Private ReadOnly MshrIdServicio As Short = 0
    Private MshrIdModuloSel As Short = 0
    Private MobjModuloSel As ClsModuloServicio = Nothing
    Private MdtbModulosSer As DataTable = Nothing
    ' Calcular
    Private ReadOnly MsepMenu As New Separator
#End Region

#Region "Constructor"
    Friend Sub New(ashrIdServicio As Short, aobjAno As ClsAno)
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuServicios
        MobjAno = aobjAno
        MshrIdServicio = ashrIdServicio
        MenuTamanoIcono = GenuTamanoIcono
        GenuTamanoIcono = EnuTamanoIconos.EnuMediano
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolCtrlsLlave As New Collection From {
            txtIdServicio
        }
        SAdicioneCtrlsRestringidos()
        SCargueForma(EnuElementosAdicionalesDef.None, 29, lcolCtrlsLlave, txtNombreServicio, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
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
        Dim lshrIdAno As Short
        If MobjAno IsNot Nothing Then
            ' Servicio Anual
            lshrIdAno = MobjAno.ObjIdAnoShr.ObjValorPro
        Else
            ' Servicio Permanente
            lshrIdAno = 0
        End If
        MobjObjetoWin = New ClsServicio(MobjAno, EnuModoInstanciaObjDef.enuNavegable)
        If Not MobjObjetoWin.FblnEstaVacioOrigenDatos Then
            If MshrIdServicio > 0 Then
                Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lshrIdAno,
                        MshrIdServicio}
                MobjObjetoWin.SAbra(lobjValorLlave)
            Else
                MobjObjetoWin.SVayaAlPrimero()
            End If
        End If
        ObjObjetoWin = MobjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuBaseRetFte) = lblBaseRetefuente
        StcValidaControl(EnuValidEntrada.enuBaseRetIca) = lblBaseReteIca
        StcValidaControl(EnuValidEntrada.EnuModoCausaMora) = lblModoCausaMora
        StcValidaControl(EnuValidEntrada.enuCtaCr) = lblCtaCredito
        StcValidaControl(EnuValidEntrada.enuCtaDb) = lblCtaDebito
        StcValidaControl(EnuValidEntrada.enuCtaDev) = lblCtaDevol
        StcValidaControl(EnuValidEntrada.enuCtaMora) = lblCtaMora
        StcValidaControl(EnuValidEntrada.enuCtaIva) = lblCtaIva
        StcValidaControl(EnuValidEntrada.enuPeriodoIni) = lblPeriodoInicio
        StcValidaControl(EnuValidEntrada.enuCantiPerio) = lblCantidadPeri
        StcValidaControl(EnuValidEntrada.enuNombre) = lblNomServicio
        StcValidaControl(EnuValidEntrada.enuTarIva) = lblTarifaIva
        StcValidaControl(EnuValidEntrada.enuTarRetFte) = lblTarifaRetefuente
        StcValidaControl(EnuValidEntrada.enuTarRetIca) = lblTarifaReteIca
        StcValidaControl(EnuValidEntrada.enuTipoBase) = lblTipoBaseCalculo
        StcValidaControl(EnuValidEntrada.enuTipoSer) = lblTipoServicio
        StcValidaControl(EnuValidEntrada.enuEsFactProgram) = chkFacturacionProgramable
        StcValidaControl(EnuValidEntrada.enuGeneraProgram) = chkEsCalculado
        StcValidaControl(EnuValidEntrada.enuEsServId) = chkEsServicioId
        StcValidaControl(EnuValidEntrada.enuTipoTerCtaCr) = lblTipoTerCtaCrédito
        StcValidaControl(EnuValidEntrada.enuIdTerCtaCe) = lblTerCtaCrédito
        StcValidaControl(EnuValidEntrada.enuDiaFra) = lblDiaFactura
        StcValidaControl(EnuValidEntrada.enuDiasVen) = lblDiasVence
        StcValidaControl(EnuValidEntrada.enuDiasGra) = lblPeriodoGracia
        StcValidaControl(EnuValidEntrada.enuGraciaFinMes) = chkGraciaFinMes
        StcValidaControl(EnuValidEntrada.EnuFactProp) = chkFactuPropYPreAgr
        StcValidaControl(EnuValidEntrada.enuConcepto) = lblConceptoSer
        StcValidaControl(EnuValidEntrada.enuModuloSer) = lblModCon
        '
        MdtbModulosSer = MobjObjetoWin.FdtbModulosServicio
        SPuebleComboBoxes()
        SEstablezcaToolTipGral()
        SVisibiliceControles()
        ' 
        HbttAceptar.TabIndex = 50
        HbttCancelar.TabIndex = 51
        If MobjObjetoWin.BlnEsCuotaAdministracion AndAlso
                MobjObjetoWin.ObjMiAno.ObjEstaCerradoAnoBln.ObjValorPro Then
            SHabiliteBotonTlb(False, HbttModificar)
            SHabiliteMenuItem(False, HmnuModificar)
        End If
    End Sub
    Protected Overrides Sub SMuestreDatos()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        Dim lstrMens = String.Empty
        HblnMostrandoDatos = True
        If lblnNoHayDatos AndAlso EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            lstrMens = "No hay Servicios para ser mostrados!"
            SLevanteEveNoti(lstrMens, "", 0, EnuSeveridadNot.EnuInformacion)
            txtIdServicio.IsEnabled = False
        End If
        With MobjObjetoWin
            txtIdServicio.Text = .ObjIdServicioShr.ToString
            txtNombreServicio.Text = .ObjNombreServicioStr.ObjValorPro
            txtAno.Content = .ObjIdAno_ServicioShr.ObjValorPro
            chkActivo.IsChecked = .ObjEstaActivoServicioBln.ObjValorPro
        End With
        SVisibiliceControles()
        SMuestreGenerales()
        SMuestreValores()
        SMuestreParContables()
        Title = My.Resources.Servicio & My.Resources.DosPuntosEspacio
        If Not String.IsNullOrEmpty(txtIdServicio.Text) Then
            Title &= txtIdServicio.Text
            Title &= " " & MobjObjetoWin.ObjNombreServicioStr.ObjValorPro
        End If
        SValide()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            Dim lenuSevNot = EnuSeveridadNot.EnuInformacion
            If Not BlnVentanaAux OrElse Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                SVerifiqueSectoresModulo(lstrMens)
                If Not String.IsNullOrEmpty(lstrMens) Then
                    lenuSevNot = EnuSeveridadNot.EnuFalta
                End If
            End If
            If String.IsNullOrEmpty(lstrMens) Then
                SVerifiqueEstado(lstrMens)
            End If
            SHabiliteReporte()
            If txtIdServicio.Focus Then
                txtIdServicio.SelectAll()
            End If
            If Not String.IsNullOrEmpty(lstrMens) Then
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            End If
        End If
        HblnMostrandoDatos = False
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntrada.EnuBaseRetFte) =
                        .ObjBaseMinimaReteFuenteDec.BlnEsValido
                StcValidValido(EnuValidEntrada.enuBaseRetIca) = .ObjBaseMinimaReteIcaDec.BlnEsValido
                StcValidValido(EnuValidEntrada.EnuModoCausaMora) =
                        .ObjModoCausaInteresesByt.BlnEsValido
                StcValidValido(EnuValidEntrada.enuCtaCr) = .ObjCodigoCuentaCrStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuCtaDb) = .ObjCodigoCuentaDbStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuCtaDev) = .ObjCodigoCuentaDevStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuCtaMora) = .ObjCodigoCuentaMoraStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuCtaIva) = .ObjCodigoCuentaIvaStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuPeriodoIni) = .ObjPeriodoInicioStr.BlnEsValido
                StcValidValido(EnuValidEntrada.EnuCantiPerio) =
                        .ObjCantPeriodos_ServicioShr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuNombre) = .ObjNombreServicioStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuTarIva) = .ObjTarifaIvaDbl.BlnEsValido
                StcValidValido(EnuValidEntrada.enuTarRetFte) = .ObjTarifaRetFteDbl.BlnEsValido
                StcValidValido(EnuValidEntrada.enuTarRetIca) = .ObjTarifaRetIcaDbl.BlnEsValido
                StcValidValido(EnuValidEntrada.enuTipoBase) = .ObjTipoBaseCalculoByt.BlnEsValido
                StcValidValido(EnuValidEntrada.enuTipoSer) = .ObjIdTipoServicioByt.BlnEsValido
                StcValidValido(EnuValidEntrada.EnuEsFactProgram) =
                        .ObjEsFactProgramableBln.BlnEsValido
                StcValidValido(EnuValidEntrada.enuGeneraProgram) = .ObjGeneraProgramBln.BlnEsValido
                StcValidValido(EnuValidEntrada.enuEsServId) = .ObjEsServicioIdBln.BlnEsValido
                StcValidValido(EnuValidEntrada.EnuTipoTerCtaCr) =
                        .ObjIdTipoTerCtaCrSerByt.BlnEsValido
                StcValidValido(EnuValidEntrada.enuIdTerCtaCe) = .ObjIdTerceroCtaCrDbl.BlnEsValido
                StcValidValido(EnuValidEntrada.enuDiaFra) = .ObjDiaFacturaShr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuDiasVen) = .ObjDiasVencimientoShr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuDiasGra) = .ObjDiasGraciaShr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuGraciaFinMes) = .ObjGraciaFinMesBln.BlnEsValido
                StcValidValido(EnuValidEntrada.EnuFactProp) = .ObjFactAPropYPreAgrBln.BlnEsValido
                StcValidValido(EnuValidEntrada.enuConcepto) = .ObjConceptoServicioStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuModuloSer) = FblnModulosSerOk()
            End With
        End If
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
        SHabiliteAcciones()
        If Not StcValidValido(EnuValidEntrada.enuModuloSer) Then
            Dim lblnNo = FblnModulosSerOk()
        End If
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjIdCarpeta_ServicioShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCentroUtil_ServicioShr.ObjValorPro = GshrIdCentroUtil
            .ObjNombreServicioStr.ObjValorPro = txtNombreServicio.Text
            .ObjConceptoServicioStr.ObjValorPro = txtConceptoSer.Text
            .ObjIdAno_ServicioShr.ObjValorPro = txtAno.Content
            If .ObjEsAjusteBln.ObjValorPro Then
                chkEstaAjustadoEsAjuste.Content = My.Resources.EsAjuste
            End If
            .ObjEstaActivoServicioBln.ObjValorPro = chkActivo.IsChecked
            .ObjDiaFacturaShr.ObjValorPro = txtDiaFactura.Text
            .ObjDiasVencimientoShr.ObjValorPro = txtDiasVence.Text
            .ObjVenceFinMesBln.ObjValorPro = chkVenceFinMes.IsChecked
            .ObjDiasGraciaShr.ObjValorPro = txtDiasGracia.Text
            .ObjGraciaFinMesBln.ObjValorPro = chkGraciaFinMes.IsChecked
            .ObjFactAPropYPreAgrBln.ObjValorPro = chkFactuPropYPreAgr.IsChecked
            .ObjModoCausaInteresesByt.ObjValorPro = cboModoCausaMora.SelectedIndex
            .ObjEsFactProgramableBln.ObjValorPro = chkFacturacionProgramable.IsChecked
            .ObjEsServicioIdBln.ObjValorPro = chkEsServicioId.IsChecked
            .ObjGeneraProgramBln.ObjValorPro = chkEsCalculado.IsChecked
            .ObjTipoBaseCalculoByt.ObjValorPro = cboTipoBaseCalculo.SelectedIndex
            .ObjEstaGenaradaProgramBln.ObjValorPro = chkEstaCalculado.IsChecked
            If tbiValores.Visibility = Visibility.Visible Then
                .ObjPeriodoInicioStr.ObjValorPro = cboAnoPeriodo.SelectedItem & cboMesPeriodo.SelectedItem
                .ObjCantPeriodos_ServicioShr.ObjValorPro = txtCantidadPer.Text
            End If
            .ObjCodigoCuentaDbStr.ObjValorPro = txtCtaDebito.Text
            .ObjCodigoCuentaCrStr.ObjValorPro = txtCtaCredito.Text
            .ObjCodigoCuentaMoraStr.ObjValorPro = txtCtaMora.Text
            .ObjCodigoCuentaDevStr.ObjValorPro = txtCtaDevolucion.Text
            .ObjCodigoCuentaIvaStr.ObjValorPro = txtCtaIva.Text
            .ObjTarifaIvaDbl.ObjValorPro = FdblTasa(txtTarifaIva.Text)
            .ObjEsExcluidoIvaBln.ObjValorPro = chkEsExcluido.IsChecked
            .ObjTarifaRetFteDbl.ObjValorPro = FdblTasa(txtTarifaretefuente.Text)
            .ObjTarifaRetIcaDbl.ObjValorPro = FdblTasa(txtTarifaReteIca.Text)
            .ObjBaseMinimaReteFuenteDec.ObjValorPro = txtBaseretefuente.Text
            .ObjBaseMinimaReteIcaDec.ObjValorPro = txtBasereteica.Text
        End With
        SValide()
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        ' Adicionar Menues Reportes
        MnuReporte = FmnuiMenuItem("MnuReportes", "R_eportes", "RecMnuItemPriInf")
        Dim lmnuVlrPorPredio = FmnuiMenuItem("MnuValorPredio", "_Valores a Cobrar", "RecMnuItemSec")
        lmnuVlrPorPredio.ToolTip = "Genera un Listado del Valor que debe pagar cada Predio."
        MnuReporte.Items.Add(lmnuVlrPorPredio)
        MenuVen.Items.Insert(1, MnuReporte)
        MnuCalcular = FmnuiMenuItemPan("MnuCalcular", "_Calcular Valores a cobrar", 2, "")
        MnuLimpiar = FmnuiMenuItemPan("MnuLimpiar", "_Eliminar Valores a cobrar", 3, "")
        MnuConsultarHistorico = FmnuiMenuItem("MnuConsultarHistorico",
                "Consultar _Historico del Servicio", "RecMnuItemSec")
        Dim lentIndice = HmnuAcciones.Items.Count - 1
        HmnuAcciones.Items.Insert(lentIndice, MsepMenu)
        HmnuAcciones.Items.Insert(lentIndice, MnuCalcular)
        HmnuAcciones.Items.Insert(lentIndice, MnuLimpiar)
        HmnuAcciones.Items.Insert(lentIndice, MnuConsultarHistorico)
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    Protected Overrides Sub SHabiliteMenues()
        With GobjPanorama.ObjUsuarioActual
            MyBase.SHabiliteMenues()
            If MobjObjetoWin.ObjIdTipoServicioByt.ObjValorPro =
                    EnuTipoServicio.EnuPermanente AndAlso
                    MobjObjetoWin.ObjGeneraProgramBln.ObjValorPro Then
                MnuConsultarHistorico.Visibility = Visibility.Visible
                MnuConsultarHistorico.IsEnabled = True
            Else
                MnuConsultarHistorico.Visibility = Visibility.Collapsed
            End If
            If MobjObjetoWin.BlnEsImportado Then
                SHabiliteMenuItem(False, MnuCalcular)
                SHabiliteMenuItem(False, MnuLimpiar)
            End If
            If MobjObjetoWin.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
                SHabiliteMenuItem(False, MnuCalcular)
            ElseIf Not MobjObjetoWin.ObjGeneraProgramBln.ObjValorPro Then
                SHabiliteMenuItem(False, MnuCalcular)
            End If
        End With
    End Sub
    Protected Overrides Sub SCree()
        Dim lblnCrear As Boolean = True
        If MobjAno IsNot Nothing Then
            lblnCrear = GobjParametros.FblnPuedeCrearSerAnual(MobjAno.ObjIdAnoShr.ObjValorPro)
            If lblnCrear Then
                If GobjParametros.ColAnos.Count = 1 AndAlso MobjAno.ColServiciosAno.Count = 0 Then
                    lblnCrear = MsgBox("Ya está debidamente parametrizado el Año?", MsgBoxStyle.YesNo,
                        "Seguir?") = vbYes
                End If
            Else
                SLevanteEveNoti("Para crear un Servicio Anual, el Año anterior debe estar " &
                        "cerrado y tener al menos un Servicio!", "", 0,
                         EnuSeveridadNot.EnuInformacion)
            End If
        End If
        If lblnCrear Then
            MobjObjetoWin.SCreeObj(Nothing)
            EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando
            SPuebleCombosPeriodo()
            SMuestreDatos()
            SFrmAdicione()
            SLevanteEveNoti("", "", 0, EnuSeveridadNot.EnuOk)
            SVisibiliceBotonesEncontrar()
            SVisibiliceControles()
            SValide()
            If MnuReporte.Visibility = Visibility.Visible Then
                MnuReporte.IsEnabled = False
            End If
            txtIdServicio.IsEnabled = False
            txtNombreServicio.Focus()
        End If
    End Sub
    Protected Overrides Sub SModifique()
        MyBase.SModifique()
        SVisibiliceBotonesEncontrar()
        SVisibiliceControles()
        SPuebleCombosPeriodo()
        SMuestreValores()
        SMuestreParContables()
        If MnuReporte.Visibility = Visibility.Visible Then
            MnuReporte.IsEnabled = False
        End If
    End Sub
    Protected Overrides Sub SGuarde()
        MyBase.SGuarde()
        SRefresqueWin()
    End Sub
    Protected Overrides Sub SEstablezcaWinConsultando()
        MyBase.SEstablezcaWinConsultando()
        MdtbModulosSer = Nothing
    End Sub
    Protected Overrides Sub SFinaliceOperacion()
        MyBase.SFinaliceOperacion()
        SHabiliteReporte()
        If MnuReporte.Visibility = Visibility.Visible Then
            MnuReporte.IsEnabled = True
        End If
    End Sub
    Protected Overrides Sub SRefresqueWin()
        MyBase.SRefresqueWin()
        tbiGenerales.IsSelected = True
        SHabiliteMenues()
        SHabiliteReporte()
    End Sub
    Private Sub SHabiliteReporte()
        If MobjObjetoWin.ObjGeneraProgramBln.ObjValorPro OrElse
                (MobjObjetoWin.ObjEsFactProgramableBln.ObjValorPro AndAlso
                MobjObjetoWin.BlnEsImportado) Then
            MnuReporte.Visibility = Visibility.Visible
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                MnuReporte.IsEnabled = True
            End If
        Else
            MnuReporte.Visibility = Visibility.Collapsed
        End If
    End Sub
#End Region

#Region "Mostrar"
    Private Sub SMuestreGenerales()
        With MobjObjetoWin
            txtConceptoSer.Text = .ObjConceptoServicioStr.ObjValorPro
            txtTipoServicio.Content = ClsOrionCop.FstrNombreDatoConstanteOri(
                    EnuGrupoConstantesOriDef.EnuTipoServicio,
                    MobjObjetoWin.ObjIdTipoServicioByt.ObjValorPro)
            txtDiaFactura.Text = .ObjDiaFacturaShr.ObjValorPro
            txtDiasVence.Text = .ObjDiasVencimientoShr.ObjValorPro
            chkVenceFinMes.IsChecked = .ObjVenceFinMesBln.ObjValorPro
            txtDiasGracia.Text = .ObjDiasGraciaShr.ObjValorPro
            chkGraciaFinMes.IsChecked = .ObjGraciaFinMesBln.ObjValorPro
            txtFechaFactura.Content = Format(MobjObjetoWin.DtmFechaFacturacionPeriodoActual,
                    GCSTRFMTFECHASIMPLE)
            txtFechaVence.Content = Format(MobjObjetoWin.DtmFechaVencePeriActual,
                    GCSTRFMTFECHASIMPLE)
            txtFechaGracia.Content = Format(MobjObjetoWin.DtmFechaGraciaPeriActual,
                    GCSTRFMTFECHASIMPLE)
            cboModoCausaMora.SelectedIndex = MobjObjetoWin.ObjModoCausaInteresesByt.ObjValorPro
            chkFactuPropYPreAgr.IsChecked = .ObjFactAPropYPreAgrBln.ObjValorPro
            chkEsServicioId.IsChecked = .ObjEsServicioIdBln.ObjValorPro
            chkFacturacionProgramable.IsChecked = .ObjEsFactProgramableBln.ObjValorPro
            chkEsCalculado.IsChecked = .ObjGeneraProgramBln.ObjValorPro
            cboTipoBaseCalculo.SelectedIndex = .ObjTipoBaseCalculoByt.ObjValorPro
            chkEstaCalculado.IsChecked = .ObjEstaGenaradaProgramBln.ObjValorPro
            If .ObjEsAjusteBln.ObjValorPro Then
                chkEstaAjustadoEsAjuste.Content = My.Resources.EsAjuste
            End If
            chkEstaAjustadoEsAjuste.IsChecked = .ObjEstaAjustadoBln.ObjValorPro OrElse
                .ObjEsAjusteBln.ObjValorPro
        End With
    End Sub
    Private Sub SMuestreValores()
        Dim lblnMostrarModulos = False
        If MobjObjetoWin.ObjGeneraProgramBln.ObjValorPro Then
            txtValorPres.Content = Format(MobjObjetoWin.DecValor, "c")
            lblnMostrarModulos = True
            SEstablezcaModuloSel()
        ElseIf MobjObjetoWin.ObjEsFactProgramableBln.ObjValorPro AndAlso
                MobjObjetoWin.BlnEsImportado Then
            lblnMostrarModulos = True
            txtValorPres.Content = Format(ClsOrionCop.FdecValorTotalCalculadoServicio(
                        MobjObjetoWin.ObjIdAno_ServicioShr.ObjValorPro,
                        MobjObjetoWin.ObjIdServicioShr.ObjValorPro), "c")
        End If
        If lblnMostrarModulos Then
            SActualiceTblModSer()
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                MdtbModulosSer = MobjObjetoWin.FdtbModulosServicio
            End If
            dgrContribuciones.DataContext = MdtbModulosSer
            dgrContribuciones.SelectedIndex = 0
        End If
        SMuestrePeriodo()
        txtCantidadPer.Text = MobjObjetoWin.ObjCantPeriodos_ServicioShr.ToString
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If MobjObjetoWin IsNot Nothing Then
                If MobjObjetoWin.ObjEsAjusteBln.ObjValorPro Then
                    txtValorCalculado.Content = Format(MobjObjetoWin.DecValorAjuste, "c")
                Else
                    If MobjObjetoWin.ObjGeneraProgramBln.ObjValorPro OrElse
                                MobjObjetoWin.BlnEsImportado Then
                        txtValorCalculado.Content =
                                    Format(ClsOrionCop.FdecValorTotalCalculadoServicio(
                                    MobjObjetoWin.ObjIdAno_ServicioShr.ObjValorPro,
                                    MobjObjetoWin.ObjIdServicioShr.ObjValorPro), "c")
                    Else
                        txtValorCalculado.Content = Format(0, "c")
                    End If
                End If
            End If
        Else
            txtValorCalculado.Content = Format(0, "c")
        End If
        SVisibiliceControles()
    End Sub
    Private Sub SMuestrePeriodo()
        Dim lstrPeriodo As String
        Dim lstrIdAno As String
        Dim lstrMes As String
        If cboAnoPeriodo.Items.Count = 0 OrElse cboMesPeriodo.Items.Count = 0 Then
            SPuebleCombosPeriodo()
        End If
        If MobjObjetoWin.ObjGeneraProgramBln.ObjValorPro OrElse MobjObjetoWin.BlnEsImportado Then
            If MobjObjetoWin.ObjPeriodoInicioStr.BlnEsValido Then
                MblnPoblandoComboBox = True
                lstrPeriodo = MobjObjetoWin.ObjPeriodoInicioStr.ObjValorPro
                lstrIdAno = lstrPeriodo.Substring(0, 4)
                lstrMes = Right(lstrPeriodo, 2)
                cboAnoPeriodo.SelectedItem = lstrIdAno
                cboMesPeriodo.SelectedItem = lstrMes
                MblnPoblandoComboBox = False
            End If
        End If
    End Sub
    Private Sub SMuestreParContables()
        With MobjObjetoWin
            txtCtaDebito.Text = .ObjCodigoCuentaDbStr.ObjValorPro
            txtNomCtaDebito.Content = .ObjCodigoCuentaDbStr.StrNombreCuenta
            txtNomCtaDebito.ToolTip = .ObjCodigoCuentaDbStr.StrNombreCuenta
            txtCtaCredito.Text = .ObjCodigoCuentaCrStr.ObjValorPro
            txtNomCtaCredito.Content = .ObjCodigoCuentaCrStr.StrNombreCuenta
            txtNomCtaCredito.ToolTip = .ObjCodigoCuentaCrStr.StrNombreCuenta
            txtCtaDevolucion.Text = .ObjCodigoCuentaDevStr.ObjValorPro
            txtNomCtaDevolucion.Content = .ObjCodigoCuentaDevStr.StrNombreCuenta
            txtNomCtaDevolucion.ToolTip = .ObjCodigoCuentaDevStr.StrNombreCuenta
            cboTipoTerCtaCr.SelectedIndex = .ObjIdTipoTerCtaCrSerByt.ObjValorPro
            txtIdTerceroCr.Text = .ObjIdTerceroCtaCrDbl.ObjValorPro
            txtNomTerceroCtaCe.Content = .ObjIdTerceroCtaCrDbl.StrNombreTercero
            txtNomTerceroCtaCe.ToolTip = txtNomTerceroCtaCe.Content
            txtCtaMora.Text = .ObjCodigoCuentaMoraStr.ObjValorPro
            txtNomCtaMora.Content = .ObjCodigoCuentaMoraStr.StrNombreCuenta
            txtCtaMora.ToolTip = .ObjCodigoCuentaMoraStr.StrNombreCuenta
            txtCtaIva.Text = .ObjCodigoCuentaIvaStr.ObjValorPro
            txtNomCtaIva.Content = .ObjCodigoCuentaIvaStr.StrNombreCuenta
            txtNomCtaIva.ToolTip = .ObjCodigoCuentaIvaStr.StrNombreCuenta
            txtTarifaIva.Text = Format(.ObjTarifaIvaDbl.ObjValorPro, "p")
            chkEsExcluido.IsChecked = .ObjEsExcluidoIvaBln.ObjValorPro
            txtTarifaretefuente.Text = Format(.ObjTarifaRetFteDbl.ObjValorPro, "p")
            txtBaseretefuente.Text = Format(.ObjBaseMinimaReteFuenteDec.ObjValorPro, "c")
            txtTarifaReteIca.Text = Format(.ObjTarifaRetIcaDbl.ObjValorPro, "p")
            txtBasereteica.Text = Format(.ObjBaseMinimaReteIcaDec.ObjValorPro, "c")
        End With
    End Sub
#End Region

#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        Me.Cursor = Cursors.Wait
        If HwinBusqueda Is Nothing Then
            HwinBusqueda = New WinBusqueda With {
                .WinPadre = Me
            }
        End If
        If FblnDefinioBusqueda() Then
            HwinBusqueda.ShowDialog()
        End If
        HwinBusqueda = Nothing
        Me.Cursor = Cursors.Arrow
    End Sub
    Private Sub SBusqueTer()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            StrResultadoBusqueda = String.Empty
            Me.Cursor = Cursors.Wait
            If HwinBusqueda Is Nothing Then
                HwinBusqueda = New WinBusqueda With {
                    .WinPadre = Me
                }
            End If
            If FblnDefinioBusqueda() Then
                HwinBusqueda.ShowDialog()
            End If
            HwinBusqueda = Nothing
            Me.Cursor = Cursors.Arrow
            If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                txtIdTerceroCr.Text = StrResultadoBusqueda
                MobjObjetoWin.ObjIdTerceroCtaCrDbl.ObjValorPro = StrResultadoBusqueda
                If MobjObjetoWin.ObjIdTerceroCtaCrDbl.BlnExisteTercero Then
                    txtNomTerceroCtaCe.Content = MobjObjetoWin.ObjIdTerceroCtaCrDbl.StrNombreTercero
                End If
            End If
        End If
    End Sub
    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        If bttEncontrarTerCr.IsFocused Then
            SDefineBusquedaApe()
            SDefineBusquedaNom()
            SDefineBusquedaNomRS()
        Else
            SDefineCuentaCont()
        End If
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
    Private Sub SDefineBusquedaNom()
        Dim lstrCamposMostrar As String() = {ClsIdTerceroDbl.SstrNombreCampoBd,
                ClsNombrePrimeroStr.SstrNombreCampoBd, ClsNombreSegundoStr.SstrNombreCampoBd,
                ClsApellidoPrimeroStr.SstrNombreCampoBd, ClsApellidoSegundoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombrePrimeroStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdTerceroDbl.SstrNombreCampoBd
        Dim lstrTabla As String = ClsTercero.SstrNombreTabla
        Dim lstrFiltro As String = lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Primer Nombre", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
    Private Sub SDefineBusquedaApe()
        Dim lstrCamposMostrar As String() = {ClsIdTerceroDbl.SstrNombreCampoBd,
                ClsNombrePrimeroStr.SstrNombreCampoBd, ClsNombreSegundoStr.SstrNombreCampoBd,
                ClsApellidoPrimeroStr.SstrNombreCampoBd, ClsApellidoSegundoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsApellidoPrimeroStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdTerceroDbl.SstrNombreCampoBd
        Dim lstrTabla As String = ClsTercero.SstrNombreTabla
        Dim lstrFiltro As String = lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Primer Apellido", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
    Private Sub SDefineBusquedaNomRS()
        Dim lstrCamposMostrar As String() = {ClsIdTerceroDbl.SstrNombreCampoBd,
                ClsRazonSocialStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsRazonSocialStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdTerceroDbl.SstrNombreCampoBd
        Dim lstrTabla As String = ClsTercero.SstrNombreTabla
        Dim lstrFiltro As String = lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Razón Social", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SAdicioneCtrlsRestringidos()
        SAdicioneControlRestringido(tbcServicios)
        SAdicioneControlRestringido(bttSectores)
        SAdicioneControlRestringido(bttEncontrarCtaCr)
        SAdicioneControlRestringido(bttEncontrarCtaMora)
        SAdicioneControlRestringido(bttEncontrarCtaDev)
        SAdicioneControlRestringido(bttEncontrarCtaDb)
        SAdicioneControlRestringido(bttEncontrarCtaIva)
        SAdicioneControlRestringido(bttEncontrarTerCr)
        SAdicioneControlRestringido(chkEstaAjustadoEsAjuste)
        SAdicioneControlRestringido(chkEstaCalculado)
        SAdicioneControlRestringido(txtAdvertencia)
    End Sub

    Private Function FblnHabiliteModulos() As Boolean
        Dim lblnHabilite = True
        If MobjObjetoWin.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
            lblnHabilite = Not GobjParametros.FblnPerActEsDicPrimerAno
        End If
        Return lblnHabilite
    End Function

    Private Sub SHabiliteAcciones()
        If MobjObjetoWin.ObjEsAjusteBln.ObjValorPro Then
            SHabiliteMenuItem(False, MnuCalcular)
            SHabiliteMenuItem(False, MnuLimpiar)
            SHabiliteMenuItem(False, HmnuModificar)
        Else
            If MobjObjetoWin.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
                If MobjObjetoWin.ObjMiAno.ObjModuloPorServicioBln.ObjValorPro Then
                    SHabiliteMenuItem(False, MnuCalcular)
                End If
            ElseIf Not MobjObjetoWin.ObjGeneraProgramBln.ObjValorPro Then
                SHabiliteMenuItem(False, MnuCalcular)
                SHabiliteMenuItem(False, MnuLimpiar)
            End If
        End If
    End Sub

    Private Sub SVisibiliceControles()
        If MobjObjetoWin IsNot Nothing Then
            If MobjObjetoWin.ObjEsFactProgramableBln.ObjValorPro Then
                If (MobjObjetoWin.ObjGeneraProgramBln.ObjValorPro AndAlso
                        FblnHabiliteModulos()) OrElse MobjObjetoWin.BlnEsImportado Then
                    tbiValores.Visibility = Visibility.Visible
                    grbEstado.Visibility = Visibility.Visible
                    bttSectores.Visibility = Visibility.Visible
                    bttSectores.IsEnabled = EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando
                    lblTipoBaseCalculo.Visibility = Visibility.Visible
                    cnvBotonesCont.Visibility = Visibility.Visible
                    dgrContribuciones.Visibility = Visibility.Visible
                    cboTipoBaseCalculo.Visibility = Visibility.Visible
                    MsepMenu.Visibility = Visibility.Visible
                    If MobjObjetoWin.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual OrElse
                            MobjObjetoWin.BlnEsImportado Then
                        MnuLimpiar.Visibility = Visibility.Collapsed
                    Else
                        MnuLimpiar.Visibility = Visibility.Visible
                    End If
                    If Not MobjObjetoWin.BlnEsImportado Then
                        MnuCalcular.Visibility = Visibility.Visible
                    End If
                    If MobjObjetoWin.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual Then
                        lblAdvertencia.Visibility = Visibility.Visible
                        txtAdvertencia.Visibility = Visibility.Visible
                        chkEstaAjustadoEsAjuste.Visibility = Visibility.Visible
                        chkEsServicioId.Visibility = Visibility.Hidden
                    Else
                        MsepMenu.Visibility = Visibility.Visible
                        lblAdvertencia.Visibility = Visibility.Hidden
                        txtAdvertencia.Visibility = Visibility.Hidden
                    End If
                    lblValorCalculado.Visibility = Visibility.Visible
                    txtValorCalculado.Visibility = Visibility.Visible
                Else
                    tbiValores.Visibility = Visibility.Collapsed
                    MnuCalcular.Visibility = Visibility.Collapsed
                    MnuLimpiar.Visibility = Visibility.Collapsed
                    grbEstado.Visibility = Visibility.Collapsed
                    bttSectores.Visibility = Visibility.Collapsed
                    lblTipoBaseCalculo.Visibility = Visibility.Collapsed
                    cboTipoBaseCalculo.Visibility = Visibility.Collapsed
                End If
                chkEsServicioId.Visibility = Visibility.Collapsed
            Else
                tbiValores.Visibility = Visibility.Collapsed
                MnuCalcular.Visibility = Visibility.Collapsed
                MnuLimpiar.Visibility = Visibility.Collapsed
                grbEstado.Visibility = Visibility.Collapsed
                bttSectores.Visibility = Visibility.Collapsed
                lblTipoBaseCalculo.Visibility = Visibility.Collapsed
                cboTipoBaseCalculo.Visibility = Visibility.Collapsed
                If MobjObjetoWin.ObjEsServicioIdBln.ObjValorPro Then
                    chkEsCalculado.Visibility = Visibility.Collapsed
                    chkFacturacionProgramable.Visibility = Visibility.Collapsed
                Else
                    chkEsCalculado.Visibility = Visibility.Visible
                End If
                chkEsServicioId.Visibility = Visibility.Visible
            End If
            SHabiliteCtrls()
        End If
    End Sub

    Private Sub SHabiliteCtrls()
        If MobjObjetoWin.BlnEsCuotaAdministracion Then
            chkFacturacionProgramable.Style = FindResource("RecCtlNoHabilitado")
            chkEsCalculado.Style = FindResource("RecCtlNoHabilitado")
        End If
        If ClsOrionCop.FblnEstaServicioActivo(MobjObjetoWin) Then
            chkEsCalculado.IsEnabled = False
            chkFacturacionProgramable.IsEnabled = False
        Else
            If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando AndAlso
                    MobjObjetoWin.ObjIdTipoServicioByt.ObjValorPro =
                    EnuTipoServicio.EnuPermanente Then
                chkEsCalculado.IsEnabled = True
                chkFacturacionProgramable.IsEnabled = True
            Else
                chkEsCalculado.IsEnabled = False
                chkFacturacionProgramable.IsEnabled = False
            End If
        End If
        SHabiliteCtrlsValores()
    End Sub

    Private Sub SHabiliteCtrlsValores()
        If tbiValores.Visibility = Visibility.Visible Then
            dgrContribuciones.IsEnabled = True
            If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                If MobjObjetoWin.ObjGeneraProgramBln.ObjValorPro Then
                    cboTipoBaseCalculo.Style = FindResource("RecCtlHabilitado")
                End If
                If MobjObjetoWin.BlnEsCuotaAdministracion Then
                    bttAgregarContribucion.IsEnabled = True
                    bttModificarContribucion.IsEnabled = dgrContribuciones.Items.Count > 0
                    bttEliminarContribucion.IsEnabled = dgrContribuciones.Items.Count > 0
                    cboMesPeriodo.Style = FindResource("RecCtlNoHabilitado")
                    cboAnoPeriodo.Style = FindResource("RecCtlNoHabilitado")
                    txtCantidadPer.Style = FindResource("RecCtlNoHabilitado")
                Else
                    If MobjObjetoWin.ObjGeneraProgramBln.ObjValorPro Then
                        cboMesPeriodo.Style = FindResource("RecCtlHabilitado")
                        cboAnoPeriodo.Style = FindResource("RecCtlHabilitado")
                        txtCantidadPer.Style = FindResource("RecCtlHabilitado")
                    End If
                    If MobjObjetoWin.ObjTipoBaseCalculoByt.ObjValorPro =
                            EnuTipoBaseCalculo.EnuCoeficientePro OrElse
                            MobjObjetoWin.ObjTipoBaseCalculoByt.ObjValorPro =
                            EnuTipoBaseCalculo.EnuUnidad Then
                        bttAgregarContribucion.IsEnabled = True
                        bttModificarContribucion.IsEnabled = dgrContribuciones.Items.Count > 0
                        bttEliminarContribucion.IsEnabled = dgrContribuciones.Items.Count > 0
                    End If
                End If
                If MobjObjetoWin.BlnEsImportado Then
                    bttAgregarContribucion.IsEnabled = False
                    bttModificarContribucion.IsEnabled = False
                    bttEliminarContribucion.IsEnabled = False
                End If
                If MobjModuloSel IsNot Nothing AndAlso MobjModuloSel.EnuEstadoActualizacion <>
                        EnuEstadoObjetoDef.enuCreando Then
                    SEstablezcaModuloSel()
                End If
            Else
                If tbiValores.Visibility = Visibility.Visible Then
                    bttAgregarContribucion.IsEnabled = False
                    bttModificarContribucion.IsEnabled = False
                    bttEliminarContribucion.IsEnabled = False
                End If
            End If
        End If
    End Sub

    Private Sub SVisibiliceBotonesEncontrar()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            bttEncontrarCtaCr.Visibility = Visibility.Visible
            bttEncontrarCtaDb.Visibility = Visibility.Visible
            bttEncontrarCtaIva.Visibility = Visibility.Visible
            bttEncontrarCtaDev.Visibility = Visibility.Visible
            bttEncontrarTerCr.Visibility = Visibility.Visible
            bttEncontrarCtaMora.Visibility = Visibility.Visible
        Else
            bttEncontrarCtaCr.Visibility = Visibility.Collapsed
            bttEncontrarCtaDb.Visibility = Visibility.Collapsed
            bttEncontrarCtaDev.Visibility = Visibility.Collapsed
            bttEncontrarCtaIva.Visibility = Visibility.Collapsed
            bttEncontrarTerCr.Visibility = Visibility.Collapsed
            bttEncontrarCtaMora.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub SPuebleComboBoxes()
        MblnPoblandoComboBox = True
        Dim ldrwDataRow = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuTipoBaseCalculo)
        SPuebleComboBox(ldrwDataRow, cboTipoBaseCalculo)
        ldrwDataRow = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuTipoTerCtaCrSer)
        SPuebleComboBox(ldrwDataRow, cboTipoTerCtaCr)
        ldrwDataRow = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuModoCausaMora)
        SPuebleComboBox(ldrwDataRow, cboModoCausaMora)
        SPuebleCombosPeriodo()
        MblnPoblandoComboBox = False
    End Sub

    Private Sub SPuebleCombosPeriodo()
        Dim lshrIdAnoActual = GobjParametros.ObjAnoActual.ObjIdAnoShr.ObjValorPro
        MblnPoblandoComboBox = True
        cboAnoPeriodo.Items.Clear()
        cboMesPeriodo.Items.Clear()
        If MobjObjetoWin IsNot Nothing AndAlso
                (MobjObjetoWin.ObjGeneraProgramBln.ObjValorPro OrElse
                (MobjObjetoWin.ObjEsFactProgramableBln.ObjValorPro AndAlso
                MobjObjetoWin.BlnEsImportado)) Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                lshrIdAnoActual = CShort(MobjObjetoWin.ObjPeriodoInicioStr.ToString.Substring(0, 4))
            End If
            For i As Integer = 1 To 12
                cboMesPeriodo.Items.Add(Right("00" & i.ToString, 2))
            Next
            For i As Integer = -1 To 4
                cboAnoPeriodo.Items.Add((lshrIdAnoActual + i).ToString)
            Next
            If GobjParametros.FblnPerActEsDicPrimerAno Then
                cboAnoPeriodo.Items.Clear()
                cboAnoPeriodo.Items.Add("0000")
                cboMesPeriodo.Items.Clear()
                cboMesPeriodo.Items.Add("00")
            End If
        End If
        cboAnoPeriodo.SelectedIndex = 0
        cboMesPeriodo.SelectedIndex = 0
        MblnPoblandoComboBox = False
    End Sub

    Private Sub SEstablezcaToolTipGral()
        bttSectores.ToolTip = My.Resources.TTAbrirSectoresServicio
    End Sub

    Private Sub SAbraServicio()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                Dim lshrIdAno As Short
                If txtIdServicio.Text <> MobjObjetoWin.ObjIdServicioShr.ToString() Then
                    If MobjAno IsNot Nothing Then
                        ' Servicio Anual
                        lshrIdAno = MobjAno.ObjIdAnoShr.ObjValorPro
                    Else
                        ' Servicio Permanente
                        lshrIdAno = 0
                    End If
                    Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lshrIdAno,
                        txtIdServicio.Text}
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
                If Not lblnNoHayError Then
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If

    End Sub

    Private Sub SAbraSectores()
        Dim lwinSectoresServicio As New WinSectoresServicio(MobjObjetoWin)
        lwinSectoresServicio.Show()
    End Sub

    Private Sub SAbraHistorico()
        MobjObjetoWin.SRefresqueObj()
        Dim lwinHistorico As New WinHistServicio(MobjObjetoWin)
        lwinHistorico.Show()
    End Sub

    Private Sub SCalculeValoresACobrar()
        Dim lblnNoHayError = False, lblnCalcular As Boolean, lblnCalculo As Boolean
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty
        Dim lenuSevNot As EnuSeveridadNot = EnuSeveridadNot.None
        Try
            Mouse.OverrideCursor = Cursors.Wait
            GobjPanDat.SInicialiceTransaccion()
            GobjPanDat.SControleProcesoObj(True)
            GobjParametros.SVerifiqueApp(False, True)
            lblnCalcular = (GobjParametros.EnuEstadoInstalacion = EnuEstadoInstalacion.Todos)
            If Not lblnCalcular Then
                lstrMens = "No se puede hacer calculos antes de terminar la Parametrización!"
                lenuSevNot = EnuSeveridadNot.EnuInformacion
            End If
            If lblnCalcular Then
                lblnCalcular = Not MobjObjetoWin.BlnEsCuotaAdministracion
                If Not lblnCalcular Then
                    lstrMens = "Cuando el Cálculo está relacionado con la Cuota de Administración, " &
                        "éste debe hacerse desde la Ventana del Año actual!"
                    lenuSevNot = EnuSeveridadNot.EnuInformacion
                End If
            End If
            If lblnCalcular Then
                If MobjObjetoWin.DecValor = 0 Then
                    lblnCalcular = MsgBox("Realmente quiere eliminar la programación de este servicio?",
                        vbYesNo, "Eliminar programación")
                    If lblnCalcular Then
                        lstrMens = "Se eliminará la programación de facturas!"
                        lenuSevNot = EnuSeveridadNot.EnuInformacion
                        SLevanteEveNoti(lstrMens, String.Empty, 0, lenuSevNot)
                        lstrMens = String.Empty
                    End If
                End If
            End If
            If lblnCalcular Then
                lblnCalculo = ClsCalculosServicios.FblnCalculoItemsServicio(MobjObjetoWin)
            End If
            lblnNoHayError = True
        Catch ex As ProveedorBdPanException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
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
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            SRefrescarClic()
            If String.IsNullOrEmpty(lstrMens) Then
                lstrMens = "Proceso finalizado exitosamente!"
                lenuSevNot = EnuSeveridadNot.EnuInformacion
            End If
            SLevanteEveNoti(lstrMens, String.Empty, 0, lenuSevNot)
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub

    Private Sub SlimpieValoresACobrar()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty
        Dim lblnNoHayError = False, lblnLimpiar As Boolean
        Dim lenuSevNot As EnuSeveridadNot = EnuSeveridadNot.None
        Try
            Mouse.OverrideCursor = Cursors.Wait
            GobjPanDat.SInicialiceTransaccion()
            GobjPanDat.SControleProcesoObj(True)
            GobjParametros.SVerifiqueApp(False, True)
            lblnLimpiar = GobjParametros.EnuEstadoInstalacion = EnuEstadoInstalacion.Todos
            If Not lblnLimpiar Then
                lstrMens = "No se puede hacer calculos antes de terminar la Parametrización!"
                lenuSevNot = EnuSeveridadNot.EnuInformacion
            End If
            If lblnLimpiar Then
                lblnLimpiar = Not MobjObjetoWin.BlnEsCuotaAdministracion
                If Not lblnLimpiar Then
                    lstrMens = "Esta operación solo es ejecutable en servicios permanentes!"
                    lenuSevNot = EnuSeveridadNot.EnuInformacion
                End If
            End If
            If lblnLimpiar Then
                If MobjObjetoWin.DecValor > 0 AndAlso
                        MobjObjetoWin.ObjEstaGenaradaProgramBln.ObjValorPro Then
                    lblnLimpiar = MsgBox("Realmente quiere eliminar la programación de este servicio?",
                            vbYesNo, "Eliminar programación")
                    If lblnLimpiar Then
                        lstrMens = "Se eliminará la programación de facturas!"
                        lenuSevNot = EnuSeveridadNot.EnuInformacion
                        SLevanteEveNoti(lstrMens, String.Empty, 0, lenuSevNot)
                        lstrMens = String.Empty
                        ClsCalculosServicios.SLimpieItemsServicio(MobjObjetoWin)
                        GobjPanorama.SRegistreAccionLogApp("ClsCalculosServicios",
                                "Limpieza ítems programa facturación del servicio permanente " &
                                MobjObjetoWin.ObjIdServicioShr.ObjValorPro)
                    End If
                End If
            End If
            lblnNoHayError = True
        Catch ex As ProveedorBdPanException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
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
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            SRefrescarClic()
            If String.IsNullOrEmpty(lstrMens) Then
                lstrMens = "Proceso finalizado exitosamente!"
                lenuSevNot = EnuSeveridadNot.EnuInformacion
            End If
            SLevanteEveNoti(lstrMens, String.Empty, 0, lenuSevNot)
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub

    Private Sub SVerifiqueEstado(ByRef astrMens As String)
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            GobjPanDat.SControleProcesoObj(True)
            If MobjObjetoWin.ObjGeneraProgramBln.ObjValorPro Then
                If Not (MobjObjetoWin.ObjIdTipoServicioByt.ObjValorPro =
                        EnuTipoServicio.EnuAnual AndAlso
                        GobjParametros.FblnPerActEsDicPrimerAno) AndAlso
                        Not MobjObjetoWin.ObjEsAjusteBln.ObjValorPro Then
                    If MobjObjetoWin.ColModulosServicio.Count = 0 Then
                        astrMens = "Se deben asignar los Módulos que contribuyen al Servicio!"
                    ElseIf MobjObjetoWin.BlnEsImportado Then
                        astrMens = "Los Valores del Servicio son Importados!"
                    Else
                        If MobjObjetoWin.ObjIdTipoServicioByt.ObjValorPro =
                                EnuTipoServicio.EnuAnual Then
                            If Not MobjObjetoWin.FblnVlrsModsIngresados Then
                                If MobjObjetoWin.ObjMiAno.ObjModuloPorServicioBln.ObjValorPro Then
                                    astrMens = "Se debe ingresar el Valor del Presupuesto " &
                                        "anual en el Año!"
                                Else
                                    If MobjObjetoWin.ObjMiAno.ObjTipoCalculoCuotaByt.ObjValorPro =
                                            EnuTipoBaseCalculo.EnuCoeficientePro OrElse
                                            MobjObjetoWin.ObjMiAno.ObjTipoCalculoCuotaByt.
                                            ObjValorPro = EnuTipoBaseCalculo.EnuUnidad Then
                                        astrMens = "Se deben ingresar los Valores de los Módulos " &
                                        "de Contribución!"
                                    End If
                                End If
                            ElseIf Not MobjObjetoWin.ObjEstaGenaradaProgramBln.ObjValorPro Then
                                If GobjParametros.EnuEstadoInstalacion =
                                        EnuEstadoInstalacion.Todos Then
                                    If MobjAno IsNot Nothing AndAlso
                                            MobjAno.ObjModuloPorServicioBln.ObjValorPro Then
                                        astrMens = "Se deben calcular los Valores a cobrar desde " &
                                                "la ventana del año!"
                                    Else
                                        astrMens = "Se deben calcular los Valores a cobrar!"
                                    End If
                                End If
                            End If
                        Else
                            If Not MobjObjetoWin.ObjEstaGenaradaProgramBln.ObjValorPro Then
                                astrMens = "Se deben calcular los Valores a cobrar!"
                            ElseIf Not MobjObjetoWin.FblnVlrsModsIngresados Then
                                astrMens = "Se deben ingresar los Valores de los Módulos " &
                                        "de Contribución!"
                            End If
                        End If
                    End If
                End If
            End If
            GobjPanDat.SControleProcesoObj(False)
        End If
    End Sub

    Private Sub SVerifiqueSectoresModulo(ByRef astrMens As String)
        If MobjObjetoWin IsNot Nothing AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.CenuConsultando Then
            MblnSectoresAsignados = MobjObjetoWin.FblnSectoresAsignadosaModulos
            If Not MblnSectoresAsignados Then
                If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                    SHabiliteMenuItem(False, MnuCalcular)
                    SHabiliteMenuItem(False, MnuLimpiar)
                    astrMens = "No todos los Módulos de Contribución tienen Sectores asignados"
                End If
            End If
        End If
    End Sub

    Private Sub SRegistreCtaContable(actrlControl As Control)
        Dim lstrMens = String.Empty
        Dim ctrlTextBox As TextBox = TryCast(actrlControl, TextBox)
        Dim lstrIdCtaContable As String = ctrlTextBox.Text
        Dim lstrNomServicio As String = MobjObjetoWin.ObjNombreServicioStr.ToString()
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
                    Dim lstrNombreCuenta As String = ctrlTextBox.ToolTip & lstrNomServicio
                    lobjCtaCont.SCreeObj({GshrIdCarpeta, lstrIdCtaContable})
                    lobjCtaCont.ObjIdCarpetaCuentaShr.ObjValorPro = GshrIdCarpeta
                    lobjCtaCont.ObjIdCuentaContStr.ObjValorPro = lstrIdCtaContable
                    lobjCtaCont.ObjNombreCuentaStr.ObjValorPro = lstrNombreCuenta
                    lobjCtaCont.SActualice(True)
                End If
            End If
            With MobjObjetoWin
                Select Case actrlControl.Name
                    Case "txtCtaCredito"
                        .ObjCodigoCuentaCrStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaCredito.Content = lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
                    Case "txtCtaDebito"
                        .ObjCodigoCuentaDbStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaDebito.Content = lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
                    Case "txtCtaIva"
                        .ObjCodigoCuentaIvaStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaIva.Content = lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
                    Case "txtCtaDevolucion"
                        .ObjCodigoCuentaDevStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaDevolucion.Content = lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
                    Case "txtCtaMora"
                        .ObjCodigoCuentaMoraStr.ObjValorPro = lstrIdCtaContable
                        txtNomCtaMora.Content = lobjCtaCont.ObjNombreCuentaStr.ObjValorPro
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

    Private Function FblnRegistroEntrada(actrlControl As Control) As Boolean
        Dim lblnRegistro = True
        With MobjObjetoWin
            Select Case actrlControl.Name
                Case "txtNombreServicio"
                    .ObjNombreServicioStr.ObjValorPro = txtNombreServicio.Text
                Case "txtConceptoSer"
                    .ObjConceptoServicioStr.ObjValorPro = txtConceptoSer.Text
                Case "chkFacturacionProgramable"
                    .ObjEsFactProgramableBln.ObjValorPro = chkFacturacionProgramable.IsChecked
                Case "chkEsCalculado"
                    .ObjGeneraProgramBln.ObjValorPro = chkEsCalculado.IsChecked
                Case "cboTipoBaseCalculo"
                    .ObjTipoBaseCalculoByt.ObjValorPro = cboTipoBaseCalculo.SelectedIndex
                Case "txtCantidadPer"
                    .ObjCantPeriodos_ServicioShr.ObjValorPro = txtCantidadPer.Text
                Case "cboAnoPeriodo", "cboMesPeriodo"
                    Dim lstrPeriodo As String = cboAnoPeriodo.SelectedItem & cboMesPeriodo.SelectedItem
                    .ObjPeriodoInicioStr.ObjValorPro = lstrPeriodo
                Case "txtTarifaIva"
                    .ObjTarifaIvaDbl.ObjValorPro = FdblTasa(txtTarifaIva.Text)
                Case "chkEsExcluido"
                    .ObjEsExcluidoIvaBln.ObjValorPro = chkEsExcluido.IsChecked
                Case "txtTarifaretefuente"
                    .ObjTarifaRetFteDbl.ObjValorPro = FdblTasa(txtTarifaretefuente.Text)
                Case "txtTarifaReteIca"
                    .ObjTarifaRetIcaDbl.ObjValorPro = FdblTasa(txtTarifaReteIca.Text)
                Case Else
                    lblnRegistro = False
            End Select
        End With
        Return lblnRegistro
    End Function

    Private Sub SCrearTercero(astrIdTercero As String)
        Dim lstrMens = String.Empty
        If ClsPanorama.FblnEsValidoNumero(astrIdTercero, 1, Double.MaxValue, True,
                        EnuTipoValor.enuDouble) Then
            If MsgBox("El Tercero ingresado no existe. Desea crearlo ahora?",
                      MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Crear Tercero ?") = MsgBoxResult.Yes Then
                Dim lobjTercero As New ClsTercero(EnuModoInstanciaObjDef.enuNavegable)
                Dim ldblIdTercero = CType(astrIdTercero, Double)
                lobjTercero.SCreeObj({ldblIdTercero})
                lobjTercero.ObjIdTerceroDbl.ObjValorPro = ldblIdTercero
                Dim lobjOrionCop = New ClsOrionCop(GCOBJREGISTRO, False)
                Dim lwinVentana As New WinTerceros() With {
                    .ObjObjetoWin = lobjTercero,
                    .EnuOperacionEnWin = EnuOperacionEnWin.CenuCreando,
                    .WinPadre = Me
                }
                lwinVentana.ShowDialog()
                MobjObjetoWin.ObjIdTerceroCtaCrDbl.ObjValorPro = txtIdTerceroCr.Text
            End If
        Else
            lstrMens = "El Valor ingresado no es válido!"
        End If
        SMuestreDatos()
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Private Sub SGenereReporte()
        Mouse.OverrideCursor = Cursors.Wait
        Dim lstrMens As String
        Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
            .ShrIdAno = MobjObjetoWin.ObjIdAno_ServicioShr.ObjValorPro,
            .EntIdServicio = MobjObjetoWin.ObjIdServicioShr.ObjValorPro,
            .EnuReporte = EnuReporteDef.enuItemsProgramaFact
            }
        lstrMens = lobjRep.SGenereItemsProgramaFact()
        Mouse.OverrideCursor = Cursors.Arrow
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        Else
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuOk)
        End If
    End Sub

    Protected Overrides Function FblnNotificaOk(aenuIdMens As EnuIdMens) As Boolean
        Dim lblnOk = False
        If aenuIdMens = EnuIdMens.EnuModSinSector Then
            lblnOk = MobjObjetoWin.FblnSectoresAsignadosaModulos
        End If
        Return lblnOk
    End Function
#End Region

#Region "Modulos del servicio"
    Private Sub SEstablezcaModuloSel()
        Dim ldrvFilaActual As DataRowView
        ldrvFilaActual = dgrContribuciones.SelectedItem
        MshrIdModuloSel = 0
        If ldrvFilaActual IsNot Nothing Then
            MshrIdModuloSel = ldrvFilaActual("IdModuloContribucion")
            If MshrIdModuloSel > 0 Then
                If MobjObjetoWin.ColModulosServicio.Count > 0 Then
                    If MobjObjetoWin.ColModulosServicio.Contains(MshrIdModuloSel.ToString) Then
                        MobjModuloSel = MobjObjetoWin.ColModulosServicio(MshrIdModuloSel.ToString())
                    Else
                        MobjModuloSel = MobjObjetoWin.ColModulosServicio(1)
                    End If
                Else
                    MobjModuloSel = Nothing
                End If
            End If
        Else
            MshrIdModuloSel = 0
        End If
    End Sub
    Friend Sub SAcepteModuloSer(ablnVinculando As Boolean, ByRef astrMens As String)
        If ablnVinculando Then
            Dim lblnVincule As Boolean
            Dim lshrIdNewMod As Short = MobjModuloSel.ObjIdModulo_ModuloServicioShr.ObjValorPro
            lblnVincule = Not MobjObjetoWin.FblnModuloMeContribuye(lshrIdNewMod)
            If lblnVincule Then
                If MobjObjetoWin.ObjIdTipoServicioByt.ObjValorPro = EnuTipoServicio.EnuAnual AndAlso
                        MobjAno.ObjModuloPorServicioBln.ObjValorPro Then
                    lblnVincule = Not MobjObjetoWin.ObjMiAno.FblnModuloYaContribuye(lshrIdNewMod)
                End If
                If Not lblnVincule Then
                    astrMens = "El módulo ya está contribuyendo con otro servicio del año!"
                End If
            Else
                astrMens = "El Módulo ya está contribuyendo al Servicio!"
            End If
            If lblnVincule Then
                MobjObjetoWin.SAdicioneNewModuloSer(MobjModuloSel)
                SLevanteEveOk()
                SMuestreValores()
                SValide()
                SHabiliteCtrlsValores()
            End If
            If Not String.IsNullOrEmpty(astrMens) Then
                SLevanteEveNoti(astrMens, "", 0, EnuSeveridadNot.EnuInformacion)
            End If
        Else
            SMuestreValores()
            SValide()
        End If
        Dim NoUsado = FblnModulosSerOk()
    End Sub
    Private Function FblnModulosSerOk()
        Dim lblnOk = True
        If MobjObjetoWin.ObjGeneraProgramBln.ObjValorPro Then
            lblnOk = MdtbModulosSer.Rows.Count > 0
            If Not lblnOk Then
                lblnOk = MobjObjetoWin.ObjEsAjusteBln.ObjValorPro
            End If
            If Not lblnOk Then
                lblnOk = GobjParametros.FblnPerActEsDicPrimerAno
            End If
            If lblnOk Then
                If (MobjObjetoWin.BlnEsCuotaAdministracion AndAlso
                            Not MobjObjetoWin.ObjMiAno.ObjModuloPorServicioBln.ObjValorPro) OrElse
                            (Not MobjObjetoWin.BlnEsCuotaAdministracion AndAlso MobjObjetoWin.
                            ObjGeneraProgramBln.ObjValorPro) Then
                    For Each lobjModSer As ClsModuloServicio In MobjObjetoWin.ColModulosServicio
                        lblnOk = lobjModSer.ObjValorPres_ModuloServicioDec.ObjValorPro > 0
                        If Not lblnOk Then
                            SLevanteEveNoti("Todos los módulos deben tener un valor!", "", 0,
                                    EnuSeveridadNot.EnuInformacion)
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        Return lblnOk
    End Function
    Private Sub SActualiceTblModSer()
        If MdtbModulosSer Is Nothing Then
            MdtbModulosSer = MobjObjetoWin.FdtbModulosServicio
        ElseIf EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            MdtbModulosSer.Rows.Clear()
            For Each lobjModuloSer As ClsModuloServicio In MobjObjetoWin.ColModulosServicio
                Dim ldrwNewFila As DataRow = MdtbModulosSer.NewRow
                ldrwNewFila(StrCampoCarpeta) = GshrIdCarpeta
                ldrwNewFila(StrCampoCentroutil) = GshrIdCentroUtil
                ldrwNewFila(ClsIdAno_ModuloServicioShr.SstrNombreCampoBd) =
                        lobjModuloSer.ObjIdAno_ModuloServicioShr.ObjValorPro
                ldrwNewFila(ClsIdServicio_ModuloServicioShr.SstrNombreCampoBd) =
                        lobjModuloSer.ObjIdServicio_ModuloServicioShr.ObjValorPro
                ldrwNewFila(ClsIdModulo_ModuloServicioShr.SstrNombreCampoBd) =
                        lobjModuloSer.ObjIdModulo_ModuloServicioShr.ObjValorPro
                ldrwNewFila(ClsValorPres_ModuloServicioDec.SstrNombreCampoBd) =
                        lobjModuloSer.ObjValorPres_ModuloServicioDec.ObjValorPro
                MdtbModulosSer.Rows.Add(ldrwNewFila)
            Next
            MobjObjetoWin.SRepuebleNombresModulos(MdtbModulosSer)
        End If
    End Sub
    Private Sub SAgregarCont()
        MobjModuloSel = MobjObjetoWin.FobjNewModuloSer()
        Dim lwinModuloSer As New WinModulosServicios(MobjModuloSel, True) With {
            .WinPadre = Me
        }
        lwinModuloSer.ShowDialog()
        If GblnOK Then
            SValide()
        End If
        SHabiliteCtrlsValores()
        SEstablezcaModuloSel()
    End Sub
    Private Sub SModificarCont()
        If MshrIdModuloSel > 0 Then
            SEstablezcaModuloSel()
            If MobjModuloSel.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
                MobjModuloSel.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando
            End If
            SLevanteEveOk()
            Dim lwinModuloSer As New WinModulosServicios(MobjModuloSel, False) With {
                .WinPadre = Me
            }
            lwinModuloSer.Show()
        End If
        SHabiliteCtrlsValores()
    End Sub
    Private Sub SElimineModCont()
        If MshrIdModuloSel > 0 Then
            MobjObjetoWin.SElimineModuloSer(MshrIdModuloSel)
            SMuestreValores()
            SValide()
            FblnModulosSerOk()
        End If
        SHabiliteCtrlsValores()
        SEstablezcaModuloSel()
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Dim lbttBoton As Button = lelmElemento
            Select Case lbttBoton.Name
                Case "bttSectores"
                    SAbraSectores()
                Case "bttEncontrarCtaCr", "bttEncontrarCtaDb", "bttEncontrarCtaIva",
                        "bttEncontrarCtaDev", "bttEncontrarCtaMora"
                    If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                        SBuscar()
                        If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                            Select Case lbttBoton.Name
                                Case "bttEncontrarCtaCr"
                                    txtCtaCredito.Text = StrResultadoBusqueda
                                    MobjObjetoWin.ObjCodigoCuentaCrStr.ObjValorPro = StrResultadoBusqueda
                                Case "bttEncontrarCtaDb"
                                    txtCtaDebito.Text = StrResultadoBusqueda
                                    MobjObjetoWin.ObjCodigoCuentaDbStr.ObjValorPro = StrResultadoBusqueda
                                Case "bttEncontrarCtaMora"
                                    txtCtaMora.Text = StrResultadoBusqueda
                                    MobjObjetoWin.ObjCodigoCuentaMoraStr.ObjValorPro = StrResultadoBusqueda
                                Case "bttEncontrarCtaDev"
                                    txtCtaDevolucion.Text = StrResultadoBusqueda
                                    MobjObjetoWin.ObjCodigoCuentaDevStr.ObjValorPro = StrResultadoBusqueda
                                Case "bttEncontrarCtaIva"
                                    txtCtaIva.Text = StrResultadoBusqueda
                                    MobjObjetoWin.ObjCodigoCuentaIvaStr.ObjValorPro = StrResultadoBusqueda
                            End Select
                        End If
                        SMuestreDatos()
                    End If
                Case "bttEncontrarTerCr"
                    SBusqueTer()
                Case "bttAgregarContribucion"
                    SAgregarCont()
                Case "bttModificarContribucion"
                    SModificarCont()
                Case "bttEliminarContribucion"
                    SElimineModCont()
            End Select
        End If
    End Sub

    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            Select Case lelmElemento.Name
                Case "MnuConsultarHistorico"
                    SAbraHistorico()
                Case "MnuCalcular"
                    SCalculeValoresACobrar()
                Case "MnuLimpiar"
                    SlimpieValoresACobrar()
                Case "MnuValorPredio"
                    Dim lstrMens = "Se esta generando el Reporte!"
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    SGenereReporte()
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
        If Not HblnSeEstaCerrando AndAlso Not HblnMostrandoDatos Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is CheckBox OrElse
                    TypeOf lelmElemento Is ComboBox Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                    Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                    Try
                        If lelmElemento.Name.ToString.StartsWith("txtCta") Then
                            SRegistreCtaContable(lelmElemento)
                        Else
                            If Not FblnRegistroEntrada(lelmElemento) Then
                                With MobjObjetoWin
                                    Select Case lelmElemento.Name
                                        Case "txtDiaFactura"
                                            .ObjDiaFacturaShr.ObjValorPro = txtDiaFactura.Text
                                        Case "txtDiasVence"
                                            .ObjDiasVencimientoShr.ObjValorPro = txtDiasVence.Text
                                        Case "chkVenceFinMes"
                                            .ObjVenceFinMesBln.ObjValorPro = chkVenceFinMes.IsChecked
                                        Case "chkActivo"
                                            .ObjEstaActivoServicioBln.ObjValorPro = chkActivo.IsChecked
                                        Case "txtDiasGracia"
                                            .ObjDiasGraciaShr.ObjValorPro = txtDiasGracia.Text
                                        Case "txtCtaDebito"
                                            .ObjCodigoCuentaDbStr.ObjValorPro = txtCtaDebito.Text
                                        Case "txtCtaCredito"
                                            .ObjCodigoCuentaCrStr.ObjValorPro = txtCtaCredito.Text
                                        Case "txtCtaDevolucion"
                                            .ObjCodigoCuentaDevStr.ObjValorPro = txtCtaDevolucion.Text
                                        Case "txtCtaMora"
                                            .ObjCodigoCuentaMoraStr.ObjValorPro = txtCtaMora.Text
                                        Case "txtCtaIva"
                                            .ObjCodigoCuentaIvaStr.ObjValorPro = txtCtaIva.Text
                                        Case "txtBaseretefuente"
                                            .ObjBaseMinimaReteFuenteDec.ObjValorPro = txtBaseretefuente.Text
                                        Case "txtBasereteica"
                                            .ObjBaseMinimaReteIcaDec.ObjValorPro = txtBasereteica.Text
                                        Case "txtIdTerceroCr"
                                            .ObjIdTerceroCtaCrDbl.ObjValorPro = txtIdTerceroCr.Text
                                            Dim lblnEsNumValido = ClsPanorama.FblnEsValidoNumero(
                                                    .ObjIdTerceroCtaCrDbl.ObjValorPro, GCDBLMINTERC, GCDBLMAXTERC,
                                                    True, EnuTipoValor.EnuDouble)
                                            If lblnEsNumValido Then
                                                If Not .ObjIdTerceroCtaCrDbl.BlnExisteTercero Then
                                                    SCrearTercero(txtIdTerceroCr.Text)
                                                End If
                                            End If
                                        Case "txtIdServicio"
                                            SAbraServicio()
                                    End Select
                                End With
                            End If
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
            End If
        End If
    End Sub

    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdServicio.KeyDown
        If e.Key = Key.Return OrElse e.Key = Key.Tab Then
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando AndAlso
                    MobjObjetoWin.ObjIdServicioShr.ToString() <> txtIdServicio.Text Then
                SAbraServicio()
            End If
        End If
    End Sub

    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If Not MblnPoblandoComboBox AndAlso Not HblnMostrandoDatos Then
            If TypeOf lelmElemento Is ComboBox Then
                Dim lstrPeriodo = String.Empty
                With MobjObjetoWin
                    Select Case lelmElemento.Name
                        Case "cboTipoBaseCalculo"
                            .ObjTipoBaseCalculoByt.ObjValorPro = cboTipoBaseCalculo.SelectedIndex
                        Case "cboAnoPeriodo", "cboMesPeriodo"
                            lstrPeriodo = cboAnoPeriodo.SelectedItem & cboMesPeriodo.SelectedItem
                            .ObjPeriodoInicioStr.ObjValorPro = lstrPeriodo
                        Case "cboTipoTerCtaCr"
                            .ObjIdTipoTerCtaCrSerByt.ObjValorPro = cboTipoTerCtaCr.SelectedIndex
                            txtIdTerceroCr.IsEnabled = False
                            bttEncontrarTerCr.Visibility = Visibility.Collapsed
                            If cboTipoTerCtaCr.SelectedIndex = 1 Then
                                txtIdTerceroCr.IsEnabled = True
                                bttEncontrarTerCr.Visibility = Visibility.Visible
                            End If
                            If cboTipoTerCtaCr.SelectedIndex <> 1 Then
                                txtIdTerceroCr.Text = String.Empty
                            End If
                        Case "cboModoCausaMora"
                            .ObjModoCausaInteresesByt.ObjValorPro =
                                    cboModoCausaMora.SelectedIndex
                            If cboModoCausaMora.SelectedIndex = EnuModoCausaMora.EnuNoCausa Then
                                txtDiasGracia.Text = "0"
                            End If
                    End Select
                End With
                SRegistre()
                SMuestreDatos()
            End If
        End If
    End Sub

    Private Sub Chk_Click(sender As Object, e As RoutedEventArgs) Handles chkEsCalculado.Click,
            chkFacturacionProgramable.Click, chkEsServicioId.Click, chkFactuPropYPreAgr.Click,
            chkGraciaFinMes.Click
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando AndAlso Not HblnMostrandoDatos Then
            If TypeOf sender Is CheckBox Then
                Dim lchkCheckB As CheckBox = sender
                With MobjObjetoWin
                    Select Case lchkCheckB.Name
                        Case "chkFacturacionProgramable"
                            .ObjEsFactProgramableBln.ObjValorPro = chkFacturacionProgramable.IsChecked
                            chkEsCalculado.IsChecked = False
                            .ObjTipoBaseCalculoByt.ObjValorPro = 0
                            chkEsCalculado.IsEnabled = chkFacturacionProgramable.IsChecked
                            SPuebleCombosPeriodo()
                        Case "chkFactuPropYPreAgr"
                            .ObjFactAPropYPreAgrBln.ObjValorPro = chkFactuPropYPreAgr.IsChecked
                        Case "chkEsCalculado"
                            .ObjGeneraProgramBln.ObjValorPro = chkEsCalculado.IsChecked
                            If Not chkEsCalculado.IsChecked Then
                                cboTipoBaseCalculo.SelectedIndex = 0
                                .ObjGeneraProgramBln.ObjValorPro = False
                                .ObjCantPeriodos_ServicioShr.ObjValorPro = 0
                                .ObjPeriodoInicioStr.ObjValorPro = GCSTRPERIODONULO
                            End If
                        Case "chkGraciaFinMes"
                            .ObjGraciaFinMesBln.ObjValorPro = chkGraciaFinMes.IsChecked
                            If .ObjGraciaFinMesBln.ObjValorPro Then
                                txtDiasGracia.Text = 0
                            End If
                        Case "chkEsServicioId"
                            .ObjEsServicioIdBln.ObjValorPro = chkEsServicioId.IsChecked
                            If chkEsServicioId.IsChecked Then
                                cboModoCausaMora.Visibility = Visibility.Hidden
                                cboModoCausaMora.SelectedIndex = EnuModoCausaMora.EnuNoCausa
                                chkFacturacionProgramable.IsChecked = False
                                chkEsCalculado.IsChecked = False
                                chkFacturacionProgramable.Visibility = Visibility.Hidden
                                chkGraciaFinMes.IsChecked = False
                                chkEsCalculado.Visibility = Visibility.Hidden
                                .ObjEsFactProgramableBln.ObjValorPro = False
                                .ObjGeneraProgramBln.ObjValorPro = False
                            Else
                                cboModoCausaMora.Visibility = Visibility.Visible
                                chkFacturacionProgramable.Visibility = Visibility.Visible
                                chkEsCalculado.Visibility = Visibility.Visible
                                chkGraciaFinMes.Visibility = Visibility.Visible
                            End If
                    End Select
                End With
                SVisibiliceControles()
                SMuestreDatos()
            End If
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

    Private Sub Tbc_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) _
            Handles tbcServicios.SelectionChanged
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If Not MblnPoblandoComboBox AndAlso TypeOf lelmElemento Is TabControl Then
            If MobjObjetoWin IsNot Nothing Then
                Select Case True
                    Case tbiGenerales.IsSelected
                        SMuestreGenerales()
                    Case tbiValores.IsSelected
                        SMuestreValores()
                    Case tbiParContables.IsSelected
                        SMuestreParContables()
                End Select
            End If
        End If
    End Sub
#End Region
End Class