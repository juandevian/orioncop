Public Class WinCuentaClientes
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuEstadoDeu = 0
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsCliente = Nothing
    Private ReadOnly MobjPredioAgr As New ClsPredio(EnuModoInstanciaObjDef.enuUnico)
    Private MdtbFacturasEstado As DataTable = Nothing
    Private MstrIdPredioAgr As String = String.Empty
    Private MblnPoblandoCbo As Boolean = False
    Private MobjEstadoCuenta As ClsEstadoCuenta = Nothing
    Private MnuReportes As MenuItem = Nothing
    Private MnuRepEstado As MenuItem = Nothing
    Private MnuRepMovito As MenuItem = Nothing
    Private MblnEsValidoEstado As Boolean = False
    Private MstrServicio As String = "A"
    Private MstrKeySer As String() = Array.Empty(Of String)
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomEstaCta
    '
    Friend Property StrIdPredioAgrupador As String = String.Empty
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuEstadoCuenta
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolControlesLlave As New Collection From {
            txtIdCliente,
            cboPredioAgru
        }
        SAdicioneCtlsRestingidos()
        SCargueForma(EnuElementosAdicionalesDef.None + EnuElementosAdicionalesDef.enuImprimir, 1,
                lcolControlesLlave, cboEstadoActualDeuda, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SAbraCliente(txtIdCliente.Text, True)
        SValide()
        txtIdCliente.Focus()
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
            ObjObjetoWin = New ClsCliente()
            If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                ObjObjetoWin.SVayaAlPrimero()
            End If
        End If
        MobjObjetoWin = ObjObjetoWin
        MobjObjetoWin.SModifiqueParaEstado()
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuEstadoDeu) = lblEstadoActualDeuda
        '
        dtpFechaDesde.SelectedDate = DateSerial(Year(Date.Today), 1, 1)
        dtpFechaHasta.SelectedDate = Date.Today
        dtpFechaDesdeRec.SelectedDate = DateSerial(Year(Date.Today), 1, 1)
        dtpFechaHastaRec.SelectedDate = Date.Today
        dtpFechaDesdeMora.SelectedDate = DateSerial(Year(Date.Today), 1, 1)
        dtpFechaHastaMora.SelectedDate = Date.Today
        dtpFechaMovDesde.SelectedDate = DateSerial(Year(Date.Today), 1, 1)
        dtpFechaMovHasta.SelectedDate = Date.Today
        If GobjParametros.BlnEFacAutorizado Then
            tbiFras.Header = "Facturas"
            lblFacturasCli.Content = "FACTURAS DEL CLIENTE"
            lblFactCli.Content = "Mostrar las Facturas del Cliente entre el"
            temSaldo.Header = "Saldo Factura"
            lblFactEstado.Content = "Facturas debidas"
        End If
        SPuebleComboBoxes()
        HbttAceptar.TabIndex = 40
        HbttCancelar.TabIndex = 41
        SHabiliteImprimir()
    End Sub
    Private Sub SVacie()
        txtNombreCliente.Content = String.Empty
        MstrIdPredioAgr = String.Empty
        MblnPoblandoCbo = True
        cboPredioAgru.Items.Clear()
        MblnPoblandoCbo = False
        SVacieDeuda()
        SMuestreInforme()
    End Sub
    Protected Overrides Sub SNavegueObj()
        HblnCargandoForma = True
        SPuebleCboPredAgru()
        HblnCargandoForma = False
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
            SLevanteEveNoti("No hay Cuentas de Cliente para ser mostradas!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            txtIdCliente.IsEnabled = False
        End If
        GobjPanDat.SControleProcesoObj(True)
        With MobjObjetoWin
            txtIdCliente.Text = .ObjIdClienteDbl.ToString
            txtNombreCliente.Content = .ObjNombreCompletoStr.ObjValorPro
        End With
        Title = My.Resources.CuentaCliente & My.Resources.DosPuntosEspacio & txtNombreCliente.Content
        If Not HblnCargandoForma Then
            SMuestreInforme()
            SMuestreEstadoDeuda()
        End If
        SValide()
        GobjPanDat.SControleProcesoObj(False)
        If txtIdCliente.IsEnabled Then
            If txtIdCliente.Focus() Then
                txtIdCliente.SelectAll()
            End If
        End If
    End Sub
    Protected Overrides Sub SRefresqueWin()
        MyBase.SRefresqueWin()
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.CenuConsultando Then
            SInicialiceValido()
        Else
            MblnEsValidoEstado = (cboEstadoActualDeuda.SelectedIndex > 0) OrElse
                    cboPredioAgru.Items.Count = 0
            StcValidValido(EnuValidEntradaDef.enuEstadoDeu) = MblnEsValidoEstado
        End If
        SHabiliteBotonesTlb()
        SHabiliteImprimir()
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
        ' Adicionar Menues Reportes
        MnuReportes = FmnuiMenuItem("MnuReportes", "R_eportes", "RecMnuItemPriInf")
        MenuVen.Items.Insert(1, MnuReportes)
        MnuRepMovito = FmnuiMenuItem("MnuRepMovito", "Mo_vimiento de la Cuenta", "RecMnuItemSec")
        MnuRepMovito.ToolTip = "Genera un Reporte del movimiento de la Cuenta."
        MnuReportes.Items.Add(MnuRepMovito)
        MnuRepEstado = FmnuiMenuItem("MnuRepEstado", "Estado de Cuenta", "RecMnuItemSec")
        MnuRepEstado.ToolTip = "Genera un Reporte del Estado actual de la Cuenta."
        MnuReportes.Items.Add(MnuRepEstado)
        Dim lmnuRepDetServicio = FmnuiMenuItem("MnuRepDetServicio", "Estado Cta. detallado por Servicio",
                "RecMnuItemSec")
        MnuReportes.Items.Add(lmnuRepDetServicio)
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SGuarde()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim lstrIdPreAgr = String.Empty
        If cboPredioAgru.SelectedIndex > 0 Then
            lstrIdPreAgr = cboPredioAgru.SelectedItem
        End If
        Try
            GobjPanDat.SControleProcesoObj(True)
            If EnuOperacionEnWin = EnuOperacionEnVentana.CenuModificando Then
                If FblnEstanTodosBien() Then
                    MobjObjetoWin.SCambieEstadoDeuda(cboEstadoActualDeuda.SelectedIndex, lstrIdPreAgr)
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
                If Not EnuIdVentana = EnuIdVentanaDef.EnuLogOn Then
                    SFinaliceOperacion()
                    SPuebleCboPredAgru()
                    SRefrescarClic()
                End If
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
                SFinaliceOperacion()
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Protected Overrides Sub SImprima()
        If MobjObjetoWin.BlnExiste Then
            Mouse.OverrideCursor = Cursors.Wait
            SLevanteEveNoti("Imprimiendo", String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            GobjPanDat.SControleProcesoObj(True)
            If tbcEstadoCta.SelectedIndex = 0 Then
                SImprimaEstado()
            ElseIf tbcEstadoCta.SelectedIndex = 1 Then
                SImprimaMovimiento()
            End If
            GobjPanDat.SControleProcesoObj(False)
            Mouse.OverrideCursor = Cursors.Arrow
            SLevanteEveNoti(String.Empty, String.Empty, 0, EnuSeveridadNot.EnuOk)
        End If
    End Sub
#End Region
#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable Then
                    txtIdCliente.Text = StrResultadoBusqueda
                    SAbraCliente(StrResultadoBusqueda, False)
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
        SDefineBusquedaPredioAgr()
        SDefineBusquedaNombreCompleto()
        Return True
    End Function
    Private Sub SDefineBusquedaPredioAgr()
        Dim lstrTablaPri As String = ClsFactura.SstrNombreTabla
        Dim lstrCamposTabPri As String() = {"DISTINCT " & ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd,
                ClsIdCliente_FactDbl.SstrNombreCampoBd}
        Dim lstrCamTabPriRel As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                 ClsIdCliente_FactDbl.SstrNombreCampoBd}
        Dim lstrTablaSec = ClsCliente.SstrNombreTabla
        Dim lstrCamposTabSec = {ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCamTabSecRel = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri
        Dim lstrCampoBusqueda As String = ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdCliente_FactDbl.SstrNombreCampoBd
        HwinBusqueda.SDefinaBusqueda("Predio Agrupador", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, True)
    End Sub
    Private Sub SDefineBusquedaNombreCompleto()
        Dim lstrTablaPri = ClsCliente.SstrNombreTabla
        Dim lstrCamposTabPri = {"DISTINCT " & ClsNombreCompletoStr.SstrNombreCampoBd,
                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrCamTabPriRel = {StrCampoCarpeta, StrCampoCentroUtil,
                ClsIdClienteDbl.SstrNombreCampoBd}
        Dim lstrTablaSec As String = ClsFactura.SstrNombreTabla
        Dim lstrCamposTabSec As String() = {ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd}
        Dim lstrCamTabSecRel As String() = {StrCampoCarpeta, StrCampoCentroUtil,
                 ClsIdCliente_FactDbl.SstrNombreCampoBd}
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdClienteDbl.SstrNombreCampoBd
        HwinBusqueda.SDefinaBusqueda("Nombre Cliente", lstrTablaPri, lstrTablaSec,
                lstrCamposTabPri, lstrCamposTabSec, lstrCamTabPriRel, lstrCamTabSecRel,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro, False)
    End Sub
#End Region
#Region "Procedimientos Propios"
#Region "Generales"
    Private Sub SAdicioneCtlsRestingidos()
        SAdicioneControlRestringido(dtpFechaDesde)
        SAdicioneControlRestringido(dtpFechaHasta)
        SAdicioneControlRestringido(dtpFechaDesdeRec)
        SAdicioneControlRestringido(dtpFechaHastaRec)
        SAdicioneControlRestringido(dtpFechaDesdeMora)
        SAdicioneControlRestringido(dtpFechaHastaMora)
        SAdicioneControlRestringido(dtpFechaMovDesde)
        SAdicioneControlRestringido(dtpFechaMovHasta)
        SAdicioneControlRestringido(cboPredioAgru)
        SAdicioneControlRestringido(cboServicio)
        SAdicioneControlRestringido(dgrFacturasEstado)
        SAdicioneControlRestringido(dgrFacturasCliente)
        SAdicioneControlRestringido(dgrRecibosCliente)
        SAdicioneControlRestringido(dgrAnticipos)
        SAdicioneControlRestringido(dgrNotasCr)
        SAdicioneControlRestringido(dgrNotasDbCliente)
        SAdicioneControlRestringido(dgrMovimiento)
        SAdicioneControlRestringido(rdbMovAnticipos)
        SAdicioneControlRestringido(rdbMovDeuda)
    End Sub
    Private Sub SAbraCliente(astrIdCliente As String, ablnCargandoVentana As Boolean)
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                If MobjObjetoWin.ObjIdClienteDbl.ToString <> astrIdCliente Then
                    MobjObjetoWin.SAbra({GshrIdCarpeta, GshrIdCentroUtil, astrIdCliente})
                    If Not MobjObjetoWin.BlnExiste Then
                        MobjPredioAgr.SVacie()
                        SVacie()
                    Else
                        SPuebleCboPredAgru()
                    End If
                ElseIf ablnCargandoVentana Then
                    SPuebleCboPredAgru()
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
    Private Sub SPuebleCboPredAgru()
        If Not IsNothing(MobjObjetoWin) AndAlso MobjObjetoWin.BlnExiste Then
            Dim larlPrediosAgr = MobjObjetoWin.FarlPrediosAgrDelCliente
            If Not larlPrediosAgr.Contains("") Then
                larlPrediosAgr.Add("")
            End If
            MblnPoblandoCbo = True
            cboPredioAgru.Items.Clear()
            For Each lstrPreAgr As String In larlPrediosAgr
                If String.IsNullOrEmpty(lstrPreAgr) Then
                    cboPredioAgru.Items.Add(GCSTRSINPA)
                Else
                    If lstrPreAgr.Contains(",") Then
                        Dim lstrPredios As String() = lstrPreAgr.Split(",")
                        For Each lstrIdPre As String In lstrPredios
                            If String.IsNullOrEmpty(lstrIdPre) Then
                                lstrIdPre = GCSTRSINPA
                            End If
                            If Not cboPredioAgru.Items.Contains(lstrIdPre) Then
                                cboPredioAgru.Items.Add(lstrIdPre)
                            End If
                        Next
                    Else
                        If Not cboPredioAgru.Items.Contains(lstrPreAgr) Then
                            cboPredioAgru.Items.Add(lstrPreAgr)
                        End If
                    End If
                End If
            Next
            If cboPredioAgru.SelectedIndex = 0 Then
                cboPredioAgru.SelectedIndex = -1
            End If
            MblnPoblandoCbo = False
            If String.IsNullOrEmpty(StrIdPredioAgrupador) Then
                If cboPredioAgru.Items.Count >= 2 Then
                    cboPredioAgru.SelectedIndex = 1
                Else
                    cboPredioAgru.SelectedIndex = 0
                End If
            Else
                cboPredioAgru.SelectedItem = StrIdPredioAgrupador
            End If
        End If
        If cboPredioAgru.Items.Count > 1 Then
            Dim lstrMens = "El presente Cliente tiene cuentas con más de un Predio!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuAdvertencia)
        End If
    End Sub
    Private Sub SPuebleComboBoxes()
        MblnPoblandoCbo = True
        Dim ldrwEstadoDeuda = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuEstadoDeuda)
        SPuebleComboBox(ldrwEstadoDeuda, cboEstadoActualDeuda)
        MblnPoblandoCbo = False
    End Sub
    Private Sub SPuebleCboServicios()
        MblnPoblandoCbo = True
        cboServicio.Items.Clear()
        cboServicio.Items.Add(My.Resources.Todos)
        Dim lstrIdpreAgr As String, i = 0
        Dim lshrIdAno As Short, lshridSer As Short
        Dim lstrKeySer As String
        lstrIdpreAgr = If(cboPredioAgru.SelectedItem = GCSTRSINPA, String.Empty,
                cboPredioAgru.SelectedItem)
        Dim ldtbServClieConDeuda = MobjObjetoWin.FdtbServiciosConDeuda({lstrIdpreAgr})
        ReDim MstrKeySer(ldtbServClieConDeuda.Rows.Count)
        MstrKeySer(i) = "A"
        For Each ldrwSer As DataRow In ldtbServClieConDeuda.Rows
            lshrIdAno = ClsPanorama.FobjValorCampo(ldrwSer(
                    ClsIdAno_ServicioItemFactShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            lshridSer = ClsPanorama.FobjValorCampo(ldrwSer(
                    ClsIdServicio_ItemFactShr.SstrNombreCampoBd), EnuTipoValor.enuShort)
            lstrKeySer = If(lshrIdAno > 0, "0", lshridSer.ToString())
            If Not MstrKeySer.Contains(lstrKeySer) Then
                i += 1
                cboServicio.Items.Add(GobjParametros.FobjServicio(
                        lshrIdAno.ToString() & "," &
                        lshridSer.ToString()).ObjConceptoServicioStr.ObjValorPro)
                MstrKeySer(i) = lstrKeySer
            End If
        Next
        If cboServicio.SelectedIndex = 0 Then
            MstrServicio = MstrKeySer(0)
        End If
        cboServicio.SelectedIndex = 0
        MblnPoblandoCbo = False
    End Sub
    Private Sub SHabiliteImprimir()
        Dim lblnImprimir = (tbcEstadoCta.SelectedIndex <= 1)
        If lblnImprimir Then
            SHabiliteBotonTlb(True, HbttImprimir)
            SHabiliteMenuItem(True, HmnuImprimir)
        End If
    End Sub
    Private Sub SMuestreInforme()
        GobjPanDat.SControleProcesoObj(True)
        Select Case tbcEstadoCta.SelectedIndex
            Case 0
                SMuestreEstado()
            Case 1
                If rdbMovDeuda.IsChecked Then
                    SMuestreMovimiento()
                Else
                    SMuestreMovimientoAnt()
                End If
            Case 2
                SMuestreFacturas()
            Case 3
                SMuestreRecibos()
            Case 4
                SMuestreAnticipos()
            Case 5
                SMuestreNotasCr()
            Case 6
                SMuestreNotasDb()
        End Select
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Private Sub SMuestreEstadoDeuda()
        If Not IsNothing(MobjObjetoWin) Then
            cboEstadoActualDeuda.SelectedIndex = MobjObjetoWin.FenuEstadoDeuda(MstrIdPredioAgr)
            If Not String.IsNullOrEmpty(MstrIdPredioAgr) Then
                MobjPredioAgr.SAbra({GshrIdCarpeta, GshrIdCentroUtil, MstrIdPredioAgr})
                Dim ldblIdClienteConDeuda = MobjPredioAgr.FdblIdClienteDiferenteConDeuda(
                        MobjObjetoWin.ObjIdClienteDbl.ObjValorPro)
                If ldblIdClienteConDeuda > 0 Then
                    Dim lstrMens = "El Predio Agrupador actual tiene deuda " &
                            "con el Cliente " & ldblIdClienteConDeuda & "!"
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuAdvertencia)
                End If
            End If
        End If
    End Sub
    Private Sub SImprimaEstado()
        Dim ldecIntPorCausar As Decimal
        Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
            .EnuReporte = EnuReporteDef.enuEstadoCtaCli
            }
        ldecIntPorCausar = MobjObjetoWin.FdecIntMoraPorCausar({MstrIdPredioAgr},
                Date.Today)
        With lobjRep
            Dim lstrIdPredAgr = String.Empty
            .DblIdCliente = MobjObjetoWin.ObjIdClienteDbl.ObjValorPro
            If cboPredioAgru.SelectedItem <> GCSTRSINPA Then
                lstrIdPredAgr = cboPredioAgru.SelectedItem
            End If
            .StrIdPredioAgru = lstrIdPredAgr
            .DecIntPortCausar = ldecIntPorCausar
            .SGenereReporte()
        End With
    End Sub
    Private Sub SImprimaMovimiento()
        Dim ldtbMovimiento As DataTable = Nothing, lstrMens = String.Empty
        Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO)
        If rdbMovDeuda.IsChecked Then
            ldtbMovimiento = MobjObjetoWin.FdtbMovimientoDeuda(MstrIdPredioAgr,
                    dtpFechaMovDesde.SelectedDate, dtpFechaMovHasta.SelectedDate)
            lobjRep.EnuReporte = EnuReporteDef.enuMovimCuenta
        Else
            ldtbMovimiento = MobjObjetoWin.FdtbMovimientoAnticipos(MstrIdPredioAgr,
                    dtpFechaMovDesde.SelectedDate, dtpFechaMovHasta.SelectedDate)
            lobjRep.EnuReporte = EnuReporteDef.enuMovimAntici
        End If
        With lobjRep
            .DsMovimiento = ldtbMovimiento.DataSet
            .DblIdCliente = MobjObjetoWin.ObjIdClienteDbl.ObjValorPro
            .SGenereReporte()
        End With
    End Sub
    Private Sub SGenereRepDetServicio()
        Dim lblnTieneDeuda = MobjObjetoWin.FblnTieneDeuda(False)
        If lblnTieneDeuda Then
            Dim lobjPara As New ClsParametrosReportesDocs("", 0, 0) With {
            .DblIdTercero = MobjObjetoWin.ObjIdClienteDbl.ObjValorPro
        }
            Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
            .ObjParRepDocs = lobjPara,
            .EnuReporte = EnuReporteDef.enuCxCDetPorSer
            }
            lobjRep.SGenereReporte()
        Else
            SLevanteEveNoti("El Cliente no tiene Deuda actualmente!", String.Empty, 0,
                    EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
#End Region
#Region "Estado"
    Private Sub SMuestreEstado()
        If Not IsNothing(MobjObjetoWin) AndAlso MobjObjetoWin.BlnExiste Then
            Mouse.OverrideCursor = Cursors.Wait
            Dim ldecIntPorCausar = MobjObjetoWin.FdecIntMoraPorCausar(
                    {MstrIdPredioAgr}, Date.Today)
            MobjEstadoCuenta = MobjObjetoWin.FobjEstadoCtaHoy(MstrIdPredioAgr, MstrServicio,
                    MdtbFacturasEstado)
            If Not IsNothing(MobjEstadoCuenta) Then
                With MobjEstadoCuenta
                    txtDeudaCapital.Content = .ObjDeudaCapitalDec.ToString
                    txtDeudaMora.Content = .ObjDeudaIntMoraDec.ToString
                    txtIntPorCausar.Content = Format(ldecIntPorCausar, "c")
                    txtTotalDeuda.Content = Format(.DecTotalDeuda + ldecIntPorCausar, "c")
                End With
            Else
                SVacieDeuda()
            End If
            tbiEstadoCta.DataContext = MdtbFacturasEstado
            dgrFacturasEstado.SelectedIndex = 0
            txtFechaEstado.Content = Format(Date.Now, "dd/MM/yyyy hh:mm:ss tt")
            Mouse.OverrideCursor = Cursors.Arrow
        Else
            tbiEstadoCta.DataContext = Nothing
        End If
    End Sub
    Private Sub SVacieDeuda()
        txtFechaEstado.Content = String.Empty
        txtDeudaCapital.Content = Format(0, "c")
        txtDeudaMora.Content = Format(0, "c")
        txtTotalDeuda.Content = Format(0, "c")
        txtIntPorCausar.Content = Format(0, "c")
        If Not IsNothing(MdtbFacturasEstado) Then
            MdtbFacturasEstado.Rows.Clear()
        End If
        tbiEstadoCta.DataContext = MdtbFacturasEstado
    End Sub
#End Region
#Region "Determinar DataContext"
    Private Sub SMuestreFacturas()
        If MobjObjetoWin.BlnExiste Then
            Dim ldblIdCliente As Double = MobjObjetoWin.ObjIdClienteDbl.ObjValorPro
            dgrFacturasCliente.DataContext = ClsOrionCop.FdtbFacturas(ldblIdCliente,
                    MstrIdPredioAgr, dtpFechaDesde.SelectedDate, dtpFechaHasta.SelectedDate)
        Else
            dgrFacturasCliente.DataContext = Nothing
        End If
    End Sub
    Private Sub SMuestreRecibos()
        If MobjObjetoWin.BlnExiste Then
            tbiRecCaj.DataContext = MobjObjetoWin.FdtbRecibos(MstrIdPredioAgr,
                    dtpFechaDesdeRec.SelectedDate, dtpFechaHastaRec.SelectedDate)
            dgrRecibosCliente.SelectedIndex = 0
        Else
            tbiRecCaj.DataContext = Nothing
        End If
    End Sub
    Private Sub SMuestreAnticipos()
        If MobjObjetoWin.BlnExiste Then
            dgrAnticipos.DataContext = MobjObjetoWin.DtbAnticipos(MstrIdPredioAgr)
        Else
            dgrAnticipos.DataContext = Nothing
        End If
    End Sub
    Private Sub SMuestreNotasCr()
        If MobjObjetoWin.BlnExiste Then
            dgrNotasCr.DataContext = MobjObjetoWin.FdtbNotasCr(MstrIdPredioAgr)
        Else
            dgrNotasCr.DataContext = Nothing
        End If
    End Sub
    Private Sub SMuestreNotasDb()
        If MobjObjetoWin.BlnExiste Then
            dgrNotasDbCliente.DataContext = MobjObjetoWin.FdtbNotasDb(MstrIdPredioAgr,
                    dtpFechaDesdeMora.SelectedDate, dtpFechaHastaMora.SelectedDate)
        Else
            dgrNotasDbCliente.DataContext = Nothing
        End If
    End Sub
    Private Sub SMuestreMovimiento()
        If MobjObjetoWin.BlnExiste Then
            dgrMovimiento.DataContext = MobjObjetoWin.FdtbMovimientoDeuda(MstrIdPredioAgr,
                    dtpFechaMovDesde.SelectedDate, dtpFechaMovHasta.SelectedDate)
            SValide()
        Else
            dgrMovimiento.DataContext = Nothing
        End If
    End Sub
    Private Sub SMuestreMovimientoAnt()
        If MobjObjetoWin.BlnExiste Then
            dgrMovimiento.DataContext = MobjObjetoWin.FdtbMovimientoAnticipos(MstrIdPredioAgr,
                        dtpFechaMovDesde.SelectedDate, dtpFechaMovHasta.SelectedDate)
            SValide()
        Else
            dgrMovimiento.DataContext = Nothing
        End If
    End Sub
#End Region
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            Select Case lelmElemento.Name
                Case "MnuRepEstado", "MnuRepMovito"
                    SImprima()
                Case "MnuRepDetServicio"
                    SGenereRepDetServicio()
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
            If TypeOf lelmElemento Is TextBox Then
                If lelmElemento.Name = "txtIdCliente" Then
                    SAbraCliente(txtIdCliente.Text, False)
                End If
            End If
        End If
    End Sub
    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If IsNothing(lelmElemento) AndAlso TypeOf sender Is ComboBox Then
            lelmElemento = sender
        End If
        If Not MblnPoblandoCbo AndAlso Not IsNothing(lelmElemento) Then
            If TypeOf lelmElemento Is ComboBox Then
                Select Case lelmElemento.Name
                    Case "cboPredioAgru"
                        If cboPredioAgru.SelectedItem = GCSTRSINPA Then
                            MstrIdPredioAgr = String.Empty
                            Title = My.Resources.CuentaCliente & My.Resources.DosPuntosEspacio & txtNombreCliente.Content
                        Else
                            MstrIdPredioAgr = cboPredioAgru.SelectedItem
                            Title = My.Resources.CuentaCliente &
                                    My.Resources.DosPuntosEspacio &
                                    txtNombreCliente.Content & My.Resources.Separador &
                                    MstrIdPredioAgr
                        End If
                        SPuebleCboServicios()
                        txtEstadoSugeridoDeuda.Content = ClsOrionCop.FstrNombreDatoConstanteOri(
                                EnuGrupoConstantesOriDef.EnuEstadoDeuda,
                                MobjObjetoWin.FenuEstadoSugeridoDeuda(MstrIdPredioAgr))
                    Case "cboServicio"
                        cboServicio.ToolTip = cboServicio.SelectedItem
                        Dim lstrSer As String = MstrKeySer(cboServicio.SelectedIndex)
                        MstrServicio = lstrSer
                End Select
                If Not HblnCargandoForma Then
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub
    Private Sub OnRatonUp(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Select Case lelmElemento.Name
            Case "tbiEstadoCta"
                SMuestreEstado()
            Case "tbiMovto"
                If rdbMovDeuda.IsChecked Then
                    SMuestreMovimiento()
                Else
                    SMuestreMovimientoAnt()
                End If
            Case "tbiFras"
                SMuestreFacturas()
            Case "tbiAnti"
                SMuestreAnticipos()
            Case "tbiNCR"
                SMuestreNotasCr()
            Case "tbiNDB"
                SMuestreNotasDb()
            Case "tbiRecCaj"
                SMuestreRecibos()
        End Select
    End Sub
    Private Sub TxtIdCliente_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdCliente.KeyDown
        If e.Key = Key.Return Then
            If txtIdCliente.Focus Then
                If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                    SAbraCliente(txtIdCliente.Text, False)
                End If
            End If
        End If
    End Sub
    Private Sub DtpFecha_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles dtpFechaDesde.SelectedDateChanged,
            dtpFechaHasta.SelectedDateChanged, dtpFechaDesdeRec.SelectedDateChanged,
            dtpFechaHastaRec.SelectedDateChanged, dtpFechaDesdeMora.SelectedDateChanged,
            dtpFechaHastaMora.SelectedDateChanged, dtpFechaMovDesde.SelectedDateChanged,
            dtpFechaMovHasta.SelectedDateChanged
        If TypeOf sender Is DatePicker Then
            If Not HblnCargandoForma Then
                SMuestreDatos()
            End If
        End If
    End Sub
    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
                dgrFacturasEstado.MouseRightButtonUp, dgrFacturasCliente.MouseRightButtonUp,
                dgrRecibosCliente.MouseRightButtonUp, dgrAnticipos.MouseRightButtonUp,
                dgrNotasCr.MouseRightButtonUp, dgrNotasDbCliente.MouseRightButtonUp,
                dgrMovimiento.MouseRightButtonUp
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            Dim lstrPrefijo = String.Empty
            Dim lentIdObjeto = 0
            ldrvFilaActual = ldgrActual.SelectedItem
            If Not IsNothing(ldrvFilaActual) AndAlso ldgrActual.SelectedIndex >= 0 Then
                Select Case ldgrActual.Name
                    Case "dgrFacturasEstado"
                        lstrPrefijo = ldrvFilaActual("Prefijo")
                        lentIdObjeto = ldrvFilaActual("IdFactura")
                        SAbraFactura(lstrPrefijo, lentIdObjeto)
                    Case "dgrMovimiento"
                        If ldgrActual.SelectedIndex > 0 Then
                            Dim ldrwMovCta = ldrvFilaActual.Row
                            Dim lstrNroDoc As String
                            If rdbMovDeuda.IsChecked Then
                                lstrNroDoc = ClsPanorama.FobjValorCampo(ldrwMovCta("NroDocOri"),
                                    EnuTipoValor.enuString)
                            Else
                                Dim lstrPref As String = ClsPanorama.FobjValorCampo(ldrwMovCta(
                                        ClsPrefijoDocOrigen_NovAntStr.SstrNombreCampoBd), EnuTipoValor.enuString)
                                Dim lentIdDoc As Integer = ClsPanorama.FobjValorCampo(ldrwMovCta(
                                        ClsIdDocOrigen_NovAntEnt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                                lstrNroDoc = ClsPanorama.FstrNumeroDcto(lstrPrefijo, lentIdDoc)
                            End If
                            Dim lenuIdTipoDoc As EnuTipoDocOri = ClsPanorama.FobjValorCampo(ldrwMovCta(
                                    ClsIdTipoDocOrigenByt.SstrNombreCampoBd), EnuTipoValor.enuInteger)
                            If Not IsDBNull(lstrNroDoc) AndAlso Not IsNothing(lstrNroDoc) Then
                                lstrPrefijo = ClsPanorama.FstrPrefijoDcto(lstrNroDoc)
                                lentIdObjeto = ClsPanorama.FentIdDcto(lstrNroDoc)
                            End If
                            Select Case lenuIdTipoDoc
                                Case EnuTipoDocOri.EnuFactura
                                    SAbraFactura(lstrPrefijo, lentIdObjeto)
                                Case EnuTipoDocOri.EnuNotaCon
                                    SAbraNotaCon(lstrPrefijo, lentIdObjeto)
                                Case EnuTipoDocOri.EnuNotaCr
                                    SAbraNotaCr(lstrPrefijo, lentIdObjeto)
                                Case EnuTipoDocOri.EnuNotaDb
                                    SAbraNotaDb(lstrPrefijo, lentIdObjeto)
                                Case EnuTipoDocOri.EnuNotaDevAnt
                                    SAbraNotaDevAnt(lstrPrefijo, lentIdObjeto)
                                Case EnuTipoDocOri.EnuReciboCaja
                                    SAbraRecibo(lstrPrefijo, lentIdObjeto)
                                Case EnuTipoDocOri.EnuNotaRevCr
                                    SAbraNotaRCr(lstrPrefijo, lentIdObjeto)
                                Case EnuTipoDocOri.None
                                    Dim lstrMens = "No se ha seleccionado una Fila que se relacione " &
                                        "con un Documento!"
                                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                            End Select
                        End If
                    Case "dgrFacturasCliente"
                        lstrPrefijo = ldrvFilaActual(ClsPrefijo_FactStr.SstrNombreCampoBd)
                        lentIdObjeto = ldrvFilaActual(ClsIdFacturaEnt.SstrNombreCampoBd)
                        SAbraFactura(lstrPrefijo, lentIdObjeto)
                    Case "dgrRecibosCliente"
                        lstrPrefijo = ldrvFilaActual(ClsPrefijo_RecStr.SstrNombreCampoBd)
                        lentIdObjeto = ldrvFilaActual(ClsIdRecCajaEnt.SstrNombreCampoBd)
                        SAbraRecibo(lstrPrefijo, lentIdObjeto)
                    Case "dgrAnticipos"
                        lentIdObjeto = ldrvFilaActual(ClsIdAnticipoEnt.SstrNombreCampoBd)
                        SAbraAnticipo(lentIdObjeto)
                    Case "dgrNotasCr"
                        lstrPrefijo = ldrvFilaActual(ClsPrefijo_NotaCrStr.SstrNombreCampoBd)
                        lentIdObjeto = ldrvFilaActual(ClsIdNotaCrEnt.SstrNombreCampoBd)
                        SAbraNotaCr(lstrPrefijo, lentIdObjeto)
                    Case "dgrNotasDbCliente"
                        lstrPrefijo = ldrvFilaActual(ClsPrefijo_NotaDbStr.SstrNombreCampoBd)
                        lentIdObjeto = ldrvFilaActual(ClsIdNotaDbEnt.SstrNombreCampoBd)
                        SAbraNotaDb(lstrPrefijo, lentIdObjeto)
                End Select
            End If
        End If
    End Sub
    Private Sub Rdb_Click(sender As Object, e As RoutedEventArgs) Handles rdbMovDeuda.Click, rdbMovAnticipos.Click
        SMuestreDatos()
    End Sub
#End Region
End Class
