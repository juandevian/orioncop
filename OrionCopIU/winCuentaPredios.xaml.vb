Public Class WinCuentaPredios
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
#End Region
    ' Variables
    Private MobjObjetoWin As ClsPredio = Nothing
    Private ReadOnly MobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
    Private MdtbFacturasEstado As DataTable = Nothing
    Private MblnPoblandoCbo As Boolean = False
    Private MobjEstadoCuenta As ClsEstadoCuenta = Nothing
    Private MstrServicio As String = "A"
    Private MstrKeySer As String() = Nothing
    Private MnuReportes As MenuItem = Nothing
    Private MnuRepEstado As MenuItem = Nothing
    Private MnuRepMovito As MenuItem = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomEstaCta
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
            txtIdPredio
        }
        SAdicioneCtlsRestingidos()
        SCargueForma(EnuElementosAdicionalesDef.None + EnuElementosAdicionalesDef.enuImprimir, 0,
                lcolControlesLlave, cboCliente, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SAbraPredio(txtIdPredio.Text, True)
        txtIdPredio.Focus()
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
            ObjObjetoWin = New ClsPredio(EnuModoInstanciaObjDef.enuNavegable)
            If Not ObjObjetoWin.FblnEstaVacioOrigenDatos Then
                ObjObjetoWin.SVayaAlPrimero()
            End If
        End If
        MobjObjetoWin = ObjObjetoWin
        MobjObjetoWin.SModifiqueParaEstado()
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
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
        HbttAceptar.TabIndex = 40
        HbttCancelar.TabIndex = 41
        SHabiliteImprimir()
    End Sub
    Private Sub SVacie()
        txtNombreCliente.Content = String.Empty
        MblnPoblandoCbo = True
        cboCliente.Items.Clear()
        MblnPoblandoCbo = False
        SVacieDeuda()
        SMuestreInforme()
    End Sub
    Protected Overrides Sub SNavegueObj()
        HblnCargandoForma = True
        SPuebleCboClientes()
        HblnCargandoForma = False
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
            SLevanteEveNoti("No hay Cuentas de Cliente para ser mostradas!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            txtIdPredio.IsEnabled = False
        End If
        GobjPanDat.SControleProcesoObj(True)
        With MobjObjetoWin
            If txtIdPredio.Text <> .ObjIdPredioStr.ObjValorPro Then
                txtIdPredio.Text = .ObjIdPredioStr.ObjValorPro
                SAbraPredio(.ObjIdPredioStr.ObjValorPro, HblnCargandoForma)
            End If
        End With
        Title = My.Resources.CuentaPredio & My.Resources.DosPuntosEspacio & txtIdPredio.Text
        If Not HblnCargandoForma Then
            SMuestreInforme()
            SMuestreEstadoDeuda()
        End If
        SValide()
        If txtIdPredio.IsEnabled Then
            If txtIdPredio.Focus() Then
                txtIdPredio.SelectAll()
            End If
        End If
        If cboCliente.Items.Count > 1 Then
            Dim lstrMens = "El presente Predio tiene cuentas con más de un Cliente!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuAdvertencia)
        End If
        Title = My.Resources.CuentaPredio & My.Resources.DosPuntosEspacio &
                                txtIdPredio.Text & My.Resources.Separador &
                                txtNombreCliente.Content
        GobjPanDat.SControleProcesoObj(False)
    End Sub
    Protected Overrides Sub SRefresqueWin()
        MyBase.SRefresqueWin()
    End Sub
    Protected Overrides Sub SValide()
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
        MnuReportes = FmnuiMenuItem("MnuReportes", "R_eportes", "RecMnuItemPri")
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
    Protected Overrides Sub SImprima()
        If MobjObjetoWin.BlnExiste Then
            Mouse.OverrideCursor = Cursors.Wait
            Dim lstrMens = "Imprimiendo"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            GobjPanDat.SControleProcesoObj(True)
            If tbcEstadoCta.SelectedIndex = 0 Then
                SImprimaEstado()
            ElseIf tbcEstadoCta.SelectedIndex = 1 Then
                SImprimaMovimiento()
            End If
            GobjPanDat.SControleProcesoObj(False)
            SLevanteEveNoti(String.Empty, String.Empty, 0, EnuSeveridadNot.EnuOk)
            Mouse.OverrideCursor = Cursors.Arrow
        End If
    End Sub
#End Region
#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        MyBase.SBuscar()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                If MobjObjetoWin.EnuTipoObjeto = EnuModoInstanciaObjDef.enuNavegable Then
                    SAbraPredio(StrResultadoBusqueda, False)
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
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " <> ''"
        Dim lstrCampoBusqueda As String = ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd
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
        Dim lstrFiltro = ClsOrionCop.StrFiltroUbicacion_Pri & " AND " &
                ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd & " <> ''"
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdPredioAgrupador_FacStr.SstrNombreCampoBd
        HwinBusqueda.SDefinaBusqueda("Nombre Clienter", lstrTablaPri, lstrTablaSec,
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
        SAdicioneControlRestringido(cboCliente)
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
    Private Sub SAbraPredio(astrIdPredio As String, ablnCargandoVentana As Boolean)
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                If MobjObjetoWin.ObjIdPredioStr.ObjValorPro <> astrIdPredio Then
                    MobjObjetoWin.SAbra({GshrIdCarpeta, GshrIdCentroUtil, astrIdPredio})
                    If Not MobjObjetoWin.BlnExiste Then
                        MobjCliente.SVacie()
                        SVacie()
                    Else
                        SPuebleCboClientes()
                        SPuebleCboServicios()
                    End If
                ElseIf ablnCargandoVentana Then
                    If Not IsNothing(MobjObjetoWin) Then
                        SPuebleCboClientes()
                        SPuebleCboServicios()
                    End If
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
    Private Sub SAbraCliente(adblIdCliente As Double)
        MobjCliente.SAbra({GshrIdCarpeta, GshrIdCentroUtil, adblIdCliente})
        If MobjCliente.BlnExiste Then
            txtNombreCliente.Content = MobjCliente.ObjNombreCompletoStr.ObjValorPro
        End If
    End Sub
    Private Sub SPuebleCboClientes()
        If Not IsNothing(MobjObjetoWin) AndAlso MobjObjetoWin.BlnExiste Then
            Dim ldblIdClientesDelPredio As ArrayList, lstrMens As String
            ldblIdClientesDelPredio = MobjObjetoWin.FstrClientesDelPredio
            MblnPoblandoCbo = True
            cboCliente.Items.Clear()
            If Not IsNothing(ldblIdClientesDelPredio) AndAlso ldblIdClientesDelPredio.Count > 0 Then
                For Each ldblCliPre As Double In ldblIdClientesDelPredio
                    cboCliente.Items.Add(ldblCliPre.ToString)
                Next
            Else
                lstrMens = "El presente Predio no tiene Movimiento como PredioAgrupador!"
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                txtNombreCliente.Content = String.Empty
            End If
            If cboCliente.SelectedIndex = 0 Then
                cboCliente.SelectedIndex = -1
            End If
            MblnPoblandoCbo = False
            If cboCliente.Items.Count > 0 Then
                cboCliente.SelectedIndex = 0
            End If
        End If
    End Sub
    Private Sub SPuebleCboServicios()
        MblnPoblandoCbo = True
        cboServicio.Items.Clear()
        If MobjObjetoWin.BlnExiste Then
            cboServicio.Items.Add(My.Resources.Todos)
            Dim i = 0
            Dim lshrIdAno As Short, lshridSer As Short
            Dim lstrKeySer As String
            Dim ldtbServPreConDeuda = MobjCliente.FdtbServiciosConDeuda({txtIdPredio.Text})
            ReDim MstrKeySer(ldtbServPreConDeuda.Rows.Count)
            MstrKeySer(i) = "A"
            For Each ldrwSer As DataRow In ldtbServPreConDeuda.Rows
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
        End If
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
        txtEstadoActual.Content = ClsOrionCop.FstrNombreDatoConstanteOri(
                EnuGrupoConstantesOriDef.EnuEstadoDeuda,
                MobjObjetoWin.ObjIdEstadoDeuda_PredioByt.ObjValorPro)
        txtEstadoSugeridoDeuda.Content = ClsOrionCop.FstrNombreDatoConstanteOri(
                EnuGrupoConstantesOriDef.EnuEstadoDeuda,
                MobjCliente.FenuEstadoSugeridoDeuda(txtIdPredio.Text))
    End Sub
    Private Sub SImprimaEstado()
        Dim ldecIntPorCausar As Decimal
        Dim lstrIdPredAgr As String = MobjObjetoWin.ObjIdPredioStr.ToString
        Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
            .EnuReporte = EnuReporteDef.enuEstadoCtaCli
            }
        ldecIntPorCausar = MobjCliente.FdecIntMoraPorCausar({lstrIdPredAgr}, Date.Today)
        With lobjRep
            .DblIdCliente = MobjCliente.ObjIdClienteDbl.ObjValorPro
            If String.IsNullOrEmpty(lstrIdPredAgr) Then lstrIdPredAgr = GCSTRSINPA
            .StrIdPredioAgru = lstrIdPredAgr
            .DecIntPortCausar = ldecIntPorCausar
            .SGenereReporte()
        End With
    End Sub
    Private Sub SImprimaMovimiento()
        Dim ldtbMovimiento As DataTable = Nothing
        Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO)
        If rdbMovDeuda.IsChecked Then
            ldtbMovimiento = MobjCliente.FdtbMovimientoDeuda(MobjObjetoWin.ObjIdPredioStr.ObjValorPro,
                    dtpFechaMovDesde.SelectedDate, dtpFechaMovHasta.SelectedDate)
            lobjRep.EnuReporte = EnuReporteDef.enuMovimCuenta
        Else
            ldtbMovimiento = MobjCliente.FdtbMovimientoAnticipos(MobjObjetoWin.ObjIdPredioStr.ObjValorPro,
                    dtpFechaMovDesde.SelectedDate, dtpFechaMovHasta.SelectedDate)
            lobjRep.EnuReporte = EnuReporteDef.enuMovimAntici
        End If
        With lobjRep
            .DsMovimiento = ldtbMovimiento.DataSet
            .DblIdCliente = MobjCliente.ObjIdClienteDbl.ObjValorPro
            .SGenereReporte()
        End With
    End Sub
    Private Sub SGenereRepDetServicio()
        Dim lobjPara As New ClsParametrosReportesDocs("", 0, 0) With {
            .StrIdPredioAgr = MobjObjetoWin.ObjIdPredioAgrupadorStr.ObjValorPro
        }
        Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO) With {
            .ObjParRepDocs = lobjPara,
            .EnuReporte = EnuReporteDef.enuCxCDetPorSer
            }
        lobjRep.SGenereReporte()
    End Sub
#End Region
#Region "Estado"
    Private Sub SMuestreEstado()
        If MobjCliente.BlnExiste Then
            Mouse.OverrideCursor = Cursors.Wait
            Dim ldecIntPorCausar = 0D
            Dim lstrIdPredioAgr As String = MobjObjetoWin.ObjIdPredioStr.ToString()
            ldecIntPorCausar = MobjCliente.FdecIntMoraPorCausar({lstrIdPredioAgr},
                    Date.Today)
            MobjEstadoCuenta = MobjCliente.FobjEstadoCtaHoy(lstrIdPredioAgr, MstrServicio,
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
            dgrFacturasEstado.DataContext = MdtbFacturasEstado
            txtFechaEstado.Content = Format(Date.Now, "dd/MM/yyyy hh:mm:ss tt")
            Mouse.OverrideCursor = Cursors.Arrow
        Else
            dgrFacturasEstado.DataContext = Nothing
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
        If MobjCliente.BlnExiste Then
            Dim ldblIdCliente As Double = MobjCliente.ObjIdClienteDbl.ObjValorPro
            Dim lstrIdPredioAgr As String = MobjObjetoWin.ObjIdPredioStr.ToString
            dgrFacturasCliente.DataContext = ClsOrionCop.FdtbFacturas(ldblIdCliente,
                    lstrIdPredioAgr, dtpFechaDesde.SelectedDate, dtpFechaHasta.SelectedDate)
        Else
            dgrFacturasCliente.DataContext = Nothing
        End If
    End Sub
    Private Sub SMuestreRecibos()
        If MobjCliente.BlnExiste Then
            Dim lstrIdPredioAgr As String = MobjObjetoWin.ObjIdPredioStr.ObjValorPro
            dgrRecibosCliente.DataContext = MobjCliente.FdtbRecibos(lstrIdPredioAgr,
                    dtpFechaDesdeRec.SelectedDate, dtpFechaHastaRec.SelectedDate)
        Else
            dgrRecibosCliente.DataContext = Nothing
        End If
    End Sub
    Private Sub SMuestreAnticipos()
        If MobjCliente.BlnExiste Then
            Dim lstrIdPredioAgr As String = MobjObjetoWin.ObjIdPredioStr.ToString
            dgrAnticipos.DataContext = MobjCliente.DtbAnticipos(lstrIdPredioAgr)
        Else
            dgrAnticipos.DataContext = Nothing
        End If
    End Sub
    Private Sub SMuestreNotasCr()
        If MobjCliente.BlnExiste Then
            Dim lstrIdPredioAgr As String = MobjObjetoWin.ObjIdPredioStr.ToString
            dgrNotasCr.DataContext = MobjCliente.FdtbNotasCr(lstrIdPredioAgr)
        Else
            dgrNotasCr.DataContext = Nothing
        End If
    End Sub
    Private Sub SMuestreNotasDb()
        If MobjCliente.BlnExiste Then
            Dim lstrIdPredioAgr As String = MobjObjetoWin.ObjIdPredioStr.ToString
            dgrNotasDbCliente.DataContext = MobjCliente.FdtbNotasDb(lstrIdPredioAgr,
                dtpFechaDesdeMora.SelectedDate, dtpFechaHastaMora.SelectedDate)
        Else
            dgrNotasDbCliente.DataContext = Nothing
        End If
    End Sub
    Private Sub SMuestreMovimiento()
        If MobjCliente.BlnExiste Then
            Dim lstrIdPredioAgr As String = MobjObjetoWin.ObjIdPredioStr.ToString
            dgrMovimiento.DataContext = MobjCliente.FdtbMovimientoDeuda(lstrIdPredioAgr,
                    dtpFechaMovDesde.SelectedDate, dtpFechaMovHasta.SelectedDate)
            SValide()
        Else
            dgrMovimiento.DataContext = Nothing
        End If
    End Sub
    Private Sub SMuestreMovimientoAnt()
        If MobjCliente.BlnExiste Then
            Dim lstrIdPredioAgr As String = MobjObjetoWin.ObjIdPredioStr.ToString
            dgrMovimiento.DataContext = MobjCliente.FdtbMovimientoAnticipos(lstrIdPredioAgr,
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
                If lelmElemento.Name = "txtIdPredio" Then
                    SAbraPredio(txtIdPredio.Text, False)
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
                    Case "cboCliente"
                        SAbraCliente(cboCliente.SelectedItem)
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
    Private Sub TxtIdPredio_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdPredio.KeyDown
        If e.Key = Key.Return Then
            If txtIdPredio.Focus Then
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                    SAbraPredio(txtIdPredio.Text, False)
                End If
            End If
        End If
    End Sub
    Private Sub DtpFecha_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles _
            dtpFechaDesde.SelectedDateChanged, dtpFechaHasta.SelectedDateChanged, dtpFechaDesdeRec.SelectedDateChanged,
            dtpFechaHastaRec.SelectedDateChanged, dtpFechaDesdeMora.SelectedDateChanged,
            dtpFechaHastaMora.SelectedDateChanged, dtpFechaMovDesde.SelectedDateChanged,
            dtpFechaMovHasta.SelectedDateChanged
        If Not HblnCargandoForma Then
            SMuestreDatos()
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
                                    Dim lstrMens = "No se ha seleccionado una Fila que se relacione con un " &
                                        "Documento !"
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
