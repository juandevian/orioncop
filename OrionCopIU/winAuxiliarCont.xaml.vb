Imports System.Collections.ObjectModel
Imports System.ComponentModel
Public Class WinAuxiliarCont
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuFechaDesde
        enuFechaHasta
        enuCtaContIni
        enuCtaContFin
        enuCliente
        enuPreAgr
    End Enum
#End Region
    ' Variables
    Private MblnPoblandoCombo As Boolean = False
    Private MenuReporte As EnuReporteDef = EnuReporteDef.None
    Private MblnBuscandoCuenta As Boolean = False
    Private ReadOnly MobjCliente As New ClsCliente(EnuModoInstanciaObjDef.enuUnico)
    Private MdtmFechaDesde As Date = GobjParametros.ObjAnoActual.DtmFechaInicioAno
    Private MdtmFechaHasta As Date = Date.Today
    Private MstrIdCtaContIni As String = String.Empty
    Private MstrIdCtaContFin As String = String.Empty
    Private MdblCliente As Double = 0.0
    Private MstrIdPredioAgr As String = String.Empty
    Private MblnCalculaSaldo As Boolean = False
    Private MblnConTercero As Boolean = False
    Private MdtbAuxCon As DataTable = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomAuxCont
    Private ReadOnly MobjRep As New ClsRepOrionCop(GCOBJREGISTRO)
#End Region
#Region "Constructor"
    Friend Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuAuxiliarCont
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneCtlsRestringidos()
        SCargueForma(EnuElementosAdicionalesDef.enuImprimir, 6,
                Nothing, Nothing, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SModificarClic()
        If Not FblnExisteExcel() Then
            bttExportar.Visibility = Visibility.Hidden
        End If
        dtpFechaDesde.Focus()
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
        EnuTipoPermisoObjWin = GobjParametros.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuFechaDesde) = lblFechaDesde
        StcValidaControl(EnuValidEntradaDef.enuFechaHasta) = lblFechaHasta
        StcValidaControl(EnuValidEntradaDef.enuCtaContIni) = lblCuentaContIni
        StcValidaControl(EnuValidEntradaDef.enuCtaContFin) = lblCuentaContFin
        StcValidaControl(EnuValidEntradaDef.enuCliente) = lblIdCliente
        StcValidaControl(EnuValidEntradaDef.enuPreAgr) = lblPredioAgr
        SVacie()
        dtpFechaDesde.SelectedDate = GobjParametros.ObjAnoActual.DtmFechaInicioAno
        dtpFechaHasta.SelectedDate = Date.Today
        SAdecueVentana()
        HbttAceptar.TabIndex = 100
        bttExportar.TabIndex = 101
        HbttCancelar.TabIndex = 102
    End Sub
    Private Sub SVacie()
        MobjRep.SVacie()
        dtpFechaDesde.SelectedDate = GobjParametros.ObjAnoActual.DtmFechaInicioAno
        dtpFechaHasta.SelectedDate = Date.Today
        txtCuentaContIni.Text = String.Empty
        txtCuentaContFin.Text = String.Empty
        txtIdCliente.Text = "0"
        txtNomCliente.Content = String.Empty
        txtCuentaContIni.Text = String.Empty
        txtCuentaContFin.Text = String.Empty
        txtNomCuentaContIni.Content = String.Empty
        txtNomCuentaContFin.Content = String.Empty
        MdtmFechaDesde = GCDTMFECHANULA
        MdtmFechaDesde = GCDTMFECHANULA
        MstrIdCtaContIni = String.Empty
        MstrIdCtaContFin = String.Empty
        MblnCalculaSaldo = False
    End Sub
    Protected Overrides Sub SMuestreDatos()
        Title = My.Resources.TituloRep
        Select Case MenuReporte
            Case EnuReporteDef.enuAuxiliar
                Title &= My.Resources.RepAuxiliarCon
        End Select
        SValide()
        If MdblCliente > 0 AndAlso MobjCliente.BlnExiste Then
            SPuebleComboBox()
        End If
    End Sub
    Protected Overrides Sub SValide()
        SValideFechas()
        If StcValidValido(EnuValidEntradaDef.enuFechaDesde) AndAlso
                StcValidValido(EnuValidEntradaDef.enuFechaHasta) Then
            SValideCtasCont()
            If StcValidValido(EnuValidEntradaDef.enuCtaContIni) AndAlso
                    StcValidValido(EnuValidEntradaDef.enuCtaContFin) Then
                SValideCliente()
            End If
            SValidePredioAgr()
        End If
        '
        SHabiliteBotonesTlb()
        If FblnEstanTodosBien() Then
            SHabiliteBotonTlb(True, HbttImprimir)
            SHabiliteMenuItem(True, HmnuImprimir)
        End If
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjRep
            Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaDesde)
            Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaHasta)
            .StrFechaDesde = lstrFechaDesde
            .StrFechaHasta = lstrFechaHasta
            .StrIdCuentaContIni = MstrIdCtaContIni
            .StrIdCuentaContFin = MstrIdCtaContFin
            .StrIdPredioAgru = MstrIdPredioAgr
            .DblIdCliente = MdblCliente
            .BlnCalculaSaldo = MblnCalculaSaldo
            If MdblCliente > 0 Then
                MenuReporte = EnuReporteDef.enuAuxTer
            Else
                MenuReporte = EnuReporteDef.enuAuxiliar
            End If
        End With
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()

        Dim lentPosicion = HmnuAcciones.Items.Count - 2
        Dim lsepSeparad As New Separator
        HmnuImprimir = FmnuiMenuItem("hmnuImprimir", "Imprimir", "RecMnuItemSec")
        HmnuAcciones.Items.Insert(lentPosicion, HmnuImprimir)
        HmnuAcciones.Items.Insert(lentPosicion, lsepSeparad)
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    ''' <summary>
    ''' Prepara la ventana y su objeto para modificar el objeto. Invalida la función "SModifique"
    ''' de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SModifique()
        If ObjObjetoWin.EnuEstadoActualizacion = EnuEstadoObjetoDef.enuConsultando Then
            EnuOperacionEnWin = EnuOperacionEnVentana.cenuModificando
            If Not IsNothing(HbttCancelar) Then
                HbttCancelar.Content = My.Resources.Cancelar
            End If
            SHabiliteWin(True)
            bttExportar.IsEnabled = False
            HmnuImprimir.Style = FindResource("RecMnuItemSecDes")
            HbttImprimir.Style = FindResource("RecBttDesHabilitado")
            SValide()
        Else
            Throw New ErrorInesperadoPanLException("Estado inesperado del objeto!")
        End If
    End Sub
    ''' <summary>
    ''' Invalida el procedimiento "SGuarde" de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SGuarde()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            SRegistre()
            SValide()
            If FblnEstanTodosBien() Then
                Mouse.OverrideCursor = Cursors.Wait
                SEstablezcaOrigenDatos()
                If MdtbAuxCon Is Nothing Then
                    MdtbAuxCon = MobjRep.FdtbAuxiliarCon(True)
                End If
                If MdtbAuxCon.Rows.Count = 0 Then
                    lstrMens = "No hay Datos para mostrar!"
                End If
                SMuestreAuxiliar()
                bttExportar.IsEnabled = True
                HmnuImprimir.Style = FindResource("RecMnuItemSecHab")
                HbttImprimir.Style = FindResource("RecBttHabilitado")
                MobjRep.SVacie()
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
            Mouse.OverrideCursor = Cursors.Arrow
            If lblnNoHayError Then
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Protected Overrides Sub SCancele()
        SCerrarClic()
    End Sub
    Protected Overrides Sub SCerrarClic()
        Me.Close()
    End Sub
    Protected Overrides Sub SImprima()
        Mouse.OverrideCursor = Cursors.Wait
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            SValide()
            If FblnEstanTodosBien() Then
                MobjRep.EnuReporte = MenuReporte
                MobjRep.SGenereReporte()
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
            Mouse.OverrideCursor = Cursors.Arrow
            If Not lblnNoHayError Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
#End Region
#Region "Busqueda"
    Private Sub SBusqueObjeto(abttBoton As Button, ablnCuenta As Boolean)
        MblnBuscandoCuenta = ablnCuenta
        SBuscar()
        If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
            Select Case abttBoton.Name
                Case "bttEncontrarCuentaIni"
                    txtCuentaContIni.Text = StrResultadoBusqueda
                    txtCuentaContIni.Focus()
                Case "bttEncontrarCuentaFin"
                    txtCuentaContFin.Text = StrResultadoBusqueda
                    txtCuentaContFin.Focus()
                Case "bttEncontrarCliente"
                    txtIdCliente.Text = StrResultadoBusqueda
                    txtIdCliente.Focus()
            End Select
        End If
    End Sub
    Protected Overrides Sub SBuscar()
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
    End Sub
    ''' <summary>
    ''' Invalida la funcion "fblnDefinioBusqueda" de la clase base.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        If MblnBuscandoCuenta Then
            SDefineBusquedaCuentaCont()
        Else
            SDefineBusquedaPredioAgr()
            SDefineBusquedaTercero()
        End If
        Return True
    End Function
    Private Sub SDefineBusquedaCuentaCont()
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
    Private Sub SDefineBusquedaTercero()
        Dim lstrCamposMostrar As String() = {ClsIdClienteDbl.SstrNombreCampoBd,
                                            ClsNombreCompletoStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCompletoStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdClienteDbl.SstrNombreCampoBd
        Dim lstrTabla As String = ClsCliente.SstrNombreTabla
        Dim lstrFiltro As String = ClsOrionCop.StrFiltroUbicacion
        HwinBusqueda.SDefinaBusqueda("Nombre Cuenta", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
    Private Sub SDefineBusquedaPredioAgr()
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
#End Region
#Region "Procedimientos Propios"
    Private Sub SAdicioneCtlsRestringidos()
        SAdicioneControlRestringido(dgrAuxiliar)
        SAdicioneControlRestringido(bttEncontrarCliente)
        SAdicioneControlRestringido(bttEncontrarCuentaFin)
        SAdicioneControlRestringido(bttEncontrarCuentaIni)
    End Sub
    Private Sub SAdecueVentana()
        HbttModificar.Visibility = Visibility.Collapsed
        HbttRefrescar.Visibility = Visibility.Collapsed
        HbttAceptar.Content = My.Resources.BtnGenRep
        HbttAceptar.ToolTip = My.Resources.TTGeneraRep
        HbttCancelar.ToolTip = My.Resources.TTCierraVen
    End Sub
    Private Sub SPuebleComboBox()
        MblnPoblandoCombo = True
        cboPredioAgr.Items.Clear()
        If StcValidValido(EnuValidEntradaDef.enuCliente) AndAlso MdblCliente > 0 Then
            Dim lstrPrediosAgr = MobjCliente.FstrPrediosAgruClienteTodos(True)
            Dim i = 0
            If lstrPrediosAgr.Length > 1 Then
                i = lstrPrediosAgr.Length
            End If
            Dim lstrPrediosAgrConFacturas = MobjCliente.FstrPrediosAgruClienteConFacturas(False)
            For Each lstrPreAgrConFacturas As String In lstrPrediosAgrConFacturas
                If Not lstrPrediosAgr.Contains(lstrPreAgrConFacturas) Then
                    ReDim Preserve lstrPrediosAgr(i)
                    lstrPrediosAgr(i) = lstrPreAgrConFacturas
                    i += 1
                End If
            Next
            cboPredioAgr.Items.Add(My.Resources.Todos)
            Array.Sort(lstrPrediosAgr)
            For Each lstrPredAgr In lstrPrediosAgr
                If lstrPredAgr = String.Empty Then
                    lstrPredAgr = GCSTRSINPA
                End If
                cboPredioAgr.Items.Add(lstrPredAgr)
            Next
        End If
        MblnPoblandoCombo = False
        cboPredioAgr.SelectedIndex = 0
    End Sub
    Private Function FblnCambioSetings() As Boolean
        Dim lblnCambio = False
        With MobjRep
            Dim lstrFechaDesde = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaDesde)
            Dim lstrFechaHasta = ClsPanoramaDat.FstrFechaNormalizada(MdtmFechaHasta)
            lblnCambio = lstrFechaDesde <> .StrFechaDesde OrElse
                        lstrFechaHasta <> .StrFechaHasta OrElse
                        MstrIdCtaContIni <> .StrIdCuentaContIni OrElse
                        MstrIdCtaContFin <> .StrIdCuentaContFin OrElse
                        MblnCalculaSaldo <> .BlnCalculaSaldo OrElse
                        MdblCliente <> .DblIdCliente OrElse
                        MstrIdPredioAgr <> .StrIdPredioAgru
        End With
        Return lblnCambio
    End Function
    Private Sub SMuestreAuxiliar()
        Mouse.OverrideCursor = Cursors.Wait
        dgrAuxiliar.DataContext = MdtbAuxCon
        txtCantReg.Content = Format(MdtbAuxCon.Rows.Count, "#,##0")
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub SEstablezcaOrigenDatos()
        If FblnEstanTodosBien() AndAlso FblnCambioSetings() Then
            SRegistre()
            MdtbAuxCon = MobjRep.FdtbAuxiliarCon(True)
        End If
    End Sub
    Private Sub SAbraVentanaDoc(astrDocOrigen As String)
        If Not String.IsNullOrEmpty(astrDocOrigen) Then
            Dim lstrTipoDoc = astrDocOrigen.Split(" ")(0)
            Dim lstrNroDoc = astrDocOrigen.Split(" ")(1)
            Dim lstrPrefDoc = ClsPanorama.FstrPrefijoDcto(lstrNroDoc)
            Dim lentIdDoc = ClsPanorama.FentIdDcto(lstrNroDoc)
            Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil, lstrPrefDoc, lentIdDoc}
            Dim lwinVentana As ClsFormInterface
            If GobjParametros.ObjTipoInterfazByt.ObjValorPro = EnuTipoInterfazDef.EnuPorComprobante Then
                Select Case lstrTipoDoc
                    Case "FAC"
                        Dim lobjFactura As New ClsFactura()
                        lobjFactura.SAbra(lobjValorLlave)
                        lwinVentana = New WinFacturas() With {
                            .WinPadre = Me,
                            .ObjObjetoWin = lobjFactura
                        }
                        lwinVentana.ShowDialog()
                    Case "REC"
                        Dim lobjRecCaja As New ClsReciboCaja()
                        lobjRecCaja.SAbra(lobjValorLlave)
                        lwinVentana = New WinRecibosCaja With {
                            .WinPadre = Me,
                            .ObjObjetoWin = lobjRecCaja
                        }
                        lwinVentana.ShowDialog()
                    Case "NDB"
                        Dim lobjNotaDb As New ClsNotaDb()
                        lobjNotaDb.SAbra(lobjValorLlave)
                        lwinVentana = New WinNotasIntMora With {
                            .WinPadre = Me,
                            .ObjObjetoWin = lobjNotaDb
                        }
                        lwinVentana.ShowDialog()
                    Case "NCR"
                        Dim lobjNotaCr As New ClsNotaCr()
                        lobjNotaCr.SAbra(lobjValorLlave)
                        lwinVentana = New WinNotasCr With {
                            .WinPadre = Me,
                            .ObjObjetoWin = lobjNotaCr
                        }
                        lwinVentana.ShowDialog()
                    Case "NCO"
                        Dim lobjNotaCon As New ClsNotaCon()
                        lobjNotaCon.SAbra(lobjValorLlave)
                        lwinVentana = New WinNotasAplicaAnt With {
                            .WinPadre = Me,
                            .ObjObjetoWin = lobjNotaCon
                        }
                        lwinVentana.ShowDialog()
                    Case "NDA"

                    Case "NRRC"
                End Select
            ElseIf GobjParametros.ObjTipoInterfazByt.ObjValorPro = EnuTipoInterfazDef.EnuPorDocumento Then
                For Each lobjDoc As ClsDocumento In GobjParametros.ColDocumentos
                    If lobjDoc.ObjTipoDocumentoStr.ObjValorPro = lstrTipoDoc Then
                        Select Case lobjDoc.ObjIdDocumentoEnt.ObjValorPro
                            Case EnuIdDocumentoDef.EnuFacturaVenta
                                Dim lobjFactura As New ClsFactura()
                                lobjFactura.SAbra(lobjValorLlave)
                                lwinVentana = New WinFacturas() With {
                                    .WinPadre = Me,
                                    .ObjObjetoWin = lobjFactura
                                }
                                lwinVentana.ShowDialog()
                            Case EnuIdDocumentoDef.EnuReciboCaja
                                Dim lobjRecCaja As New ClsReciboCaja()
                                lobjRecCaja.SAbra(lobjValorLlave)
                                lwinVentana = New WinRecibosCaja With {
                                    .WinPadre = Me,
                                    .ObjObjetoWin = lobjRecCaja
                                }
                                lwinVentana.ShowDialog()
                            Case EnuIdDocumentoDef.EnuNotaAjuste
                            Case EnuIdDocumentoDef.EnuNotaAplicacionAnt
                                Dim lobjNotaCon As New ClsNotaCon()
                                lobjNotaCon.SAbra(lobjValorLlave)
                                lwinVentana = New WinNotasAplicaAnt With {
                                    .WinPadre = Me,
                                    .ObjObjetoWin = lobjNotaCon
                                }
                                lwinVentana.ShowDialog()
                            Case EnuIdDocumentoDef.EnuNotaCr
                                Dim lobjNotaCr As New ClsNotaCr()
                                lobjNotaCr.SAbra(lobjValorLlave)
                                lwinVentana = New WinNotasCr With {
                                    .WinPadre = Me,
                                    .ObjObjetoWin = lobjNotaCr
                                }
                                lwinVentana.ShowDialog()
                            Case EnuIdDocumentoDef.EnuNotaIntMora
                                Dim lobjNotaDb As New ClsNotaDb()
                                lobjNotaDb.SAbra(lobjValorLlave)
                                lwinVentana = New WinNotasIntMora With {
                                    .WinPadre = Me,
                                    .ObjObjetoWin = lobjNotaDb
                                }
                                lwinVentana.ShowDialog()
                            Case EnuIdDocumentoDef.EnuNotaReintegroAnt
                            Case EnuIdDocumentoDef.EnuNotaReversaCr
                        End Select
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub
    Private Sub SExporteAExcel()
        Mouse.OverrideCursor = Cursors.Wait
        SValide()
        If FblnEstanTodosBien() Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Dim lblnExportoExcel As Boolean
            Try
                Dim lobjExp As New ClsExportToExcell
                Dim lstrFecha = Today.Year.ToString & Format(Today.Month, "00") &
                        Format(Today.Day, "00")
                Dim lstrNomArchivo = lstrFecha & "_Aux_" & MstrIdCtaContIni
                If MstrIdCtaContFin <> MstrIdCtaContIni Then
                    lstrNomArchivo &= "_" & MstrIdCtaContFin
                End If
                If lstrNomArchivo.Length > 31 Then
                    lstrNomArchivo = lstrNomArchivo.Substring(0, 31)
                End If
                lstrNomArchivo = GstrTrayReportes & "\" & lstrNomArchivo
                lblnExportoExcel = lobjExp.FblnExportToExcel("Auxiliar Contable", lstrNomArchivo,
                        dgrAuxiliar, MdtbAuxCon, FstrMapeoColsRepAuxiliar)
                If lblnExportoExcel Then
                    If MsgBox("Desea abrir el Archivo de Excel?", MsgBoxStyle.Question + vbYesNo,
                        "Abrir Archivo") = MsgBoxResult.Yes Then
                        lobjExp.SAbraExcel()
                    End If
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
                Mouse.OverrideCursor = Cursors.Arrow
                If Not lblnNoHayError Then
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub
    Private Sub SLimpieDataGrid()
        If FblnCambioSetings() Then
            MdtbAuxCon = Nothing
            txtCantReg.Content = String.Empty
            dgrAuxiliar.DataContext = MdtbAuxCon
            bttExportar.IsEnabled = False
            HmnuImprimir.Style = FindResource("RecMnuItemSecDes")
            HbttImprimir.Style = FindResource("RecBttDesHabilitado")
        End If
    End Sub
    Private Function FstrMapeoColsRepAuxiliar() As String()
        Dim lstrMap(dgrAuxiliar.Columns.Count - 1) As String
        lstrMap(0) = "Fecha=FechaNovedad"
        lstrMap(1) = "Cuenta Cont=IdCta"
        lstrMap(2) = "Tercero=IdTerceroCliente"
        lstrMap(3) = "Alias Contable=AliasCont"
        lstrMap(4) = "Predio Agrupador=IdPredioAgrupador"
        lstrMap(5) = "Detalle=Detalle"
        lstrMap(6) = "Doc. Origen=NroDocOrigen"
        lstrMap(7) = "Debitos=Debito"
        lstrMap(8) = "Creditos=Credito"
        lstrMap(9) = "Saldo=Saldo"
        Return lstrMap
    End Function
#End Region
#Region "Validacion"
    Private Sub SValideFechas()
        Dim lstrMens = String.Empty
        Dim lblnEsValidaFechaDesde = (MdtmFechaDesde >= GCDTMFECHANULA AndAlso
                                    MdtmFechaDesde <= Date.Today)
        Dim lblnEsValidaFechaHasta = False
        If lblnEsValidaFechaDesde Then
            lblnEsValidaFechaHasta = (MdtmFechaHasta >= MdtmFechaDesde AndAlso
                                      MdtmFechaHasta <= Date.Today)
        End If
        StcValidValido(EnuValidEntradaDef.enuFechaDesde) = lblnEsValidaFechaDesde
        StcValidValido(EnuValidEntradaDef.enuFechaHasta) = lblnEsValidaFechaHasta
        If Not lblnEsValidaFechaDesde Then
            lstrMens = "La Fecha inicial no es válida!"
        ElseIf Not lblnEsValidaFechaHasta Then
            lstrMens = "La Fecha final no es válida!"
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        Else
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuOk)
        End If
    End Sub
    Private Sub SValideCtasCont()
        Dim lblnEsValidaCuentaIni = ClsPanorama.FblnEsValidoString(MstrIdCtaContIni, 4, 30, True)
        Dim lblnEsValidaCuentaFin = False
        Dim lstrMens = String.Empty
        If lblnEsValidaCuentaIni Then
            lblnEsValidaCuentaIni = ClsOrionCop.FblnEsValidaCtaContabilidad(MstrIdCtaContIni)
            If lblnEsValidaCuentaIni Then
                txtNomCuentaContIni.Content = ClsOrionCop.FstrNombreCuentaCon(MstrIdCtaContIni)
            Else
                txtNomCuentaContIni.Content = String.Empty
            End If
        End If
        If lblnEsValidaCuentaIni Then
            If Not String.IsNullOrEmpty(MstrIdCtaContFin) Then
                lblnEsValidaCuentaFin = ClsPanorama.FblnEsValidoString(MstrIdCtaContFin, 4, 30, True)
                If lblnEsValidaCuentaFin Then
                    lblnEsValidaCuentaFin = ClsOrionCop.FblnEsValidaCtaContabilidad(MstrIdCtaContFin)
                    If lblnEsValidaCuentaFin Then
                        txtNomCuentaContFin.Content = ClsOrionCop.FstrNombreCuentaCon(MstrIdCtaContFin)
                    Else
                        txtNomCuentaContFin.Content = String.Empty
                    End If
                End If
            Else
                txtCuentaContFin.Text = MstrIdCtaContIni
                lblnEsValidaCuentaFin = True
            End If
        End If
        If lblnEsValidaCuentaFin Then
            lblnEsValidaCuentaFin = (MstrIdCtaContFin >= MstrIdCtaContIni)
        End If
        txtNomCuentaContIni.ToolTip = txtNomCuentaContIni.Content
        txtNomCuentaContFin.ToolTip = txtNomCuentaContFin.Content
        StcValidValido(EnuValidEntradaDef.enuCtaContIni) = lblnEsValidaCuentaIni
        StcValidValido(EnuValidEntradaDef.enuCtaContFin) = lblnEsValidaCuentaFin
        If Not lblnEsValidaCuentaIni Then
            lstrMens = "La Cuenta inicial no es válida!"
        ElseIf Not lblnEsValidaCuentaFin Then
            lstrMens = "La Cuenta final no es válida!"
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SValideCliente()
        Dim lblnEsValidoCliente = ClsPanorama.FblnEsValidoNumero(MdblCliente, 0,
                GCDBLMAXTERC, True, EnuTipoValor.enuDouble)
        If lblnEsValidoCliente AndAlso MdblCliente > 0 Then
            lblnEsValidoCliente = FblnMostroCli()
            MblnConTercero = lblnEsValidoCliente
        Else
            MblnConTercero = False
        End If
        StcValidValido(EnuValidEntradaDef.enuCliente) = lblnEsValidoCliente
        If Not lblnEsValidoCliente Then
            SLevanteEveNoti("La Id. del Cliente inicial no es valida!'", "", 0,
                    EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Function FblnMostroCli() As Boolean
        Dim lblnMostro = True
        Dim lobjValorLlave As Object()
        lobjValorLlave = {GshrIdCarpeta, GshrIdCentroUtil, MdblCliente}
        MobjCliente.SAbra(lobjValorLlave)
        If MobjCliente.BlnExiste Then
            txtNomCliente.Content = MobjCliente.ObjNombreCompletoStr.ObjValorPro
        Else
            lblnMostro = False
        End If
        Return lblnMostro
    End Function
    Private Sub SValidePredioAgr()
        Dim lblnEsValidoPredioAgr = True, lstrMens As String
        If MblnConTercero Then
            lblnEsValidoPredioAgr = Not IsNothing(cboPredioAgr.SelectedItem) AndAlso cboPredioAgr.SelectedIndex >= 0
        End If
        StcValidValido(EnuValidEntradaDef.enuPreAgr) = lblnEsValidoPredioAgr
        If Not lblnEsValidoPredioAgr Then
            lstrMens = "El Predio agrupador no es válido!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Select Case lelmElemento.Name
                Case "bttEncontrarCuentaIni", "bttEncontrarCuentaFin"
                    SBusqueObjeto(lelmElemento, True)
                Case "bttEncontrarCliente"
                    SBusqueObjeto(lelmElemento, False)
                Case "bttExportar"
                    SExporteAExcel()
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
            If TypeOf lelmElemento Is DatePicker OrElse TypeOf lelmElemento Is TextBox Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                    Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                    Try
                        Select Case lelmElemento.Name
                            Case "dtpFechaDesde"
                                MdtmFechaDesde = dtpFechaDesde.SelectedDate
                            Case "dtpFechaHasta"
                                MdtmFechaHasta = dtpFechaHasta.SelectedDate
                            Case "txtCuentaContIni"
                                MstrIdCtaContIni = txtCuentaContIni.Text
                            Case "txtCuentaContFin"
                                MstrIdCtaContFin = txtCuentaContFin.Text
                            Case "txtIdCliente"
                                If IsNumeric(txtIdCliente.Text) Then
                                    MdblCliente = CType(txtIdCliente.Text, Double)
                                Else
                                    MdblCliente = 0
                                End If
                        End Select
                        SMuestreDatos()
                        SLimpieDataGrid()
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
    Private Sub CboPredioAgr_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cboPredioAgr.SelectionChanged
        If Not MblnPoblandoCombo Then
            If Not IsNothing(cboPredioAgr.SelectedIndex) Then
                MstrIdPredioAgr = cboPredioAgr.SelectedItem
                If MstrIdPredioAgr = GCSTRSINPA Then
                    MstrIdPredioAgr = String.Empty
                End If
            Else
                MstrIdPredioAgr = String.Empty
            End If
            SValide()
        End If
        SLimpieDataGrid()
    End Sub
    Private Sub ChkCalcularSaldo_Click(sender As Object, e As RoutedEventArgs) Handles chkCalcularSaldo.Click
        MblnCalculaSaldo = chkCalcularSaldo.IsChecked
        SMuestreDatos()
        SLimpieDataGrid()
    End Sub
    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
            dgrAuxiliar.MouseRightButtonUp
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If lelmElemento.Name = "dgrAuxiliar" Then
            Dim lstrDocOrigen As String
            If Not IsNothing(dgrAuxiliar.SelectedItem) Then
                Dim ldrvNovedad As DataRowView = dgrAuxiliar.SelectedItem
                lstrDocOrigen = ldrvNovedad("NroDocOrigen")
                SAbraVentanaDoc(lstrDocOrigen)
            End If
        End If
    End Sub
#End Region
End Class