Public Class WinGenFacturas
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Delegados"
    Private Delegate Sub SdgtActualizaProgressBar(dp As _
                 System.Windows.DependencyProperty,
                 value As Object)
    Private Delegate Sub SdgtActualizaLabel(dp As _
                 System.Windows.DependencyProperty,
                 Content As Object)
    Private MdgtPgbActualiza As SdgtActualizaProgressBar = Nothing
    Private MdgtLblActualiza As SdgtActualizaLabel = Nothing
#End Region
#Region "Enumeradores"
#End Region
    Private WithEvents MobjOrionCop As New ClsOrionCop(GCOBJREGISTRO, False)
    Private MstrResultado As String = String.Empty
    Private MblnCancelando As Boolean = False
    Private MblnGuardo As Boolean = False
    Private MstrNombreVentana As String = My.Resources.NomGenFras
    Private ReadOnly MwinMW As MWOrionCop = Nothing
    Private ReadOnly MobjRep As New ClsRepOrionCop(GCOBJREGISTRO)
#End Region

#Region "Constructor"
    Public Sub New(awinMW As MWOrionCop)
        MwinMW = awinMW
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuGenFacturas
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 0,
                Nothing, Nothing, True)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SAdapteEtiquetas()
        SCrearClic()
        If GobjParametros.BlnEFacAutorizado Then
            Dim lstrMens = String.Empty
            FblnEstaConectado(GobjParametros.ObjURLStr.ObjValorPro, lstrMens)
            If Not String.IsNullOrEmpty(lstrMens) Then
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuAdvertencia)
                HbttAceptar.IsEnabled = False
            End If
        End If
    End Sub

    Protected Overrides ReadOnly Property StrNombreVentana As String
        Get
            If Not GobjParametros.BlnEFacAutorizado Then
                MstrNombreVentana = "GENERAR CUENTAS DE COBRO"
            End If
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
            ObjObjetoWin = New ClsFactura()
            Dim lobjObjetoWin As ClsFactura = ObjObjetoWin
            EnuTipoPermisoObjWin = lobjObjetoWin.EnuPermisosObj
        End If
    End Sub

    Protected Overrides Sub SInicialiceControles()
        MdgtPgbActualiza = New SdgtActualizaProgressBar(AddressOf pgbAvance.SetValue)
        MdgtLblActualiza = New SdgtActualizaLabel(AddressOf lblAvance.SetValue)
        '
        HbttAceptar.TabIndex = 2
        HbttCancelar.TabIndex = 3
    End Sub

    Protected Overrides Sub SMuestreDatos()
        txtPeriodoFact.Content = GobjParametros.ObjAnoActual.StrNombrePeriodoActual.ToUpper
    End Sub

    Protected Overrides Sub SValide()
        SHabiliteBotonesTlb()
    End Sub

    Protected Overrides Sub SRegistre()
        '
    End Sub

    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    Protected Overrides Sub SGuarde()
        Dim lblnAjustoCuotaAdmin = False, lblnFacturoCompleto = True
        Cursor = Cursors.Wait
        MblnCancelando = False
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            ' lstrMens devuelve un mensaje de advertencia cuando no se causó mora por cuenta
            ' suspendida o perdida
            If GobjParametros.BlnEFacAutorizado Then
                lstrMens = "Generando Facturas definitivas!"
            Else
                lstrMens = "Generando Cuentas de Cobro definitivas!"
            End If
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            lstrMens = String.Empty
            MblnGuardo = MobjOrionCop.FblnGeneroFacturasDef(lstrMens, lblnFacturoCompleto)
            If Not MblnCancelando Then
                txtProceso.Content = My.Resources.TerPro
                HbttCancelar.IsEnabled = False
                HbttAceptar.IsEnabled = False
                STermineProcesoFact(lstrMens)
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Cursor = Cursors.Wait
            End If
            lblnNoHayError = True
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            SFinaliceOperacion()
            If lblnNoHayError Then
                lstrMens = String.Empty
                If MblnCancelando Then
                    lstrMens = "El Proceso fue cancelado por el Usuario!"
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Else
                    If Not lblnFacturoCompleto Then
                        lstrMens = "Se facturó hasta el último Número habilitado!"
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuAdvertencia)
                    Else
                        If String.IsNullOrEmpty(lstrMens) Then
                            lstrMens = "Proceso terminado exitosamente!"
                            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                        Else
                            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuAdvertencia)
                        End If
                    End If
                End If
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            Activate()
        End Try
    End Sub

    Protected Overrides Sub SCancele()
        If ClsOrionCop.BlnGenFacturas Then
            MblnCancelando = True
        Else
            SCerrarClic()
        End If
    End Sub

    Protected Overrides Sub SRefresqueWin()
        SHabiliteWin(False)
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SAdapteEtiquetas()
        If Not GobjParametros.BlnEFacAutorizado Then
            Title = "Generar Cuentas de Cobro definitivas"
            lblTitulo.Content = "Generar Cuentas de Cobro Definitivas..."
            lblPeriodo.Content = "Periodo de las Cuentas de Cobro"
            lblAvance.Content = "Cuentas de Cobro Procesadas: 0 de 0"
        End If
    End Sub

    Private Sub STermineProcesoFact(ByRef astrMens As String)
        If MblnGuardo Then
            Cursor = Cursors.Arrow
            If Not MblnCancelando Then
                If GobjParametros.BlnEFacAutorizado Then
                    BlnFactAuto = True
                    SProceseEFac(astrMens)
                Else
                    SImprimaFactAut(False)
                End If
            End If
        End If
        txtProceso.Content = My.Resources.ProTer
        HbttCancelar.IsEnabled = True
        HbttAceptar.IsEnabled = True
    End Sub
#End Region

#Region "Eventos Facturacion"
    Private Sub SEvnInicio(aobjSender As Object, e As ClsPanEventArgs) Handles _
            MobjOrionCop.EvnInicio
        Select Case e.EnuProceso
            Case EnuProcesoDef.enuPasandoAFact
                txtProceso.Content = My.Resources.GenFacturasDef
            Case EnuProcesoDef.enuGenEstadosCtaCli
                txtProceso.Content = My.Resources.GenEstadosCtasCli
            Case EnuProcesoDef.enuGenEstadosCtaPre
                txtProceso.Content = My.Resources.GenEstadosCtasPre
            Case EnuProcesoDef.enuApliAnti
                txtProceso.Content = My.Resources.AplAnticipos
            Case EnuProcesoDef.enuCausaMora
                txtProceso.Content = My.Resources.GenInte
            Case EnuProcesoDef.None
                txtProceso.Content = My.Resources.ProcesoCanc
        End Select
        pgbAvance.Minimum = 0.0
        pgbAvance.Maximum = e.DblCantAProcesar
        pgbAvance.Value = 0.0
    End Sub
    Private Sub SEvnAvance(aobjSender As Object, e As ClsPanEventArgs) Handles MobjOrionCop.EvnAvance
        If MblnCancelando Then
            e.BlnCancele = True
            e.EnuProceso = EnuProcesoDef.None
            SEvnInicio(Nothing, e)
            Exit Sub
        End If
        MstrResultado = My.Resources.EleProce & Format(e.DblCantProcesada, "##0") &
                My.Resources.De & Format(e.DblCantAProcesar, "##0")
        Dispatcher.Invoke(MdgtPgbActualiza,
                System.Windows.Threading.DispatcherPriority.Background,
                New Object() {ProgressBar.ValueProperty, e.DblCantProcesada})
        Dispatcher.Invoke(MdgtLblActualiza,
                System.Windows.Threading.DispatcherPriority.Background,
                New Object() {Label.ContentProperty, MstrResultado})
    End Sub
#End Region
End Class
