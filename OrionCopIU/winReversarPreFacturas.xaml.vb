Public Class WinReversarPreFacturas
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
    Private WithEvents MobjOrionCop As New ClsOrionCop(GCOBJREGISTRO, False)
    Private MstrNombreVentana As String = My.Resources.NomRevPreF
    Private MstrResultado As String = String.Empty
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuRevPrefacturas
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 0,
                Nothing, Nothing, True)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SAdapteEtiquetas()
        SCrearClic()
    End Sub
    Protected Overrides ReadOnly Property StrNombreVentana As String
        Get
            If Not GobjParametros.BlnEFacAutorizado Then
                MstrNombreVentana = "REVERSAR PRE-CUENTAS DE COBRO"
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
            Dim lobjFact As ClsFactura = ObjObjetoWin
            EnuTipoPermisoObjWin = lobjFact.EnuPermisosObj
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
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        FblnEstanTodosBien()
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
    Protected Overrides Sub SCree()
        EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando
        Dim ldtmfecha = ClsOrionCop.DtmFechaFacturasAReversar
        txtFechaFacturas.Content = Format(ldtmfecha, GCSTRFMTFECHASIMPLE)
    End Sub
    Protected Overrides Sub SGuarde()
        Dim lstrMens As String
        If GobjParametros.BlnEFacAutorizado Then
            lstrMens = "Reversando Pre-Facturas!!"
        Else
            lstrMens = "Reversando Pre-Cuentas de Cobro!!"
        End If
        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        lstrMens = String.Empty
        Dim lstrMensEx = String.Empty, lblnNoHayError = False
        HbttAceptar.IsEnabled = False
        HbttCancelar.IsEnabled = False
        Try
            Mouse.OverrideCursor = Cursors.Wait
            MobjOrionCop.SReversePreFacturas()
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
            SFinaliceOperacion()
            If lblnNoHayError Then
                lstrMens = "Las Pre-Facturas fueron reversadas exitosamente!"
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                txtProceso.Content = My.Resources.ProTer
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            HbttAceptar.IsEnabled = True
            HbttCancelar.IsEnabled = True
            Mouse.OverrideCursor = Cursors.Arrow
        End Try
    End Sub
    Protected Overrides Sub SCancele()
        If Not ClsOrionCop.BlnPreFacturando Then
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
            Title = "Reversar Pre-Cuentas de Cobro"
            lblTitulo.Content = "Reversar Pre-Cuentas de Cobro..."
            lblPeriodo.Content = "Periodo de las Pre-Cuentas"
            lblFecha.Content = "Fecha de las Pre-Cuentas"
        End If
    End Sub
#End Region
#Region "Eventos Prefacturas"
    Private Sub EInicio(aobjSender As Object, e As ClsPanEventArgs) Handles MobjOrionCop.EvnInicio
        If e.EnuProceso = EnuProcesoDef.enuRevPreFac Then
            txtProceso.Content = My.Resources.RevPreFac
        End If
        pgbAvance.Minimum = 0
        pgbAvance.Maximum = e.DblCantAProcesar
        pgbAvance.Value = 0.0
    End Sub
    Private Sub SEvnAvance(aobjSender As Object, e As ClsPanEventArgs) Handles MobjOrionCop.EvnAvance
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
#Region "Eventos en la Ventana"
    '
#End Region
End Class
