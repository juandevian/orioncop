Public Class WinIntegridad
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
    Private Enum EnuValidEntradaDef As Integer
        enuFecIni
        enuFecFin
    End Enum
#End Region
    ' Variables
    Private ReadOnly MstrNombreVentana As String = "VERIFICA INTEGRIDAD"
    Private MdtmFechaIni As Date = GCDTMFECHANULA
    Private MdtmFechaFin As Date = GCDTMFECHANULA
    Private MblnCancelando As Boolean = False
    Private MenuDocsInt As EnuDocsIntegridad = EnuDocsIntegridad.None
    Private WithEvents MobjOrionCop As New ClsOrionCop(GCOBJREGISTRO, False)
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuVerificaInt
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            SCargueForma(EnuElementosAdicionalesDef.None, 2,
                Nothing, Nothing, True)
            SPuebleBarraEstado(HcolLabelsBarraEstado)
            SCree()
            HbttCancelar.Content = My.Resources.CerrarBtn
            dtpFechaDesde.SelectedDate = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
            dtpFechaHasta.SelectedDate = Date.Today
            dtpFechaDesde.Focus()
            lblnNoHayError = True
        Catch ex As ErrorInesperadoPanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
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
            ObjObjetoWin = New ClsFactura()
            Dim lobjOBjetoWin As ClsFactura = ObjObjetoWin
            EnuTipoPermisoObjWin = lobjOBjetoWin.EnuPermisosObj
        End If
    End Sub

    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuFecIni) = lblFechaDesde
        StcValidaControl(EnuValidEntradaDef.enuFecFin) = lblFechaHasta
        '
        MdgtLblActualiza = New SdgtActualizaLabel(AddressOf lblAvance.SetValue)
        MdgtPgbActualiza = New SdgtActualizaProgressBar(AddressOf pgbAvance.SetValue)
        '
        HbttAceptar.TabIndex = 20
        HbttCancelar.TabIndex = 21
    End Sub

    Protected Overrides Sub SMuestreDatos()
        ' 
    End Sub
    Protected Overrides Sub SValide()
        If Not HblnCargandoForma Then
            StcValidValido(EnuValidEntradaDef.enuFecFin) = (MdtmFechaFin <= Date.Today AndAlso
                        MdtmFechaFin >= MdtmFechaIni)
            StcValidValido(EnuValidEntradaDef.enuFecIni) = (MdtmFechaIni > GCDTMFECHANULA AndAlso
                    MdtmFechaIni <= Date.Today)
            '
            SHabiliteBotonesTlb()
            FblnEstanTodosBien()
        End If
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
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        If EnuDocInt = EnuDocsIntegridad.None Then
            lstrMens = "No hay Documentos selecciobados para ser verificados!"
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Exit Sub
        End If
        HbttCancelar.Content = My.Resources.Cancelar
        lstrMens = "Verificando Integridad"
        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        pgbAvance.Visibility = Visibility.Visible
        lblAvance.Visibility = Visibility.Visible
        Dispatcher.Invoke(MdgtLblActualiza,
                        System.Windows.Threading.DispatcherPriority.Background,
                        New Object() {Label.ContentProperty, "Verificando Integridad"})
        Dim lblnHayErrorInt = False
        Dim lstrArchIntegridad As String = String.Empty
        Cursor = Cursors.Wait
        HbttAceptar.IsEnabled = False
        Try
            lstrArchIntegridad = MobjOrionCop.SCompruebeIntegridad(lblnHayErrorInt,
                    dtpFechaDesde.SelectedDate, dtpFechaHasta.SelectedDate, EnuDocInt)
            If Not lblnHayErrorInt Then
                If MblnCancelando Then
                    lstrMens = My.Resources.ProcesoCanc
                    MblnCancelando = False
                Else
                    lstrMens = "No se detectarón Problemas de Integridad!"
                    txtProceso.Content = My.Resources.ProTer
                    pgbAvance.Visibility = Visibility.Collapsed
                    lblAvance.Visibility = Visibility.Collapsed
                End If
            Else
                lstrMens = "Hay Problemas de Integridad! Por favor informe a Soporte!"
                Process.Start("notepad.exe", lstrArchIntegridad)
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
            If lblnNoHayError Then
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            HbttAceptar.IsEnabled = True
            HbttCancelar.IsEnabled = True
            HbttCancelar.Content = My.Resources.CerrarBtn
            Cursor = Cursors.Arrow
        End Try
    End Sub
    Protected Overrides Sub SCancele()
        If MobjOrionCop.BlnCorriendoIntegridad Then
            MblnCancelando = True
        Else
            If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                SCerrarClic()
            Else
                SFinaliceOperacion()
                SCerrarClic()
            End If
        End If
    End Sub
#End Region
#Region "Eventos del proceso"
    Private Sub SEvnInicio(aobjSender As Object, e As ClsPanEventArgs) Handles _
            MobjOrionCop.EvnInicio
        Select Case e.EnuProceso
            Case EnuProcesoDef.enuIntegrFac
                txtProceso.Content = "Procesando Facturas"
            Case EnuProcesoDef.enuIntegrRec
                txtProceso.Content = "Procesando Recibos de Caja"
            Case EnuProcesoDef.enuIntegrNcr
                txtProceso.Content = "Procesando Notas Crédito"
            Case EnuProcesoDef.enuIntegrNco
                txtProceso.Content = "Procesando Notas Aplicación Anticipos"
            Case EnuProcesoDef.enuIntegrNdb
                txtProceso.Content = "Procesando Notas por Intereses de Mora"
            Case EnuProcesoDef.enuInteAnt
                txtProceso.Content = "Procesando Anticipos"
            Case EnuProcesoDef.enuInteNRRC
                txtProceso.Content = "Procesando Notas Reversión R.C."
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
        Dim lstrResultado = My.Resources.EleProce & Format(e.DblCantProcesada, "##0") &
                        My.Resources.De & Format(e.DblCantAProcesar, "##0")
        Dispatcher.Invoke(MdgtPgbActualiza,
                System.Windows.Threading.DispatcherPriority.Background,
                New Object() {ProgressBar.ValueProperty, e.DblCantProcesada})
        Dispatcher.Invoke(MdgtLblActualiza,
                System.Windows.Threading.DispatcherPriority.Background,
                New Object() {Label.ContentProperty, lstrResultado})
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private ReadOnly Property EnuDocInt() As EnuDocsIntegridad
        Get
            MenuDocsInt = EnuDocsIntegridad.None
            If chkFac.IsChecked Then MenuDocsInt += EnuDocsIntegridad.enuFac
            If chkRec.IsChecked Then MenuDocsInt += EnuDocsIntegridad.enuRC
            If chkNCr.IsChecked Then MenuDocsInt += EnuDocsIntegridad.enuNcr
            If chkNdb.IsChecked Then MenuDocsInt += EnuDocsIntegridad.enuNdb
            If chkNco.IsChecked Then MenuDocsInt += EnuDocsIntegridad.enuNco
            If chkAnt.IsChecked Then MenuDocsInt += EnuDocsIntegridad.enuAnt
            If chkNRRC.IsChecked Then MenuDocsInt += EnuDocsIntegridad.enuNrrc
            If chkEstadosCta.IsChecked Then MenuDocsInt += EnuDocsIntegridad.enuEstadoCta
            If chkTodos.IsChecked Then MenuDocsInt = EnuDocsIntegridad.enuTodos
            Return MenuDocsInt
        End Get
    End Property
    Private Sub SCorraAuto()
        Dim lobjPerAnt As ClsPeriodo = ClsOrionCop.FobjPeriodoAnterior()
        Dim ldtmFechaIni = lobjPerAnt.DtmFechaInicioPeriodo
        dtpFechaDesde.SelectedDate = ldtmFechaIni
        dtpFechaHasta.SelectedDate = Date.Today
        chkTodos.IsChecked = True
        Cursor = Cursors.Wait
        SGuarde()
        Cursor = Cursors.Arrow
    End Sub
#End Region
#Region "Eventos en la Ventana"
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
                    Select Case lelmElemento.Name
                        Case "dtpFechaDesde"
                            MdtmFechaIni = dtpFechaDesde.SelectedDate
                        Case "dtpFechaHasta"
                            MdtmFechaFin = dtpFechaHasta.SelectedDate
                    End Select
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub

    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is CheckBox Then
            If lelmElemento.Name = "chkTodos" Then
                chkAnt.IsChecked = False
                chkFac.IsChecked = False
                chkNco.IsChecked = False
                chkNCr.IsChecked = False
                chkNdb.IsChecked = False
                chkRec.IsChecked = False
                chkNRRC.IsChecked = False
                chkEstadosCta.IsChecked = False
            Else
                chkTodos.IsChecked = False
            End If
        End If
    End Sub

    Private Sub Dtp_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles _
            dtpFechaDesde.SelectedDateChanged, dtpFechaHasta.SelectedDateChanged
        If sender.Equals(dtpFechaDesde) Then
            MdtmFechaIni = dtpFechaDesde.SelectedDate
        ElseIf sender.Equals(dtpFechaHasta) Then
            MdtmFechaFin = dtpFechaHasta.SelectedDate
        End If
        SValide()
    End Sub
#End Region
End Class
