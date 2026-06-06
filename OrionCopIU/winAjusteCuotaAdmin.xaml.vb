Public Class WinAjusteCuotaAdmin
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
        enuCantidadCuo = 0
    End Enum
#End Region
    ' Variables
    Private MobjObjetoAno As ClsAno = Nothing
    Private WithEvents MobjOrionCop As New ClsOrionCop(GCOBJREGISTRO, False)
    Private MblnCancelando As Boolean = False
    Private MstrResultado As String = String.Empty
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomAjuCtaAdm
    Private MblnAjustando As Boolean = False
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuAjusteCuotaAdmin
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 1,
                Nothing, Nothing, True)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SCree()
        txtCantCuotas.Text = "1"
        txtCantCuotas.Focus()
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
        MobjObjetoAno = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoAno.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuCantidadCuo) = lblCantidadCuotas
        chkNoGenerarAjuste.IsChecked = False
        lblAdvertencia.Visibility = Visibility.Hidden
        txtAdvertencia.Visibility = Visibility.Hidden
        txtCantCuotas.Text = My.Resources.Cero

        MdgtPgbActualiza = New SdgtActualizaProgressBar(AddressOf pgbAvance.SetValue)
        MdgtLblActualiza = New SdgtActualizaLabel(AddressOf lblAvance.SetValue)

        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        StcValidValido(EnuValidEntradaDef.enuCantidadCuo) = FblnCantidadValida()
        '
        SHabiliteBotonesTlb()
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
        '
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    ''' <summary>
    ''' Sub que prepara a la ventana y a su objeto para crear un nuevo objeto. Invalida el Sub
    ''' "SCree" de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SCree()
        EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando
        SFrmAdicione()
    End Sub
    ''' <summary>
    ''' Invalida el procedimiento "SGuarde" de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SGuarde()
        Dim lblnGuardo = False, lblnNoHayError = False
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty
        Dim lenuSeveNoti As EnuSeveridadNot
        MblnCancelando = False
        Dim lentCanCuo = CType(txtCantCuotas.Text, Integer)
        Try
            MblnAjustando = True
            MobjOrionCop.SAjusteCuotas(MobjObjetoAno, lentCanCuo, chkNoGenerarAjuste.IsChecked)
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
                lblnGuardo = True
                If Not MblnCancelando Then
                    lstrMens = "La generación del Retroactivo terminó exitosamente!"
                    txtProceso.Content = My.Resources.ProTer
                Else
                    lstrMens = "La generación del Retroactivo fue Cancelado por el Usuario!"
                    MblnCancelando = False
                End If
                lenuSeveNoti = EnuSeveridadNot.EnuInformacion
            Else
                lenuSeveNoti = EnuSeveridadNot.EnuExcep
            End If
            SLevanteEveNoti(lstrMens, lstrMensEx, 0, lenuSeveNoti)
            MblnAjustando = False
        End Try
    End Sub
    Protected Overrides Sub SFinaliceOperacion()
        SEstablezcaWinConsultando()
        MblnAjustando = False
    End Sub
    Protected Overrides Sub SCancele()
        If MblnAjustando Then
            MblnCancelando = True
        Else
            SCerrarClic()
        End If
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
            If TypeOf lelmElemento Is TextBox Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    Select Case lelmElemento.Name
                        Case "txtCantCuotas"
                            SMuestreDatos()
                    End Select
                End If
            End If
        End If
    End Sub
    Private Sub ChkNoGenerarAjuste_Click(sender As Object, e As RoutedEventArgs) Handles chkNoGenerarAjuste.Click
        If chkNoGenerarAjuste.IsChecked Then
            txtCantCuotas.Style = FindResource("RecCtlNoHabilitado")
            lblAdvertencia.Visibility = Visibility.Visible
            txtAdvertencia.Visibility = Visibility.Visible
            txtCantCuotas.Text = My.Resources.Cero
        Else
            txtCantCuotas.Style = FindResource("RecCtlHabilitado")
            lblAdvertencia.Visibility = Visibility.Hidden
            txtAdvertencia.Visibility = Visibility.Hidden
        End If
        SValide()
    End Sub
#End Region
#Region "Eventos Ajustes Cuotas para ProgressBar"
    Private Sub SEvnInicio(aobjSender As Object, e As ClsPanEventArgs) Handles MobjOrionCop.EvnInicio
        If e.EnuProceso = EnuProcesoDef.enuAjusCuota Then
            txtProceso.Content = My.Resources.AjusCuoAdm
        ElseIf e.EnuProceso = EnuProcesoDef.None Then
            txtProceso.Content = My.Resources.ProcesoCanc
        End If
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
        MstrResultado = My.Resources.PreProcesados & Format(e.DblCantProcesada, "##0") &
                My.Resources.De & Format(e.DblCantAProcesar, "##0")
        Dispatcher.Invoke(MdgtPgbActualiza,
                System.Windows.Threading.DispatcherPriority.Background,
                New Object() {ProgressBar.ValueProperty, e.DblCantProcesada})
        Dispatcher.Invoke(MdgtLblActualiza,
                System.Windows.Threading.DispatcherPriority.Background,
                New Object() {Label.ContentProperty, MstrResultado})
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Function FblnCantidadValida() As Boolean
        Dim lblnEsValida As Boolean
        If chkNoGenerarAjuste.IsChecked Then
            lblnEsValida = True
        Else
            Dim lstrPerInicio = If(GobjParametros.ObjAnoActual.ObjPeriodoActual.
                    ObjEstaCerradoPeriodoBln.ObjValorPro,
                    ClsOrionCop.FstrPeriodoFinal(GobjParametros.ObjAnoActual.
                    StrIdPeriodoActual, 1),
                    GobjParametros.ObjAnoActual.StrIdPeriodoActual)
            Dim lentPeriodoInicio = CType(Right(lstrPerInicio, 2), Integer)
            Dim lentCanMax As Integer = 13 - lentPeriodoInicio
            lblnEsValida = ClsPanorama.FblnEsValidoNumero(txtCantCuotas.Text, 1, lentCanMax,
                Not chkNoGenerarAjuste.IsChecked, EnuTipoValor.enuInteger)
        End If
        Return lblnEsValida
    End Function
#End Region
End Class
