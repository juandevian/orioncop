Public Class WinGenPrefacturas
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
        enuPermitido = 0
        enuFechaVen
        enuFechaGra
    End Enum
#End Region
    Private WithEvents MobjOrionCop As New ClsOrionCop(GCOBJREGISTRO, False)
    Private ReadOnly MblnDebeCalcular As Boolean = GobjParametros.ObjAnoActual.FblnDebeAjustarCuotasAdmin
    Private MstrResultado As String = String.Empty
    Private MdtmFechaFac As Date = GCDTMFECHANULA
    Private MdtmFechaVen As Date = GCDTMFECHANULA
    Private MdtmFechaGra As Date = GCDTMFECHANULA
    Private MblnFechaFacUnica As Boolean = False
    Private MblnFechaVenUnica As Boolean = False
    Private MblnFechaGraUnica As Boolean = False
    '
    Private MblnCancelando As Boolean = False
    Private MstrNombreVentana As String = My.Resources.NomGenPreF
#End Region

#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuGenPrefacturas
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 3, Nothing, Nothing, True)
        SAdapteEtiquetas()
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SCrearClic()
    End Sub
    Protected Overrides ReadOnly Property StrNombreVentana As String
        Get
            If Not GobjParametros.BlnEFacAutorizado Then
                MstrNombreVentana = "GENERAR PRE-CUENTAS DE COBRO"
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
        MdgtPgbActualiza = New SdgtActualizaProgressBar(AddressOf pgbAvance.SetValue)
        MdgtLblActualiza = New SdgtActualizaLabel(AddressOf lblAvance.SetValue)
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuPermitido) = lblTitulo
        StcValidaControl(EnuValidEntradaDef.enuFechaVen) = lblFechaVen
        StcValidaControl(EnuValidEntradaDef.enuFechaGra) = lblFechaGra
        '
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        txtPeriodoFact.Content = GobjParametros.ObjAnoActual.StrNombrePeriodoActual.ToUpper
        txtFechaFact.Content = If(MblnFechaFacUnica, Format(MdtmFechaFac, GCSTRFMTFECHASIMPLE),
                "Según Servicio de facturación!")
        If MblnFechaVenUnica Then
            dtpFechaVen.SelectedDate = MdtmFechaVen
        Else
            lblFechaVenTxt.Content = "Según Concepto de Facturación!"
        End If
        If MblnFechaGraUnica Then
            dtpFechaGra.SelectedDate = MdtmFechaGra
        Else
            lblFechaGraTxt.Content = "Según Concepto de Facturación!"
        End If
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnFecVenVal = FblnFechaVenValida()
        Dim lblnFecGraVal = FblnFechaGraValida()
        StcValidValido(EnuValidEntradaDef.enuPermitido) = Not MblnDebeCalcular
        StcValidValido(EnuValidEntradaDef.enuFechaVen) = lblnFecVenVal
        StcValidValido(EnuValidEntradaDef.enuFechaGra) = lblnFecGraVal
        FblnEstanTodosBien()
        SHabiliteBotonesTlb()
        If Not lblnFecVenVal Then
            SLevanteEveNoti("Fecha de vencimiento no valida!", String.Empty, 0,
                EnuSeveridadNot.EnuInformacion)
        ElseIf Not lblnFecGraVal Then
            SLevanteEveNoti("Fecha de gracia no valida!", String.Empty, 0,
                EnuSeveridadNot.EnuInformacion)
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
    Protected Overrides Sub SCree()
        SRegistreFechas()
        EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando
        SFrmAdicione()
        If dtpFechaVen.Visibility = Visibility.Visible Then
            dtpFechaVen.Focus()
        End If
        If GobjParametros.BlnEFacAutorizado Then
            Dim lstrMens = String.Empty
            If ClsOrionCop.FblnClientesSinEmail(lstrMens) Then
                MsgBox(lstrMens, vbInformation + vbOKOnly, "Clientes sin email")
                SLevanteEveNoti("Es necesario que los Clientes tengan registrado un Email!",
                            String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                HbttAceptar.IsEnabled = False
            End If
        End If
    End Sub
    Protected Overrides Sub SGuarde()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False,
                lblnGuardo = False
        Dim lenuSeveNoti As EnuSeveridadNot
        MblnCancelando = False
        Try
            If GobjParametros.BlnEFacAutorizado Then
                lstrMens = "Generando Pre-Facturas!"
            Else
                lstrMens = "Generando Pre-Cuentas de cobro!"
            End If
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            lstrMens = String.Empty
            HbttAceptar.IsEnabled = False
            MobjOrionCop.SGenerePrefacturas()
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
                If Not MblnCancelando Then
                    lstrMens = My.Resources.ProTer & " exitosamente!"
                    txtProceso.Content = My.Resources.ProTer
                Else
                    lstrMens = "El Proceso fue cancelado por el Usuario!"
                    MblnCancelando = False
                End If
                lenuSeveNoti = EnuSeveridadNot.EnuInformacion
            Else
                lenuSeveNoti = EnuSeveridadNot.EnuExcep
            End If
            SLevanteEveNoti(lstrMens, lstrMensEx, 0, lenuSeveNoti)
            HbttAceptar.IsEnabled = True
        End Try
    End Sub
    Protected Overrides Sub SCancele()
        If ClsOrionCop.BlnPreFacturando Then
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
            Title = "Generación de Pre-Cuentas de Cobro"
            lblTitulo.Content = "Generar Pre-Cuentas de Cobro..."
            lblPeriodo.Content = "Periodo de las Pre-Cuentas"
            lblFechaFact.Content = "Fecha de las Pre-Cuentas"
            lblFechaVen.Content = "Fecha vencen las Pre-Cuentas"
            lblProceso.Content = "Generando Pre-Cuentas a:"
        End If

    End Sub
    Private Sub SRegistreFechas()
        MblnFechaFacUnica = GobjParametros.FblnFechaFacturacionUnica(MdtmFechaFac)
        MblnFechaVenUnica = GobjParametros.FblnFechaVenceUnica(MdtmFechaVen)
        MblnFechaGraUnica = GobjParametros.FblnFechaGraciaUnica(MdtmFechaGra)
        SMuestreControles()
        SMuestreDatos()
    End Sub
    Private Sub SMuestreControles()
        If MblnFechaFacUnica AndAlso MblnFechaVenUnica Then
            dtpFechaVen.Visibility = Visibility.Visible
            lblFechaVenTxt.Visibility = Visibility.Collapsed
        Else
            dtpFechaVen.Visibility = Visibility.Collapsed
            lblFechaVenTxt.Visibility = Visibility.Visible
        End If
        If MblnFechaFacUnica AndAlso MblnFechaGraUnica AndAlso MblnFechaGraUnica Then
            If MdtmFechaGra = GCDTMFECHAMAXI Then
                dtpFechaGra.Visibility = Visibility.Collapsed
                lblFechaGra.Visibility = Visibility.Collapsed
                lblFechaGraTxt.Visibility = Visibility.Collapsed
            Else
                dtpFechaGra.Visibility = Visibility.Visible
                lblFechaGra.Visibility = Visibility.Visible
            End If
        Else
            dtpFechaGra.Visibility = Visibility.Collapsed
            lblFechaGraTxt.Visibility = Visibility.Visible
        End If
    End Sub
    Private Function FblnFechaVenValida() As Boolean
        Dim lblnEsValida = True
        If MblnFechaVenUnica Then
            For Each lobjFechasFact As ClsFechasServicio In GobjParametros.ColFechasServicio
                lblnEsValida = MdtmFechaVen >= lobjFechasFact.DtmFechaFac
                If Not lblnEsValida Then Exit For
            Next
        End If
        Return lblnEsValida
    End Function
    Private Function FblnFechaGraValida() As Boolean
        Dim lblnEsValida = True
        If MblnFechaGraUnica Then
            For Each lobjFechasFact As ClsFechasServicio In GobjParametros.ColFechasServicio
                lblnEsValida = GobjParametros.FblnFechaGraciaValida(lobjFechasFact, MdtmFechaGra)
                If Not lblnEsValida Then Exit For
            Next
        End If
        Return lblnEsValida
    End Function
#End Region

#Region "Eventos Prefacturas"
    Private Sub EInicio(aobjSender As Object, e As ClsPanEventArgs) Handles MobjOrionCop.EvnInicio
        Select Case e.EnuProceso
            Case EnuProcesoDef.enuPreFacCli
                txtProceso.Content = My.Resources.GenPreClientes
            Case EnuProcesoDef.enuPreFacProPA
                txtProceso.Content = My.Resources.GenPrePropPA
            Case EnuProcesoDef.enuPreFacProSPA
                txtProceso.Content = My.Resources.GenPrePropSPA
            Case EnuProcesoDef.enuPreFacPre
                txtProceso.Content = My.Resources.GenPrePredios
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
            EInicio(Me, e)
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

#Region "Eventos en la Ventana"
    Private Sub Dtp_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles _
            dtpFechaVen.SelectedDateChanged, dtpFechaGra.SelectedDateChanged
        If TypeOf sender Is DatePicker Then
            Dim ldtpSender As DatePicker = sender
            If ldtpSender.Name = "dtpFechaVen" Then
                MdtmFechaVen = dtpFechaVen.SelectedDate
                If MblnFechaVenUnica Then
                    GobjParametros.SEstablezcaFechasServicios(dtpFechaVen.SelectedDate, True)
                End If
            Else
                MdtmFechaGra = dtpFechaGra.SelectedDate
                GobjParametros.SEstablezcaFechasServicios(MdtmFechaGra, False)
            End If
        End If
        SMuestreDatos()
    End Sub
#End Region
End Class
