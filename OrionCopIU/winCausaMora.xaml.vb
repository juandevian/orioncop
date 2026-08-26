Public Class WinCausaMora
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
    Private Enum EnuTipoAccion As Byte
        None
        EnuSoloCierra
        EnuSoloCausa
        EnuAmbas
    End Enum
#End Region
    Private WithEvents MobjOrionCop As New ClsOrionCop(GCOBJREGISTRO, False)
    Private WithEvents MobjReportes As New ClsRepOrionCop(GCOBJREGISTRO)
    Private MstrResultado As String = String.Empty
    Private MblnCancelando As Boolean = False
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomCausaMora
    Private MenuTipoAccion As EnuTipoAccion = EnuTipoAccion.None
    Private ReadOnly MblnCierraMes As Boolean = False
    Private MobjInterfazCont As ClsCBInterfazContableOri = Nothing
    '
    Private ReadOnly MstrArchivoCopia As String = String.Empty
#End Region

#Region "Constructor"
    Public Sub New(ablnCierrePeriodo As Boolean)
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuCausaMora
        MblnCierraMes = ablnCierrePeriodo
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            SCargueForma(EnuElementosAdicionalesDef.None, 0,
                Nothing, Nothing, True)
            SPuebleBarraEstado(HcolLabelsBarraEstado)
            SCree()
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
            ObjObjetoWin = New ClsNotaDb()
            ObjObjetoWin.SInicialiceObj()
        End If
        Dim lobjObjetoWin As ClsNotaDb = ObjObjetoWin
        EnuTipoPermisoObjWin = lobjObjetoWin.EnuPermisosObj
        MenuTipoAccion = If(GobjParametros.ObjExigeFechaHoyDocsBln.ObjValorPro,
                If(MblnCierraMes, EnuTipoAccion.EnuSoloCierra, EnuTipoAccion.EnuSoloCausa),
                EnuTipoAccion.EnuAmbas)
        Select Case GobjParametros.ObjIdAppContableByt.ObjValorPro
            Case EnuAppConta.EnuApoloBD, EnuAppConta.EnuApoloAP
                MobjInterfazCont = New ClsInterfazApolo(GCOBJREGISTRO)
            Case EnuAppConta.EnuContaPyme
                MobjInterfazCont = New ClsInterfazContaPyme(GCOBJREGISTRO)
            Case EnuAppConta.EnuSIIGO
                MobjInterfazCont = New ClsInterfazContaSIIGO(GCOBJREGISTRO)
            Case EnuAppConta.EnuSIIGON
                Dim lobjIntCon As New ClsInterfazContaSIIGO(GCOBJREGISTRO)
                MobjInterfazCont = lobjIntCon
                lobjIntCon.BlnNube = True
            Case EnuAppConta.EnuColon
                MobjInterfazCont = New ClsInterfazContaColon(GCOBJREGISTRO)
            Case EnuAppConta.EnuMekano
                MobjInterfazCont = New ClsInterfazMekano(GCOBJREGISTRO)
        End Select
    End Sub
    Protected Overrides Sub SInicialiceControles()
        SMuestreCtrls()
        '
        MdgtLblActualiza = New SdgtActualizaLabel(AddressOf lblResultado.SetValue)
        MdgtPgbActualiza = New SdgtActualizaProgressBar(AddressOf pgbCausacion.SetValue)
        '
        HbttAceptar.TabIndex = 1
        HbttCancelar.TabIndex = 2
    End Sub
    Protected Overrides Sub SMuestreDatos()
        Select Case MenuTipoAccion
            Case EnuTipoAccion.EnuSoloCierra
                txtFechaCausa.Content = "No causa Intereses de Mora."
                txtCierreMes.Content = "Si: " & GobjParametros.ObjAnoActual.StrNombrePeriodoActual
            Case EnuTipoAccion.EnuSoloCausa
                txtFechaCausa.Content = Format(Date.Today, GCSTRFMTFECHASIMPLE)
                txtCierreMes.Content = "No. Solo Causa Intereses de Mora."
            Case EnuTipoAccion.EnuAmbas
                Dim ldtmFechaCausaraMora = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
                ldtmFechaCausaraMora = ldtmFechaCausaraMora.AddDays(1)
                txtFechaCausa.Content = Format(ldtmFechaCausaraMora, GCSTRFMTFECHASIMPLE)
                txtCierreMes.Content = "Si: " & GobjParametros.ObjAnoActual.StrNombrePeriodoActual
        End Select
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
    Protected Overrides Sub SCree()
        SFrmAdicione()
        EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando
    End Sub

    Protected Overrides Sub SGuarde()
        Dim lblnCauso = False, lblnNoHayError = False, lstrMensImp = String.Empty
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lenuSevNot = EnuSeveridadNot.None
        Try
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            Select Case MenuTipoAccion
                Case EnuTipoAccion.EnuSoloCierra
                    SCierreMes()
                Case EnuTipoAccion.EnuSoloCausa
                    lblnCauso = FblnCausoMora(lstrMens, lenuSevNot)
                Case EnuTipoAccion.EnuAmbas
                    lblnCauso = FblnCerroYCauso(lstrMens, lenuSevNot)
            End Select
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Catch ex As ArgumentNullException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Finally
            If lblnNoHayError Then
                If Not MblnCancelando Then
                    GobjPanDat.SConfirmeTransaccion()
                    GobjPanDat.SControleProcesoObj(False)
                    If MenuTipoAccion <> EnuTipoAccion.EnuSoloCierra Then
                        If lblnCauso Then
                            If String.IsNullOrEmpty(lstrMens) Then
                                lstrMens = If(EnuTipoAccion.EnuSoloCausa, "Causación de Intereses " &
                                        "terminada exitosamente!", "Causación de Intereses y Cierre " &
                                        "de Mes terminados exitosamente!")
                                lenuSevNot = EnuSeveridadNot.EnuInformacion
                            End If
                        End If
                    Else
                        lstrMens = "Cierre de Mes terminado exitosamente!"
                        lenuSevNot = EnuSeveridadNot.EnuInformacion
                    End If
                    txtProceso.Content = My.Resources.ProTer
                Else
                    lstrMens = "Proceso cancelado por el Usuario!"
                    lenuSevNot = EnuSeveridadNot.EnuInformacion
                End If
                SEstablezcaWinConsultando()
            Else
                GobjPanDat.SControleProcesoObj(False, True)
                GobjPanDat.SAborteTransaccion()
                lenuSevNot = EnuSeveridadNot.EnuExcep
            End If
            SLevanteEveNoti(lstrMens, lstrMensEx, 0, lenuSevNot)
            HbttAceptar.IsEnabled = True
            HbttCancelar.IsEnabled = True
            GblnOK = lblnCauso
        End Try
    End Sub

    Protected Overrides Sub SCancele()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuCreando Then
            If GblnCausandoFM Then
                MblnCancelando = True
            Else
                GblnOK = False
                SCerrarClic()
            End If
        Else
            SCerrarClic()
        End If
    End Sub

    Protected Overrides Sub SCerrarClic()
        MyBase.SCerrarClic()
        MblnCancelando = False
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SMuestreCtrls()
        Select Case MenuTipoAccion
            Case EnuTipoAccion.EnuSoloCierra
                lblTitulo.Content = "Cierra Mes ..."
            Case EnuTipoAccion.EnuSoloCausa
                lblTitulo.Content = "Causa Intereses de Mora ..."
            Case EnuTipoAccion.EnuAmbas
                lblTitulo.Content = "Causa Intereses de Mora y Cierra Mes ..."
        End Select
    End Sub

    Private Sub SCierreMes()
        Dim lentItemsMulta = 0
        HbttAceptar.IsEnabled = False
        HbttCancelar.IsEnabled = False
        Mouse.OverrideCursor = Cursors.Wait
        SGenereBK()
        If GobjParametros.ObjAnoActual.ObjTipoIncentivoByt.ObjValorPro =
                EnuTipoIncentivo.EnuPenalización Then
            Dim lobjOrionCop As New ClsOrionCop(GCOBJREGISTRO, False)
            lobjOrionCop.SGenereItemsProgFactMulta(lentItemsMulta)
        End If
        SGenReportesFinMes(lentItemsMulta)
        Mouse.OverrideCursor = Cursors.Wait
        GobjParametros.SCierrePeriodo()
        If GobjParametros.ObjTipoInterfazByt.ObjValorPro <> EnuTipoInterfazDef.None Then
            SGenereInterfazCon()
        End If
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub

    ' Inicio causa mora FM
    Private Function FblnCausoMora(ByRef astrMens As String,
            ByRef aenuSeve As EnuSeveridadNot) As Boolean ' FM
        HbttCancelar.IsEnabled = True
        Dim lblnCauso = MobjOrionCop.FblnCausoMoraGeneral(astrMens)
        If lblnCauso AndAlso Not MblnCancelando Then
            HbttCancelar.IsEnabled = False
            Dim lstrIdNotasDb = ClsOrionCop.FstrIdUltimasNotasDb
            If Not String.IsNullOrEmpty(lstrIdNotasDb) Then
                SImprimaNotasDb(lstrIdNotasDb, True)
                If Not String.IsNullOrEmpty(astrMens) Then
                    aenuSeve = EnuSeveridadNot.EnuAdvertencia
                End If
                If GobjParametros.BlnEFacAutorizado Then
                    SProceseNotasApi()
                End If
            Else
                astrMens = "No se generaron Notas débito por Intereses de Mora!"
                aenuSeve = EnuSeveridadNot.EnuInformacion
            End If
            HbttCancelar.IsEnabled = True
        ElseIf MblnCancelando Then
            If MenuTipoAccion = EnuTipoAccion.EnuAmbas Then
                GobjParametros.SAbraPeriodoAnterior()
            End If
        End If
        Return lblnCauso
    End Function

    Private Function FblnCerroYCauso(ByRef astrMens As String,
            ByRef aenuSeve As EnuSeveridadNot) As Boolean
        If ClsOrionCop.FblnDebeCrearAno() Then
            astrMens = "Antes de cerrar mes, debe crear el nuevo año!"
            aenuSeve = EnuSeveridadNot.EnuInformacion
            MsgBox(astrMens, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Información")
            Return False
        End If
        SCierreMes()
        Dim lblnCauso = FblnCausoMora(astrMens, aenuSeve)
        Return lblnCauso
    End Function

    Private Sub SGenReportesFinMes(aentItemsMulta As Integer)
        Mouse.OverrideCursor = Cursors.Wait
        Dim ldtmFechaCierre As Date =
                GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaFinPeriodo
        ldtmFechaCierre = ldtmFechaCierre.AddDays(1)
        If GobjParametros.ObjAnoActual.ObjTipoIncentivoByt.ObjValorPro =
                EnuTipoIncentivo.EnuPenalización AndAlso aentItemsMulta > 0 Then
            MobjReportes.SGenereRepFacturasMultadas(ldtmFechaCierre)
        End If
        Mouse.OverrideCursor = Cursors.Wait
        Dim NoUsado = MobjReportes.SGenereEdadCartera(GentLimite1, GentLimite2, GentLimite3,
                GentLimite4, EnuTipoRepEdadCartera.enuResumido, ldtmFechaCierre,
                True, True)
        MobjReportes.SGenereAntPorAplicar(ldtmFechaCierre, True)
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub

    Private Sub SGenereInterfazCon()
        Dim lwinIntCon As New WinInterfazCont(True) With {
            .WinPadre = Me
        }
        lwinIntCon.ShowDialog()
    End Sub

    Private Sub SProceseNotasApi()
        Dim lstrMens = String.Empty
        SProceseEFac(lstrMens)
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
#End Region

#Region "Eventos Causacion Mora"
    Private Sub SEvnInicio(aobjSender As Object, e As ClsPanEventArgs) Handles _
            MobjOrionCop.EvnInicio, MobjReportes.EvnInicio
        lblResultado.Visibility = Visibility.Visible
        pgbCausacion.Visibility = Visibility.Visible
        Select Case e.EnuProceso
            Case EnuProcesoDef.EnuBK
                txtProceso.Content = My.Resources.GenBK
            Case EnuProcesoDef.EnuCausaMora
                txtProceso.Content = My.Resources.GenInte
            Case EnuProcesoDef.enuRepEdadCar
                txtProceso.Content = My.Resources.GenEdadCar
                lblResultado.Visibility = Visibility.Hidden
                pgbCausacion.Visibility = Visibility.Hidden
            Case EnuProcesoDef.enuRepAntXApl
                txtProceso.Content = My.Resources.AntPorApl
                lblResultado.Visibility = Visibility.Hidden
            Case EnuProcesoDef.None
                txtProceso.Content = My.Resources.ProcesoCanc
        End Select
        pgbCausacion.Minimum = 0.0
        pgbCausacion.Maximum = e.DblCantAProcesar - 1
        pgbCausacion.Value = 0.0
    End Sub

    Private Sub SEvnAvance(aobjSender As Object, e As ClsPanEventArgs) Handles MobjOrionCop.EvnAvance,
            MobjReportes.EvnAvance
        If MblnCancelando Then
            e.BlnCancele = True
            e.EnuProceso = EnuProcesoDef.None
            SEvnInicio(Me, e)
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