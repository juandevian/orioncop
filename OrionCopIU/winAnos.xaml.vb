Public Class WinAnos
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuPres
        enuModPorSer
        enuTipoCal
        enuTipoDsctoPP
        enuDiasDscto
        enuDiasMulta
        enuValorMulta
        enuServicioMulta
    End Enum
    Private Enum EnuIdAccion As Integer
        None
        EnuCalCuotas
        EnuCalretroAct
    End Enum
#End Region
    ' Variables
    Private MnuCalcularCuotas As MenuItemPan = Nothing
    Private MnuAjustarCuotasAdmi As MenuItemPan = Nothing
    Private ReadOnly MsepMenu As New Separator
    Private ReadOnly MobjObjetoWin As ClsAno = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomAno
#End Region

#Region "Constructor"
    Friend Sub New(aobjAno As ClsAno)
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuAno
        MobjObjetoWin = aobjAno
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolControlesLlave As New Collection From {
            txtAno
        }
        SAdicioneCtrlsRestringidos()
        SCargueForma(EnuElementosAdicionalesDef.None, 8,
                lcolControlesLlave, txtPresAno, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SLevanteEveNoti("", "", 0, EnuSeveridadNot.EnuOk)
        If MobjObjetoWin.FblnEsAnoActual Then
            MobjObjetoWin.SVerifiqueApp()
        ElseIf MobjObjetoWin.FblnDebeCalcularCuotas Then
            Dim lstrMens = "Se deben calcular las cuotas de administraciòn"
            SLevanteEveNoti(lstrMens, "", 0, EnuSeveridadNot.EnuInformacion)
        End If
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
        If IsNothing(MobjObjetoWin) Then
            Dim lstrEsrror = "El Objeto Pasado en el Constructor de la Ventana no puede ser Null"
            Throw New ErrorInesperadoPanLException(lstrEsrror)
        Else
            ObjObjetoWin = MobjObjetoWin
        End If
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub

    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuPres) = lblPresInicial
        StcValidaControl(EnuValidEntrada.enuModPorSer) = chkModuloPorSer
        StcValidaControl(EnuValidEntrada.enuTipoCal) = lblBaseCalcCA
        StcValidaControl(EnuValidEntrada.enuTipoDsctoPP) = lblTipoDcstoPP
        StcValidaControl(EnuValidEntrada.enuDiasDscto) = lblDiasDscto
        StcValidaControl(EnuValidEntrada.enuDiasMulta) = lblDiasMulta
        StcValidaControl(EnuValidEntrada.enuValorMulta) = lblValorMulta
        StcValidaControl(EnuValidEntrada.enuServicioMulta) = lblServicio
        '
        SHabiliteCtrls()
        SPuebleComboBoxes()
        dgrPeriodos.DataContext = MobjObjetoWin.DtbPeriodos
        HbttAceptar.TabIndex = 25
        HbttCancelar.TabIndex = 26
    End Sub

    Protected Overrides Sub SMuestreDatos()
        HblnMostrandoDatos = True
        With MobjObjetoWin
            txtAno.Content = .ObjIdAnoShr.ToString
            chkEstaCerrado.IsChecked = .ObjEstaCerradoAnoBln.ObjValorPro
            chkModuloPorSer.IsChecked = .ObjModuloPorServicioBln.ObjValorPro
            txtPresAno.Text = Format(.ObjValorPres_AnoDec.ObjValorPro, "c")
            txtPresCalc.Content = Format(.FdecValorCuotaAdminAno, "c")
            chkEstaCalculado.IsChecked = .FblnEstaGenCuota()
            chkEstaAjustado.IsChecked = .FblnEstaAjustada()
            CboBaseCalcCA.SelectedIndex = .ObjTipoCalculoCuotaByt.ObjValorPro
            CboTipoIncentivo.SelectedIndex = .ObjTipoIncentivoByt.ObjValorPro
        End With
        SMuestreIncentivo()
        SHabiliteCtrls()
        Title = My.Resources.Ano & My.Resources.DosPuntosEspacio
        If Not IsNothing(txtAno.Content) Then
            Title &= txtAno.Content.ToString
        End If
        SValide()
        HblnMostrandoDatos = False
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando AndAlso
                MobjObjetoWin.FblnDebeCalcularCuotas Then
            Dim lstrMens = "Se deben calcular las cuotas de administraciòn"
            SLevanteEveNoti(lstrMens, "", 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjValorPres_AnoDec.ObjValorPro = txtPresAno.Text
            .ObjModuloPorServicioBln.ObjValorPro = chkModuloPorSer.IsChecked
            .ObjTipoIncentivoByt.ObjValorPro = CboTipoIncentivo.SelectedIndex
            .ObjTipoCalculoCuotaByt.ObjValorPro = CboBaseCalcCA.SelectedIndex
            If .ObjTipoIncentivoByt.ObjValorPro = EnuTipoIncentivo.EnuDescuentoPP Then
                .ObjDiasMultaExtShr.ObjValorPro = 0
                .ObjValorMultaPagoExtDec.ObjValorPro = 0
                .ObjDiasParaDsctoPPShr.ObjValorPro = txtDiasDscto.Text
                .ObjTipoDsctoPPByt.ObjValorPro = CboTipoDctoPP.SelectedIndex
                .ObjIdServicioMultaShr.ObjValorPro = 0
            ElseIf .ObjTipoIncentivoByt.ObjValorPro = EnuTipoIncentivo.EnuPenalización Then
                .ObjDiasParaDsctoPPShr.ObjValorPro = 0
                .ObjTipoDsctoPPByt.ObjValorPro = EnuTipoDsctoPP.None
                .ObjDiasMultaExtShr.ObjValorPro = txtDiasMulta.Text
                .ObjValorMultaPagoExtDec.ObjValorPro = txtValorMulta.Text
                .ObjIdServicioMultaShr.ObjValorPro = txtIdServicioMulta.Text
            ElseIf .ObjTipoIncentivoByt.ObjValorPro = EnuTipoIncentivo.None Then
                .ObjDiasParaDsctoPPShr.ObjValorPro = 0
                .ObjTipoDsctoPPByt.ObjValorPro = 0
                .ObjDiasMultaExtShr.ObjValorPro = txtDiasMulta.Text
                .ObjValorMultaPagoExtDec.ObjValorPro = txtValorMulta.Text
                .ObjIdServicioMultaShr.ObjValorPro = 0
            End If
        End With
    End Sub

    Protected Overrides Sub SValide()
        With MobjObjetoWin
            StcValidValido(EnuValidEntrada.enuPres) = .ObjValorPres_AnoDec.BlnEsValido
            StcValidValido(EnuValidEntrada.enuModPorSer) = .ObjModuloPorServicioBln.BlnEsValido
            StcValidValido(EnuValidEntrada.enuTipoCal) = .ObjTipoCalculoCuotaByt.BlnEsValido
            StcValidValido(EnuValidEntrada.enuTipoDsctoPP) = .ObjTipoDsctoPPByt.BlnEsValido
            StcValidValido(EnuValidEntrada.enuDiasDscto) = .ObjDiasParaDsctoPPShr.BlnEsValido
            StcValidValido(EnuValidEntrada.enuDiasMulta) = .ObjDiasMultaExtShr.BlnEsValido
            StcValidValido(EnuValidEntrada.enuValorMulta) = .ObjValorMultaPagoExtDec.BlnEsValido
            StcValidValido(EnuValidEntrada.enuServicioMulta) = .ObjIdServicioMultaShr.BlnEsValido
        End With
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub

    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        Dim lsepMenu As New Separator
        MnuCalcularCuotas = FmnuiMenuItemPan("MnuCalcular", "_Calcular Cuotas Administración",
                EnuIdAccion.EnuCalCuotas, "")
        MnuAjustarCuotasAdmi = FmnuiMenuItemPan("MnuAjustarCuotasAdmi",
                "_Generar Retroactivo Cuotas Administración", EnuIdAccion.EnuCalretroAct, "")
        Dim lentIndice = HmnuAcciones.Items.Count - 1
        HmnuAcciones.Items.Insert(lentIndice, lsepMenu)
        HmnuAcciones.Items.Insert(lentIndice, MnuAjustarCuotasAdmi)
        HmnuAcciones.Items.Insert(lentIndice, MnuCalcularCuotas)
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    Protected Overrides Sub SHabiliteMenues()
        SHabiliteCalcular()
    End Sub

    Protected Overrides Sub SRefresqueWin()
        If MobjObjetoWin.EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuConsultando Then
            MyBase.SRefresqueWin()
            dgrPeriodos.DataContext = MobjObjetoWin.DtbPeriodos
            GobjParametros.SRefresqueObj()
            MobjObjetoWin.SRefresqueObj()
            If MobjObjetoWin.FblnEsAnoActual Then
                MobjObjetoWin.SVerifiqueApp()
            End If
            SHabiliteCalcular()
        End If
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub

    Protected Overrides Sub SModifique()
        Dim lstrMens = String.Empty
        If MobjObjetoWin.FblnAnoEsModificable(lstrMens) Then
            MyBase.SModifique()
            SHabiliteCtrls()
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Protected Overrides Sub SFinaliceOperacion()
        MyBase.SFinaliceOperacion()
        SHabiliteCtrls()
    End Sub

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

    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        SDefineServicio()
        Return True
    End Function

    Private Sub SDefineServicio()
        Dim lstrCamposMostrar As String() = {ClsIdServicioShr.SstrNombreCampoBd,
                ClsNombreServicioStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreServicioStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdServicioShr.SstrNombreCampoBd
        Dim lstrTabla As String = ClsServicio.SstrNombreTabla
        Dim lstrFiltro As String = ClsIdCarpetaCuentaShr.SstrNombreCampoBd & " = " &
                GshrIdCarpeta & " AND " & lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Servicio Multa", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SAdicioneCtrlsRestringidos()
        SAdicioneControlRestringido(dgrPeriodos)
        SAdicioneControlRestringido(chkEstaCerrado)
        SAdicioneControlRestringido(txtPresAno)
        SAdicioneControlRestringido(bttEncontrarSer)
    End Sub

    Private Sub SHabiliteCtrls()
        Dim lblnHabilite = EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando
        If lblnHabilite Then
            lblnHabilite = MobjObjetoWin.ObjModuloPorServicioBln.ObjValorPro AndAlso
                    MobjObjetoWin.ObjTipoCalculoCuotaByt.ObjValorPro <>
                    EnuTipoBaseCalculo.EnuImportadas
        End If
        Dim lstrStyle As String
        If lblnHabilite Then
            lstrStyle = "RecCtlHabilitado"
        Else
            lstrStyle = "RecCtlNoHabilitado"
        End If
        txtPresAno.Style = FindResource(lstrStyle)
        If MobjObjetoWin.FblnValorPresExigeCero Then
            txtPresAno.Style = FindResource("RecCtlNoHabilitado")
        End If
    End Sub

    Private Sub SHabiliteCtrlIncentivo()
        cnvDsctoPP.Visibility = Visibility.Collapsed
        cnvMulta.Visibility = Visibility.Collapsed
        If MobjObjetoWin.ObjTipoIncentivoByt.ObjValorPro = EnuTipoIncentivo.EnuDescuentoPP Then
            cnvDsctoPP.Visibility = Visibility.Visible
        ElseIf MobjObjetoWin.ObjTipoIncentivoByt.ObjValorPro =
                EnuTipoIncentivo.EnuPenalización Then
            cnvMulta.Visibility = Visibility.Visible
        End If
    End Sub

    Private Sub SMuestreIncentivo()
        With MobjObjetoWin
            CboTipoIncentivo.SelectedIndex = .ObjTipoIncentivoByt.ObjValorPro
            SHabiliteCtrlIncentivo()
            If .ObjTipoIncentivoByt.ObjValorPro =
                    EnuTipoIncentivo.EnuDescuentoPP Then
                CboTipoDctoPP.SelectedIndex = .ObjTipoDsctoPPByt.ObjValorPro
                txtDiasDscto.Text = .ObjDiasParaDsctoPPShr.ObjValorPro
            ElseIf .ObjTipoIncentivoByt.ObjValorPro =
                    EnuTipoIncentivo.EnuPenalización Then
                txtDiasMulta.Text = .ObjDiasMultaExtShr.ObjValorPro
                txtValorMulta.Text = Format(.ObjValorMultaPagoExtDec.ObjValorPro, "c")
                txtIdServicioMulta.Text = .ObjIdServicioMultaShr.ObjValorPro
                If .ObjIdServicioMultaShr.BlnEsValido AndAlso
                        .FobjServicioMulta IsNot Nothing Then
                    txtServicioMulta.Content = .FobjServicioMulta.ObjNombreServicioStr.ObjValorPro.ToString()
                Else
                    txtServicioMulta.Content = String.Empty
                End If
            End If
        End With
    End Sub

    Private Sub SPuebleComboBoxes()
        Dim ldrwTiposBaseCalc = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuTipoBaseCalculo)
        SPuebleComboBox(ldrwTiposBaseCalc, CboBaseCalcCA)
        ldrwTiposBaseCalc = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuTipoIncentivo)
        SPuebleComboBox(ldrwTiposBaseCalc, CboTipoIncentivo)
        ldrwTiposBaseCalc = ClsOrionCop.FdrwConstantesOri(EnuGrupoConstantesOriDef.EnuTipoDsctoPP)
        SPuebleComboBox(ldrwTiposBaseCalc, CboTipoDctoPP)
    End Sub

    Private Sub SVerifiqueSerMulta()
        MobjObjetoWin.ObjIdServicioMultaShr.ObjValorPro = txtIdServicioMulta.Text
        If Not MobjObjetoWin.ObjIdServicioMultaShr.BlnEsValido Then
            Dim lstrMens = "El Servicio ingresado para la Multa no existe! Desea Crearlo ahora?"
            If MsgBox(lstrMens, vbYesNo, "Crear servicio multa") = MsgBoxResult.Yes Then
                If FblnCreoServicioMulta() Then
                    SLevanteEveNoti("El servicio fue creado existosamente!", "", 0,
                                EnuSeveridadNot.EnuInformacion)
                End If
            End If
        End If
    End Sub

    Private Function FblnCreoServicioMulta() As Boolean
        Dim lblnCreo = False
        Dim lshrIdServicio = GobjParametros.ColServiciosPer.Count + 1
        Dim lstrNombreServicio = "Multa por pago Extemporáneo"
        Dim lwinVentana = New WinServicios(lshrIdServicio, lstrNombreServicio, Nothing) With {
            .EnuOperacionEnWin = EnuOperacionEnWin.CenuCreando,
            .WinPadre = Me
        }
        lwinVentana.ShowDialog()
        Dim lobjServicio As ClsServicio = lwinVentana.ObjObjetoWin
        If lobjServicio IsNot Nothing Then
            GobjParametros.SRefresqueObj()
            MobjObjetoWin.ObjIdServicioMultaShr.ObjValorPro =
                    lobjServicio.ObjIdServicioShr.ObjValorPro
            SMuestreIncentivo()
            lblnCreo = True
        End If
        Return lblnCreo
        SLevanteEveOk()
    End Function
#End Region

#Region "Calculos"
    Private Sub SHabiliteCalcular()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            SHabiliteMenuPan(MnuCalcularCuotas)
            SHabiliteMenuPan(MnuAjustarCuotasAdmi)
            Dim lstrMens = String.Empty
            If MnuCalcularCuotas.IsEnabled Then
                Dim lblnPuedeCalcular = MobjObjetoWin.FblnCalcularCuotas(lstrMens)
                SHabiliteMenuItemPan(lblnPuedeCalcular, MnuCalcularCuotas)
                If Not lblnPuedeCalcular Then
                    If MobjObjetoWin.ObjTipoCalculoCuotaByt.ObjValorPro <>
                            EnuTipoBaseCalculo.EnuImportadas Then
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                End If
            End If
            If MnuAjustarCuotasAdmi.IsEnabled Then
                Dim lblnDebeAjustar = MobjObjetoWin.FblnDebeAjustarCuotasAdmin
                SHabiliteMenuItemPan(lblnDebeAjustar, MnuAjustarCuotasAdmi)
            End If
        End If
    End Sub

    Private Sub SCalculeCuotas()
        Dim lblnNoHayError = False, lblnCalculo = False
        Dim lstrMens As String, lstrMensEx = String.Empty
        lstrMens = "Calculando Cuotas de Administración"
        SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuInformacion)
        Mouse.OverrideCursor = Cursors.Wait
        Try
            lstrMens = String.Empty
            GobjPanDat.SControleProcesoObj(True)
            GobjPanDat.SInicialiceTransaccion()
            lblnCalculo = MobjObjetoWin.FblnCalculoCuotasAdmin(lstrMens)
            SRefresqueWin()
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
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                If lblnCalculo Then
                    GobjPanDat.SConfirmeTransaccion()
                Else
                    GobjPanDat.SAborteTransaccion()
                End If
                GobjPanDat.SControleProcesoObj(False)
            Else
                GobjPanDat.SAborteTransaccion()
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
        SRefrescarClic()
        If String.IsNullOrEmpty(lstrMensEx) Then
            If Not String.IsNullOrEmpty(lstrMens) Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuInformacion)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuOk)
            End If
        Else
            SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
        End If
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub

    Private Sub SAjusteCuotas()
        If GobjParametros.EnuEstadoInstalacion = EnuEstadoInstalacion.Todos Then
            If Not MobjObjetoWin.FblnEstaAjustada Then
                Dim lwinAjusteCuotas As New WinAjusteCuotaAdmin With {
                    .ObjObjetoWin = MobjObjetoWin,
                    .WinPadre = Me
                }
                lwinAjusteCuotas.ShowDialog()
                SRefrescarClic()
            End If
        End If
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            Select Case lelmElemento.Name
                Case "MnuCalcular"
                    SCalculeCuotas()
                Case "MnuAjustarCuotasAdmi"
                    SAjusteCuotas()
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
            If TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is RadioButton Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                    SRegistre()
                    If lelmElemento.Name = "txtIdServicioMulta" Then
                        SVerifiqueSerMulta()
                    End If
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub

    Private Sub OnCboCambio(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is ComboBox Then
            If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando AndAlso
                    Not HblnMostrandoDatos Then
                With MobjObjetoWin
                    Select Case lelmElemento.Name
                        Case "CboBaseCalcCA"
                            .ObjTipoCalculoCuotaByt.ObjValorPro = CboBaseCalcCA.SelectedIndex
                        Case "CboTipoIncentivo"
                            .ObjTipoIncentivoByt.ObjValorPro = CboTipoIncentivo.SelectedIndex
                            .ObjDiasParaDsctoPPShr.ObjValorPro = 0
                            .ObjDiasMultaExtShr.ObjValorPro = 0
                            .ObjValorMultaPagoExtDec.ObjValorPro = 0
                            .ObjTipoDsctoPPByt.ObjValorPro = EnuTipoDsctoPP.None
                            .ObjIdServicioMultaShr.ObjValorPro = 0
                            SHabiliteCtrlIncentivo()
                        Case "CboTipoDctoPP"
                            .ObjTipoDsctoPPByt.ObjValorPro = CboTipoDctoPP.SelectedIndex
                    End Select
                End With
                SMuestreDatos()
            End If
        End If
    End Sub

    Private Sub OnBttClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            With MobjObjetoWin
                If TypeOf lelmElemento Is CheckBox Then
                    If lelmElemento.Name = "chkModuloPorSer" Then
                        .ObjModuloPorServicioBln.ObjValorPro = chkModuloPorSer.IsChecked
                    End If
                    SMuestreDatos()
                ElseIf TypeOf lelmElemento Is Button Then
                    If lelmElemento.Name = "bttEncontrarSer" Then
                        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
                            SBuscar()
                            If BlnBusquedaOk AndAlso
                                    Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                                .ObjIdServicioMultaShr.ObjValorPro = StrResultadoBusqueda
                            End If
                            SMuestreDatos()
                        End If
                    End If
                End If
            End With
        End If
    End Sub
#End Region
End Class