Imports System.Windows.Controls
Public Class WinAgrupadoresServicios
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuNombre
        enuDiaFra
        enuDiasVen
        enuDiasGra
        enuGraciaFinMes
        enuPie1
        enuPie2
        enuFactProp
        enuFactPropSinPA
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsAgrupadorServicios = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomAgrSer
    Private ReadOnly MshrIdAgruSerInicio As Short = 0
#End Region
#Region "Constructor"
    Public Sub New(ashrIdAgruSer As Short)
        InitializeComponent()
        MshrIdAgruSerInicio = ashrIdAgruSer
        HenuIdVentana = EnuIdVentanaDef.enuAgrupadorServicios
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolCtrlsLlave As New Collection From {
            txtIdAgrupador
        }
        SCargueForma(EnuElementosAdicionalesDef.None, 9, lcolCtrlsLlave, txtNombre, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
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
        If IsNothing(MobjObjetoWin) Then
            MobjObjetoWin = New ClsAgrupadorServicios(EnuModoInstanciaObjDef.enuNavegable)
            If Not MobjObjetoWin.FblnEstaVacioOrigenDatos Then
                Dim lobjValorLlave As Object() = {GshrIdCarpeta, GshrIdCentroUtil,
                        MshrIdAgruSerInicio}
                MobjObjetoWin.SAbra(lobjValorLlave)
            End If
        End If
        ObjObjetoWin = MobjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuNombre) = lblNombreAgru
        StcValidaControl(EnuValidEntrada.enuDiaFra) = lblDiaFactura
        StcValidaControl(EnuValidEntrada.enuDiasVen) = lblDiasVence
        StcValidaControl(EnuValidEntrada.enuDiasGra) = lblPeriodoGracia
        StcValidaControl(EnuValidEntrada.enuGraciaFinMes) = chkGraciaFinMes
        StcValidaControl(EnuValidEntrada.enuPie1) = lblPieFra1
        StcValidaControl(EnuValidEntrada.enuPie2) = lblPieFra2
        StcValidaControl(EnuValidEntrada.enuFactProp) = rdbProPorPreAgr
        StcValidaControl(EnuValidEntrada.enuFactPropSinPA) = rdbProSinPreAgr
        '
        rdbProPorPreAgr.ToolTip = "Seleccionado: Genera una Factura por" & vbCrLf &
                "Propietario y por Predio Agrupador"
        rdbProPorPreAgr.ToolTip = "Seleccionado: Genera una Factura por" & vbCrLf &
                "Propietario y sin Predio Agrupador"
        HbttAceptar.TabIndex = 50
        HbttCancelar.TabIndex = 51
    End Sub
    Protected Overrides Sub SMuestreDatos()
        HblnMostrandoDatos = True
        If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
            SLevanteEveNoti("No hay Conceptos de Facturación para ser mostrados!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            txtIdAgrupador.IsEnabled = False
        Else
            With MobjObjetoWin
                txtIdAgrupador.Text = .ObjIdAgrupadorServiciosShr.ObjValorPro
                txtNombre.Text = .ObjNombreAgrupadorServiciosStr.ObjValorPro
                rdbProPorPreAgr.IsChecked = .ObjFactAPropYPreAgrBln.ObjValorPro
                rdbProSinPreAgr.IsChecked = .ObjFactAPropSinPreAgrBln.ObjValorPro
                If .ObjFactAPropYPreAgrBln.ObjValorPro = False AndAlso
                        .ObjFactAPropSinPreAgrBln.ObjValorPro = False Then
                    rdbNoFactPorPro.IsChecked = True
                End If
                txtDiaFactura.Text = .ObjDiaFacturaShr.ObjValorPro
                txtDiasVence.Text = .ObjDiasVencimientoShr.ObjValorPro
                chkVenceFinMes.IsChecked = .ObjVenceFinMesBln.ObjValorPro
                txtDiasGracia.Text = .ObjDiasGraciaShr.ObjValorPro
                chkGraciaFinMes.IsChecked = .ObjGraciaFinMesBln.ObjValorPro
                txtPieFac1.Text = .ObjPieFacturaUnoStr.ObjValorPro
                txtPieFac2.Text = .ObjPieFacturaDosStr.ObjValorPro
            End With
        End If
        SMuestreFechas()
        ' AVV Reescribir
        Title = "" 'My.Resources.AgrupadorServicios & My.Resources.DosPuntosEspacio
        If Not IsNothing(txtIdAgrupador.Text) Then
            Title &= txtIdAgrupador.Text.ToString
            Title &= " - " & txtNombre.Text
        End If
        SValide()
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            If txtIdAgrupador.Focus Then
                txtIdAgrupador.SelectAll()
            End If
        End If
        HblnMostrandoDatos = False
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentanaDef.cenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntrada.enuNombre) = .ObjNombreAgrupadorServiciosStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuDiaFra) = .ObjDiaFacturaShr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuDiasVen) = .ObjDiasVencimientoShr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuDiasGra) = .ObjDiasGraciaShr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuGraciaFinMes) = .ObjPeriodoGraciaFinMesBln.BlnEsValido
                StcValidValido(EnuValidEntrada.enuPie1) = .ObjPieFacturaUnoStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuPie2) = .ObjPieFacturaDosStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuFactProp) = .ObjFactAPropYPreAgrBln.BlnEsValido
                StcValidValido(EnuValidEntrada.enuFactPropSinPA) = .ObjFactAPropSinPreAgrBln.BlnEsValido
            End With
        End If
        '
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjIdAgrupadorServiciosShr.ObjValorPro = txtIdAgrupador.Text
            .ObjNombreAgrupadorServiciosStr.ObjValorPro = txtNombre.Text
            .ObjFactAPropYPreAgrBln.ObjValorPro = rdbProPorPreAgr.IsChecked
            .ObjFactAPropSinPreAgrBln.ObjValorPro = rdbProSinPreAgr.IsChecked
            .ObjDiaFacturaShr.ObjValorPro = txtDiaFactura.Text
            .ObjDiasVencimientoShr.ObjValorPro = txtDiasVence.Text
            .ObjVenceFinMesBln.ObjValorPro = chkVenceFinMes.IsChecked
            .ObjDiasGraciaShr.ObjValorPro = txtDiasGracia.Text
            .ObjPeriodoGraciaFinMesBln.ObjValorPro = chkGraciaFinMes.IsChecked
            .ObjPieFacturaUnoStr.ObjValorPro = txtPieFac1.Text
            .ObjPieFacturaDosStr.ObjValorPro = txtPieFac2.Text
        End With
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
    Protected Overrides Sub SCree()
        MyBase.SCree()
        SHabilitePieFac()
        'MobjObjetoWin.ObjPieFacturaUnoStr.ObjValorPro = GobjCentroUtilOriCop.FstrPieFacturaUno
        'MobjObjetoWin.ObjPieFacturaDosStr.ObjValorPro = GobjCentroUtilOriCop.FstrPieFacturaDos
        SMuestreDatos()
        txtIdAgrupador.IsEnabled = False
        txtNombre.Focus()
    End Sub
    Protected Overrides Sub SModifique()
        MyBase.SModifique()
        SHabilitePieFac()
    End Sub
    Protected Overrides Sub SGuarde()
        MyBase.SGuarde()
        SRefresqueWin()
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SMuestreFechas()
        txtFechaFactura.Content = Format(MobjObjetoWin.DtmFechaFacturacionPeriodoActual, GCSTRFMTFECHASIMPLE)
        txtFechaVence.Content = Format(MobjObjetoWin.DtmFechaVencePeriActual, GCSTRFMTFECHASIMPLE)
        txtFechaGracia.Content = Format(MobjObjetoWin.DtmFechaGraciaPeriActual, GCSTRFMTFECHASIMPLE)
    End Sub
    Private Sub SHabilitePieFac()
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            'If GobjCentroUtilOriCop.ColAgrupadoresServicios.Count = 0 OrElse
            '        MobjObjetoWin.ObjIdAgrupadorServiciosShr.ObjValorPro = 1 Then
            '    txtPieFac1.Style = FindResource("RecCtlHabilitado")
            'Else
            '    txtPieFac1.Style = FindResource("RecCtlNoHabilitado")
            'End If
        End If
    End Sub
    Private Sub SAbraAgrupador()
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                If txtIdAgrupador.Text <> MobjObjetoWin.ObjIdAgrupadorServiciosShr.ToString() Then
                    Dim lobjVlrLlave = {GshrIdCarpeta, GshrIdCentroUtil,
                            txtIdAgrupador.Text}
                    MobjObjetoWin.SAbra(lobjVlrLlave)
                    SMuestreDatos()
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
        If Not HblnSeEstaCerrando AndAlso Not HblnMostrandoDatos Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is CheckBox Then
                If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
                    With MobjObjetoWin
                        Select Case lelmElemento.Name
                            Case "txtIdAgrupador"
                                .ObjIdAgrupadorServiciosShr.ObjValorPro = txtIdAgrupador.Text
                            Case "txtNombre"
                                .ObjNombreAgrupadorServiciosStr.ObjValorPro = txtNombre.Text
                            Case "txtDiaFactura"
                                .ObjDiaFacturaShr.ObjValorPro = txtDiaFactura.Text
                            Case "txtDiasVence"
                                .ObjDiasVencimientoShr.ObjValorPro = txtDiasVence.Text
                            Case "chkVenceFinMes"
                                .ObjVenceFinMesBln.ObjValorPro = chkVenceFinMes.IsChecked
                            Case "txtDiasGracia"
                                .ObjDiasGraciaShr.ObjValorPro = txtDiasGracia.Text
                            Case "chkGraciaFinMes"
                                .ObjPeriodoGraciaFinMesBln.ObjValorPro = chkGraciaFinMes.IsChecked
                            Case "txtPieFac1"
                                .ObjPieFacturaUnoStr.ObjValorPro = txtPieFac1.Text
                            Case "txtPieFac2"
                                .ObjPieFacturaDosStr.ObjValorPro = txtPieFac2.Text
                            Case "txtIdAgrupador"
                                SAbraAgrupador()
                        End Select
                    End With
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub
    Private Sub Txt_KeyDown(sender As Object, e As KeyEventArgs) Handles txtIdAgrupador.KeyDown
        If e.Key = Key.Return OrElse e.Key = Key.Tab Then
            If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando AndAlso
                    MobjObjetoWin.ObjIdAgrupadorServiciosShr.ToString() <> txtIdAgrupador.Text Then
                SAbraAgrupador()
            End If
        End If
    End Sub
    Private Sub Chk_Click(sender As Object, e As RoutedEventArgs) Handles chkGraciaFinMes.Click,
            chkVenceFinMes.Click
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando AndAlso
                Not HblnMostrandoDatos Then
            With MobjObjetoWin
                Select Case True
                    Case sender.Equals(chkGraciaFinMes)
                        .ObjPeriodoGraciaFinMesBln.ObjValorPro = chkGraciaFinMes.IsChecked
                    Case sender.Equals(chkVenceFinMes)
                        .ObjVenceFinMesBln.ObjValorPro = chkVenceFinMes.IsChecked
                End Select
            End With
            SMuestreDatos()
        End If
    End Sub
    Private Sub RadioButton_Click(sender As Object, e As RoutedEventArgs)
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando AndAlso
                Not HblnMostrandoDatos Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is RadioButton Then
                MobjObjetoWin.ObjFactAPropYPreAgrBln.ObjValorPro = rdbProPorPreAgr.IsChecked
                MobjObjetoWin.ObjFactAPropSinPreAgrBln.ObjValorPro = rdbProSinPreAgr.IsChecked
            End If
            SMuestreDatos()
        End If
    End Sub
#End Region
End Class