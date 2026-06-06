Public Class WinCuentaCorreoOrigen
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        EnuServidorCorreo
        EnuPuerto
        EnuCtaCorreoOrigen
        EnuContrasenaCorreo
        EnuMensTanda
        EnuInterTandas
    End Enum
#End Region
    'Variables
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomCtaCorreoOri
    Private MblnPoblandoCbo As Boolean = False
    '
    Private MobjObjetoWin As ClsCentroUtilidad = Nothing
    '
    Private ReadOnly MstrServidor As String =
            GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjServidorSmtpStr.ObjValorPro
    Private ReadOnly MblnHabilitarSSL As Boolean =
            GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjHabilitarSslBln.ObjValorPro
    Private ReadOnly MentPuerto As Integer =
            GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjPuertoHostShr.ObjValorPro
    Private ReadOnly MstrMailOrigen As String =
            GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjEmailOrigenStr.ObjValorPro
    Private ReadOnly MblnRequiereAutent As Boolean =
            GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjRequiereAutenticacionBln.ObjValorPro
    Private ReadOnly MstrContrasena As String =
            GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.FstrContrasena
#End Region
#Region "Constructores"
    Friend Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuCorrepOrigen
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneControlRestringido(bttEnviarPrueba)
        SCargueForma(EnuElementosAdicionalesDef.None, 6, Nothing, Nothing, True)
        SPuebleBarraEstadoAdmin(HcolLabelsBarraEstado)
        SValide()
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
        MobjObjetoWin = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
        ObjObjetoWin = MobjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.EnuServidorCorreo) = lblNomServidor
        StcValidaControl(EnuValidEntrada.EnuPuerto) = lblPuerto
        StcValidaControl(EnuValidEntrada.EnuCtaCorreoOrigen) = lblEmailOrigen
        StcValidaControl(EnuValidEntrada.EnuContrasenaCorreo) = lblContraseña
        StcValidaControl(EnuValidEntrada.EnuMensTanda) = lblMensPorTanda
        StcValidaControl(EnuValidEntrada.EnuInterTandas) = lblIntervalo
        '
        SPuebleCboServidores()
        '
        HbttAceptar.TabIndex = 40
        HbttCancelar.TabIndex = 41
    End Sub
    Protected Overrides Sub SMuestreDatos()
        With MobjObjetoWin
            txtServidor.Text = .ObjServidorSmtpStr.ObjValorPro
            txtPuerto.Text = .ObjPuertoHostShr.ObjValorPro
            txtEmailOrigen.Text = .ObjEmailOrigenStr.ObjValorPro
            chkReqAutenti.IsChecked = .ObjRequiereAutenticacionBln.ObjValorPro
            pswConfCon.Password = ClsPanorama.FstrContrasena(.ObjContrasenaEMailStr)
            chkHabilitar.IsChecked = .ObjHabilitarSslBln.ObjValorPro
            txtMensPorTanda.Text = .ObjMensajesPorTandaShr.ObjValorPro
            txtIntervalo.Text = .ObjIntervaloEntreTandasShr.ObjValorPro
            SEstablezcaSer()
        End With
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        With MobjObjetoWin
            StcValidValido(EnuValidEntrada.EnuServidorCorreo) = .ObjServidorSmtpStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuPuerto) = .ObjPuertoHostShr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuCtaCorreoOrigen) = .ObjEmailOrigenStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuContrasenaCorreo) = .ObjContrasenaEMailStr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuMensTanda) = .ObjMensajesPorTandaShr.BlnEsValido
            StcValidValido(EnuValidEntrada.EnuInterTandas) = .ObjIntervaloEntreTandasShr.BlnEsValido
        End With
        '
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjServidorSmtpStr.ObjValorPro = txtServidor.Text
            .ObjPuertoHostShr.ObjValorPro = txtPuerto.Text
            .ObjEmailOrigenStr.ObjValorPro = txtEmailOrigen.Text
            .ObjRequiereAutenticacionBln.ObjValorPro = chkReqAutenti.IsChecked
            .ObjContrasenaEMailStr.ObjValorPro = ClsPanorama.FstrContrasena(True, pswConfCon.Password)
            .ObjHabilitarSslBln.ObjValorPro = chkHabilitar.IsChecked
            .ObjMensajesPorTandaShr.ObjValorPro = txtMensPorTanda.Text
            .ObjIntervaloEntreTandasShr.ObjValorPro = txtIntervalo.Text
        End With
    End Sub
    Protected Overrides Sub SConfigureMenuesPropios()
        HbttGuardar.Visibility = Visibility.Visible
        HbttModificar.Visibility = Visibility.Visible
        HmnuGuardar.Visibility = Visibility.Visible
        HmnuModificar.Visibility = Visibility.Visible
    End Sub
#End Region
#Region "Invalida otros metodos de la clase base"
    Protected Overrides Sub SGuarde()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            SRegistre()
            SValide()
            If FblnEstanTodosBien() Then
                If ObjObjetoWin.BlnTengoCambios Then
                    MobjObjetoWin.SActualice(True)
                End If
                SFinaliceOperacion()
            End If
            lblnNoHayError = True
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
        End Try
    End Sub
    Protected Overrides Sub SCancele()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            With MobjObjetoWin
                If FblnEstanTodosBien() AndAlso .BlnTengoCambios Then
                    Dim lstrMensaje As String = "Los datos de la Copropiedad han cambiado!" &
                            vbCrLf & "Desea guardar los cambios?"
                    If MsgBox(lstrMensaje, vbYesNo, "Aceptar Cambios") = vbYes Then
                        If .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                            .SActualice(True)
                        End If
                    Else
                        If .EnuEstadoActualizacion = EnuEstadoObjetoDef.enuModificando Then
                            .SNormaliceEstado(True)
                        End If
                    End If
                Else
                    .SNormaliceEstado(True)
                End If
            End With
        End If
        MyBase.SCancele()
    End Sub
#End Region
#Region "Métodos de la ventana"
    Private Sub SPuebleCboServidores()
        Dim lstrServidores As ArrayList = MobjObjetoWin.FstrNombreServidoresCorreo
        With cboNomServidor
            MblnPoblandoCbo = True
            .Items.Clear()
            For Each lstrSer As String In lstrServidores
                .Items.Add(lstrSer)
            Next
            MblnPoblandoCbo = False
            .SelectedIndex = 0
        End With
    End Sub
    Private Sub SEstablezcaSer()
        Dim lstrHost = txtServidor.Text
        Dim lentIndiceCbo = 0
        If cboNomServidor.Items.Count > 0 Then
            For i = 0 To cboNomServidor.Items.Count - 1
                If MobjObjetoWin.FstrHostCorreo(i) = lstrHost Then
                    lentIndiceCbo = i
                    Exit For
                End If
            Next
            If lentIndiceCbo > 0 Then
                cboNomServidor.SelectedIndex = lentIndiceCbo
            Else
                cboNomServidor.SelectedIndex = cboNomServidor.Items.Count - 1
            End If
        End If
    End Sub
    Private Sub SEnvieMensPrueba()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim lblnEnvioEmail = False
        If Not FblnEstaConectado(String.Empty, lstrMens) Then
            lstrMens = "No está conectado a Internet en este momento!"
            SLevanteEveNoti(lstrMens, "", 0, EnuSeveridadNot.EnuInformacion)
            Exit Sub
        End If
        SLevanteEveNoti("Enviando Mensaje de Prueba!", "", 0, EnuSeveridadNot.EnuInformacion)
        Dim lstrEmilDest As New ArrayList From {
            txtEmailOrigen.Text
        }
        Try
            lblnEnvioEmail = ClsPanorama.FblnEnvioMensaje(lstrEmilDest, My.Resources.SubjeEmail,
                    My.Resources.MensPrueEmail, "", MobjObjetoWin, MobjObjetoWin.FstrContrasena,
                    lstrMens)
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString()
        Finally
            If lblnNoHayError Then
                If String.IsNullOrEmpty(lstrMens) Then
                    lstrMens = If(lblnEnvioEmail, "El mensaje de prueba fue enviado exitosamente!",
                        "El mensaje no pudo ser enviado!")
                End If
                SLevanteEveNoti(lstrMens, "", 0, EnuSeveridadNot.EnuInformacion)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Try
            GobjPanDat.SControleProcesoObj(True)
            With MobjObjetoWin
                Select Case lelmElemento.Name
                    Case "txtServidor"
                        .ObjServidorSmtpStr.ObjValorPro = txtServidor.Text
                    Case "txtPuerto"
                        .ObjPuertoHostShr.ObjValorPro = txtPuerto.Text
                    Case "txtEmailOrigen"
                        .ObjEmailOrigenStr.ObjValorPro = txtEmailOrigen.Text
                    Case "pswConfCon"
                        .ObjContrasenaEMailStr.ObjValorPro =
                                ClsPanorama.FstrContrasena(True, pswConfCon.Password)
                    Case "txtMensPorTanda"
                        .ObjMensajesPorTandaShr.ObjValorPro = txtMensPorTanda.Text
                    Case "txtIntervalo"
                        .ObjIntervaloEntreTandasShr.ObjValorPro = txtIntervalo.Text
                    Case Else
                        '
                End Select
            End With
            SMuestreDatos()
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
            If lblnNoHayError Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuOk)
                GobjPanDat.SControleProcesoObj(False)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Controls.TextBox Then
            Dim ltxtTexto As Controls.TextBox = lelmElemento
            ltxtTexto.SelectAll()
        ElseIf TypeOf lelmElemento Is PasswordBox Then
            Dim lpswPasw As PasswordBox = lelmElemento
            lpswPasw.SelectAll()
        End If
    End Sub
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Controls.Button Then
            Dim lstrNombreBtt = lelmElemento.Name
            If lstrNombreBtt = "bttEnviarPrueba" Then
                Mouse.OverrideCursor = Cursors.Wait
                SEnvieMensPrueba()
                Mouse.OverrideCursor = Cursors.Arrow
            End If
        End If
    End Sub
    Private Sub Cbo_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles _
            cboNomServidor.SelectionChanged
        Dim lstrMens = String.Empty
        If EnuOperacionEnWin <> EnuOperacionEnVentana.CenuConsultando Then
            If Not MblnPoblandoCbo Then
                Dim lentSele = cboNomServidor.SelectedIndex
                Select Case lentSele
                    Case Is = 0
                        lstrMens = "El Servidor seleccionado no es valido!"
                    Case Is = cboNomServidor.Items.Count - 1
                        With MobjObjetoWin
                            .ObjServidorSmtpStr.ObjValorPro = String.Empty
                            .ObjPuertoHostShr.ObjValorPro = 0
                            .ObjRequiereAutenticacionBln.ObjValorPro = True
                            .ObjHabilitarSslBln.ObjValorPro = True
                        End With
                    Case Else
                        With MobjObjetoWin
                            .ObjServidorSmtpStr.ObjValorPro = .FstrHostCorreo(lentSele)
                            .ObjPuertoHostShr.ObjValorPro = .FentPuertoHost(lentSele)
                            .ObjRequiereAutenticacionBln.ObjValorPro =
                                    .FblnRequiereAutenticacion(lentSele)
                            .ObjHabilitarSslBln.ObjValorPro = .FblnHabilitaSSL(lentSele)
                        End With
                End Select
            End If
        End If
        SMuestreDatos()
        If String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, "", 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub Chk_Click(sender As Object, e As RoutedEventArgs) Handles _
            chkReqAutenti.Click, chkHabilitar.Click
        If sender.Equals(chkHabilitar) Then
            MobjObjetoWin.ObjHabilitarSslBln.ObjValorPro = chkHabilitar.IsChecked
        ElseIf sender.Equals(chkReqAutenti) Then
            MobjObjetoWin.ObjRequiereAutenticacionBln.ObjValorPro = chkReqAutenti.IsChecked
        End If
        SMuestreDatos()
    End Sub
#End Region
End Class
