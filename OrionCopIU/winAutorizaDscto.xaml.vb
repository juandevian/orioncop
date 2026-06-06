Imports System.Windows.Controls
Public Class WinAutorizaDscto
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuIdUsuario
        enuContraseña
        enuAppAsignada
        enuCarpetaAsignada
        enuTienePermiso
    End Enum
#End Region
    Private MobjObjetoWin As ClsUsuario = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomAutDscto
    Friend Property BlnEfact As Boolean = False
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuAutorizaDscto
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneControlRestringido(txtUsuario)
        SAdicioneControlRestringido(pwbContrasena)
        SCargueForma(EnuElementosAdicionalesDef.None, 5,
                Nothing, Nothing, True)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        If BlnEfact Then
            pwbContrasena.Focus()
        Else
            txtUsuario.Focus()
        End If
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
            ObjObjetoWin = New ClsUsuario(EnuModoInstanciaObjDef.enuUnico, False)
        End If
        MobjObjetoWin = ObjObjetoWin
        If BlnEfact Then
            MobjObjetoWin.ObjIdUsuarioStr.ObjValorPro = GCSTRUSUARIOU
            EnuTipoPermisoObjWin = EnuPermisosDef.enuTodos
        Else
            EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
        End If
    End Sub
    Protected Overrides Sub SInicialiceControles()
        txtCarpetaActual.Content = GobjPanorama.ObjCarpetaActual.ObjNombreStr.ObjValorPro
        txtCenUtilActual.Content = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjNombreCentroUtilStr.ObjValorPro
        If BlnEfact Then
            StcValidaControl(EnuValidEntradaDef.enuIdUsuario) = lblContrasena
            StcValidaControl(EnuValidEntradaDef.enuContraseña) = lblIdUsuario
            StcValidaControl(EnuValidEntradaDef.enuAppAsignada) = lblContrasena
            StcValidaControl(EnuValidEntradaDef.enuCarpetaAsignada) = lblContrasena
            StcValidaControl(EnuValidEntradaDef.enuTienePermiso) = lblContrasena
            lblIdUsuario.Content = "Ingrese su Contraseña"
            txtUsuario.Text = "OPT"
            txtUsuario.Visibility = Visibility.Collapsed
            lblContrasena.Visibility = Visibility.Collapsed
        Else
            StcValidaControl(EnuValidEntradaDef.enuIdUsuario) = lblIdUsuario
            StcValidaControl(EnuValidEntradaDef.enuContraseña) = lblContrasena
            StcValidaControl(EnuValidEntradaDef.enuAppAsignada) = lblIdUsuario
            StcValidaControl(EnuValidEntradaDef.enuCarpetaAsignada) = lblIdUsuario
            StcValidaControl(EnuValidEntradaDef.enuTienePermiso) = lblIdUsuario
        End If
        ' 
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        txtUsuario.Text = MobjObjetoWin.ObjIdUsuarioStr.ObjValorPro
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        If Not BlnEfact Then
            With MobjObjetoWin
                Dim lstrMens = String.Empty
                StcValidValido(EnuValidEntradaDef.enuIdUsuario) = .ObjIdUsuarioStr.BlnEsValido AndAlso
                    .ObjEstaActivoUsuarioBln.ObjValorPro
                StcValidValido(EnuValidEntradaDef.enuContraseña) = .ObjContrasenaUsuarioStr.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuAppAsignada) = .BlnTieneAsignadaEstaApp
                If StcValidValido(EnuValidEntradaDef.enuIdUsuario) Then
                    StcValidValido(EnuValidEntradaDef.enuCarpetaAsignada) =
                        MobjObjetoWin.FcolCarpetasUsuario.Contains(GshrIdCarpeta.ToString)
                    If Not StcValidValido(EnuValidEntradaDef.enuCarpetaAsignada) Then
                        lstrMens = "El Usuario " & .ObjIdUsuarioStr.ToString &
                            " no tiene asignada la Carpeta Actual!"
                    Else
                        If Not StcValidValido(EnuValidEntradaDef.enuAppAsignada) Then
                            lstrMens = "El Usuario " & .ObjIdUsuarioStr.ToString &
                                " no tiene asignada esta Aplicación!"
                        Else
                            StcValidValido(EnuValidEntradaDef.enuTienePermiso) =
                                MobjObjetoWin.FblnTienePermiso(EnuIdClasesPanDef.enuReciboCaja, 1)
                            If Not StcValidValido(EnuValidEntradaDef.enuTienePermiso) Then
                                lstrMens = "El Usuario " & .ObjIdUsuarioStr.ToString &
                                    " no tiene Permiso para Descuentos!"
                            End If
                        End If
                    End If
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                End If
            End With
        Else
            SValideEFac()
        End If
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
    Protected Overrides Sub SCancele()
        GblnOK = False
        SCerrarClic()
    End Sub
    Protected Overrides Sub SGuarde()
        SValide()
        Dim lblnGuardo = FblnEstanTodosBien()
        If lblnGuardo Then
            GblnOK = True
        Else
            GblnOK = False
        End If
    End Sub
#End Region
#Region "Busqueda"
    '
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
            If TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is PasswordBox Then
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                    Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                    Try
                        Select Case lelmElemento.Name
                            Case "txtUsuario"
                                If Not BlnEfact Then
                                    MobjObjetoWin.SAbra({txtUsuario.Text})
                                End If
                            Case "pwbContrasena"
                                MobjObjetoWin.ObjContrasenaUsuarioStr.SConfirmeContrasena(pwbContrasena.Password)
                        End Select
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
                        If Not lblnNoHayError Then
                            SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                        End If
                    End Try
                End If
            End If
        End If
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SValideEFac()
        StcValidValido(EnuValidEntradaDef.enuIdUsuario) = True
        StcValidValido(EnuValidEntradaDef.enuAppAsignada) = True
        StcValidValido(EnuValidEntradaDef.enuCarpetaAsignada) = True
        StcValidValido(EnuValidEntradaDef.enuTienePermiso) = True
        StcValidValido(EnuValidEntradaDef.enuContraseña) = MobjObjetoWin.ObjContrasenaUsuarioStr.BlnEsValido
    End Sub
#End Region
End Class
