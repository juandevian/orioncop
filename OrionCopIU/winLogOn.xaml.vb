Imports System.Windows.Controls ' Ok
Partial Public Class WinLogOn
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuUsuario = 0
        enuClave
        enuCarpeta
        enuCenUtilidad
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsUsuario = Nothing
    Private ReadOnly MwinPadre As MWOrionCop
    Private MstrNombreCarpeta As String = String.Empty
    Private MshrIdCenUtilidad As Short = 0
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomLogOn
    Private MblnCboCarPoblado As Boolean = False
#End Region

#Region "Constructor"
    Friend Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuLogOn
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            SCargueForma(EnuElementosAdicionalesDef.None, 4, Nothing, Nothing, True)
            GblnOK = False
            txtUsuario.Style = FindResource("RecCtlHabilitado")
            pwbContrasena.Style = FindResource("RecCtlHabilitado")
            cboCarpeta.Style = FindResource("RecCtlHabilitado")
            cboCenUtilidad.Style = FindResource("RecCtlHabilitado")
            HbttCancelar.Content = My.Resources.Cancelar
            txtUsuario.Focus()
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentNullException
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens &= ex.Message
            lstrMensEx = ex.ToString
        Finally
            HblnCargandoForma = False
            If lblnNoHayError Then
                FblnEstanTodosBien()
                GobjPanDat.SControleProcesoObj(False)
            Else
                GblnOK = False
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
            If String.IsNullOrEmpty(GstrIdUsuario) AndAlso
                    GenuIdAplicacion = EnuListaAplicaciones.EnuOrionCop Then
                GstrIdUsuario = GCSTRUSUARIOU
            End If
            Dim lobjValorLlave As Object() = {GstrIdUsuario}
            MobjObjetoWin = New ClsUsuario(EnuModoInstanciaObjDef.enuUnico, False)
            MobjObjetoWin.EnuPermisosObj += EnuPermisosDef.enuModificar
            If GstrIdUsuario = GCSTRUSUARIOU Then
                MobjObjetoWin.SLeaUsuarioUniversal()
            ElseIf Not String.IsNullOrEmpty(GstrIdUsuario) Then
                MobjObjetoWin.SAbra(lobjValorLlave)
            End If
            ObjObjetoWin = MobjObjetoWin
        End If
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuUsuario) = lblIdUsuario
        StcValidaControl(EnuValidEntrada.enuClave) = lblContrasena
        StcValidaControl(EnuValidEntrada.enuCarpeta) = lblCarpeta
        StcValidaControl(EnuValidEntrada.enuCenUtilidad) = lblCenUtilidad
        SEstablezcaToolTipGral()
        '
        HbttAceptar.TabIndex = 8
        HbttCancelar.TabIndex = 9
    End Sub
    Protected Overrides Sub SMuestreDatos()
        txtUsuario.Text = MobjObjetoWin.ObjIdUsuarioStr.ToString
        Title = My.Resources.Usuario & MobjObjetoWin.ObjNombreUsuarioStr.ObjValorPro
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        With MobjObjetoWin
            Dim lblnEsValido = MobjObjetoWin.ObjIdUsuarioStr.ObjValorPro = GCSTRUSUARIOU
            If Not lblnEsValido Then
                lblnEsValido = .ObjIdUsuarioStr.BlnEsValido AndAlso
                    .ObjEstaActivoUsuarioBln.ObjValorPro
            End If
            StcValidValido(EnuValidEntrada.enuUsuario) = lblnEsValido
            StcValidValido(EnuValidEntrada.enuClave) = .ObjContrasenaUsuarioStr.BlnEsValido
        End With
        SHabiliteBotonesTlb()
        FblnEstanTodosBien(True)
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjIdUsuarioStr.ObjValorPro = txtUsuario.Text
            .ObjContrasenaUsuarioStr.SConfirmeContrasena(pwbContrasena.Password)
        End With
    End Sub
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region

#Region "Procedimientos sobrescritos"
    Protected Overrides Sub SCancele()
        GblnOK = False
        SCerrarClic()
    End Sub
    Protected Overrides Sub SGuarde()
        Dim lblnNoHayError As Boolean
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty
        Try
            GobjPanDat.SControleProcesoObj(True)
            If FblnEstanTodosBien() Then
                GobjPanorama.ObjCarpetaActual.SEstablezcaCenUtilidadActual(MshrIdCenUtilidad)
                MobjObjetoWin.SAsigneUsuarioActual()
            End If
            lblnNoHayError = True
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
                GblnOK = True
            Else
                GobjPanDat.SControleProcesoObj(False, True)
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SEstablezcaToolTipGral()
        HbttAceptar.ToolTip = "Este Botón permite el ingreso a la Aplicación!"
        HbttCancelar.ToolTip = "Este Botón cancela el ingreso a la Aplicación!"
    End Sub
    Private Sub SLimpieCombos()
        cboCarpeta.Items.Clear()
        cboCenUtilidad.Items.Clear()
        StcValidValido(EnuValidEntrada.enuCarpeta) = False
        StcValidValido(EnuValidEntrada.enuCenUtilidad) = False
        StcValidValido(EnuValidEntrada.enuClave) = False
        pwbContrasena.Password = String.Empty
    End Sub
    Private Function FblnValideApp(ByRef astrMens As String) As Boolean
        Dim lblnAppOk = True
        If MobjObjetoWin.ObjIdUsuarioStr.ObjValorPro <> GCSTRUSUARIOU Then
            If MobjObjetoWin.ObjIdUsuarioStr.BlnEsValido Then
                If Not MobjObjetoWin.BlnTieneAsignadaEstaApp Then
                    astrMens = "El Usuario no tiene asignada ésta Aplicación!"
                    lblnAppOk = False
                ElseIf MobjObjetoWin.FcolCarpetasUsuario.Count = 0 Then
                    astrMens = "El Usuario no tiene Carpetas asignadas!"
                    lblnAppOk = False
                End If
            End If
        End If
        Return lblnAppOk
    End Function
    Private Sub SPuebleCboCarpetas()
        Dim lentIndice = 0
        If MobjObjetoWin.ObjIdUsuarioStr.ObjValorPro = GCSTRUSUARIOU Then
            SPuebCboCarpUsuaUniv(lentIndice)
        Else
            SPuebCboCarpUsua(lentIndice)
        End If
        MblnCboCarPoblado = True
        StcValidValido(EnuValidEntrada.enuCarpeta) = True
        cboCarpeta.SelectedIndex = lentIndice
    End Sub
    Private Sub SPuebCboCarpUsua(ByRef aentIndice As Integer)
        Dim lcolCarpetasUsu = MobjObjetoWin.FcolCarpetasUsuario
        Dim lstrItem As String, lshrIdCarpeta As Short
        cboCarpeta.Items.Clear()
        For Each lobjCarp As ClsCarpetaUsuario In lcolCarpetasUsu
            lshrIdCarpeta = lobjCarp.ObjIdCarpetaUsuarioShr.ObjValorPro
            If lshrIdCarpeta = GshrIdCarpeta Then
                aentIndice = cboCarpeta.Items.Count
            End If
            lstrItem = lshrIdCarpeta.ToString & " " & lobjCarp.ObjNombreCarpetaStr.ToString
            cboCarpeta.Items.Add(lstrItem)
        Next
    End Sub
    Private Sub SPuebCboCarpUsuaUniv(ByRef aentIndice As Integer)
        Dim lstrItem As String, lshrIdCarpeta As Short
        cboCarpeta.Items.Clear()
        Dim lcolCarp = GobjAdministrador.FcolCarpetas(True)
        For Each lobjCarp As ClsCarpeta In lcolCarp
            lshrIdCarpeta = lobjCarp.ObjIdCarpetaShr.ObjValorPro
            If lshrIdCarpeta = GshrIdCarpeta Then
                aentIndice = cboCarpeta.Items.Count
            End If
            lstrItem = lshrIdCarpeta.ToString & " " & lobjCarp.ObjNombreStr.ToString
            cboCarpeta.Items.Add(lstrItem)
        Next
    End Sub
    Private Sub SPuebleCboCenUtilidad()
        Dim lshrIdCentroUtil As Short, lentIndiceCenUtil = 0
        cboCenUtilidad.Items.Clear()
        If cboCarpeta.Items.Count > 0 Then
            Dim lcolCentrosUtil = MobjObjetoWin.FcolCentrosUtilCarUsuario
            Dim lstrNomCenUtil As String
            For Each lobjCenUtil As ClsCentroUtilidad In lcolCentrosUtil
                lshrIdCentroUtil = lobjCenUtil.ObjIdCentroUtilShr.ObjValorPro
                lstrNomCenUtil = lshrIdCentroUtil.ToString & " " &
                            lobjCenUtil.ObjNombreCentroUtilStr.ToString
                cboCenUtilidad.Items.Add(lstrNomCenUtil)
                If lshrIdCentroUtil = GshrIdCentroUtil Then
                    lentIndiceCenUtil = cboCenUtilidad.Items.Count - 1
                End If
            Next
            StcValidValido(EnuValidEntrada.enuCenUtilidad) = lcolCentrosUtil.Count > 0
        End If
        If cboCenUtilidad.Items.Count > 0 Then
            cboCenUtilidad.SelectedIndex = lentIndiceCenUtil
        End If
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        ElseIf TypeOf lelmElemento Is PasswordBox Then
            Dim lpwbCOntrasena As PasswordBox = lelmElemento
            lpwbCOntrasena.SelectAll()
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Dim lblnMostrarDatos = True
        If TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is PasswordBox OrElse
                TypeOf lelmElemento Is ComboBox Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty
            Dim lblnNoHayError = False
            Dim lstrNombreControl As String = lelmElemento.Name
            Try
                GobjPanDat.SControleProcesoObj(True)
                If Not HbttCancelar.IsFocused Then
                    Select Case lstrNombreControl
                        Case "txtUsuario"
                            MobjObjetoWin.ObjIdUsuarioStr.ObjValorPro = txtUsuario.Text
                            SLimpieCombos()
                        Case "pwbContrasena"
                            If MobjObjetoWin.ObjIdUsuarioStr.BlnEsValido Then
                                MobjObjetoWin.ObjContrasenaUsuarioStr.SConfirmeContrasena(pwbContrasena.Password)
                                If MobjObjetoWin.ObjContrasenaUsuarioStr.BlnEsValido AndAlso
                                        FblnValideApp(lstrMens) Then
                                    SPuebleCboCarpetas()
                                End If
                            Else
                                pwbContrasena.Password = String.Empty
                            End If
                        Case Else
                            lblnMostrarDatos = False
                    End Select
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
                If lblnNoHayError Then
                    GobjPanDat.SControleProcesoObj(False)
                    If lblnMostrarDatos Then
                        SMuestreDatos()
                    End If
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, "", 0, EnuSeveridadNot.EnuInformacion)
                    End If
                Else
                    GobjPanDat.SControleProcesoObj(False, True)
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub
    Private Sub CboCarpeta_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) _
            Handles cboCarpeta.SelectionChanged
        If e.AddedItems.Count > 0 Then
            Dim lstrItem As String = e.AddedItems(0)
            If Not String.IsNullOrEmpty(lstrItem) Then
                MobjObjetoWin.SEstablezcaCarpetaActual(CType(lstrItem.Split(" ")(0), UShort))
                MstrNombreCarpeta = GobjPanorama.ObjCarpetaActual.ObjNombreStr.ObjValorPro
                If mblnCboCarPoblado Then
                    SPuebleCboCenUtilidad()
                End If
            End If
        End If
    End Sub
    Private Sub CboCenUtilidad_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) _
            Handles cboCenUtilidad.SelectionChanged
        If e.AddedItems.Count > 0 Then
            Dim lstrItem As String = e.AddedItems(0)
            If Not String.IsNullOrEmpty(lstrItem) Then
                MshrIdCenUtilidad = CType(lstrItem.Split(" ")(0), UShort)
            End If
        End If
    End Sub
    Private Sub ClsFormInterface_Closed(sender As Object, e As EventArgs)
        If GblnOK Then
            If Not FblnEstanTodosBien() OrElse Not HblnSeEstaCerrando Then
                GblnOK = False
            End If
        End If
    End Sub
#End Region
End Class
