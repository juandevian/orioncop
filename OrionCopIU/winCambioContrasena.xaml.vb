Imports System.Windows.Controls
Public Class WinCambioContrasena
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuContrsenaActual = 0
        enuContrasena
        enuTipoCambio
        enuFechaExpiracion
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsUsuario = Nothing
    Private MstrTipoCambio As String = String.Empty
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomCamCon
    Private MblnContActualValida As Boolean = False
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuCambioContraseña
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 4,
                Nothing, pwbContrasenaActual, True)
        GblnOK = False
        SHabiliteCtls(False)
        MobjObjetoWin.ObjTipoCambioContrasenaByt.ObjValorPro = EnuTipoCambioContrasenaDef.enuProximaVez
        dtpFechaExpiracion.Style = FindResource("RecCtlNoHabilitado")
        cboTipoCambio.SelectedIndex = 1
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
            ObjObjetoWin = GobjPanorama.ObjUsuarioActual
            MobjObjetoWin = ObjObjetoWin
        Else
            MobjObjetoWin = ObjObjetoWin
        End If
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuContrsenaActual) = lblContrasenaActual
        StcValidaControl(EnuValidEntrada.enuContrasena) = lblNuevaContrasena
        StcValidaControl(EnuValidEntrada.enuTipoCambio) = lblTipoCambio
        StcValidaControl(EnuValidEntrada.enuFechaExpiracion) = lblFechaExpiracion
        SPuebleComboBoxes()
        HbttAceptar.TabIndex = 16
        HbttCancelar.TabIndex = 17
    End Sub
    Protected Overrides Sub SMuestreDatos()
        For Each lobjPropiedad As ClsCBPropiedad In ObjObjetoWin.ColPropiedades
            Select Case lobjPropiedad.StrNombre
                Case "IdUsuario"
                    txtUsuario.Content = lobjPropiedad.ObjValorPro
                Case "FechaCreacion"
                    txtFechaCreacion.Content = CStr(lobjPropiedad.ObjValorPro)
                Case "FechaCambio"
                    txtFechaUltCambio.Content = CStr(lobjPropiedad.ObjValorPro)
                Case "FechaExpiracion"
                    dtpFechaExpiracion.Text = lobjPropiedad.ObjValorPro.ToString
            End Select
        Next
        Title = "Cambio de Contraseña"
        Title &= ": " & MobjObjetoWin.ObjNombreUsuarioStr.ObjValorPro
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        With MobjObjetoWin
            StcValidValido(EnuValidEntrada.enuContrsenaActual) = MblnContActualValida
            StcValidValido(EnuValidEntrada.enuContrasena) = .ObjContrasenaUsuarioStr.BlnEsValido
            StcValidValido(EnuValidEntrada.enuTipoCambio) = .ObjTipoCambioContrasenaByt.BlnEsValido
            StcValidValido(EnuValidEntrada.enuFechaExpiracion) = .ObjFechaExpiracionDtm.BlnEsValido
        End With
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            If Not String.IsNullOrEmpty(pwbContrasena.Password) Then
                .ObjContrasenaUsuarioStr.ObjValorPro = pwbContrasena.Password
            End If
            If Not String.IsNullOrEmpty(pwbConfirmarContrasena.Password) Then
                .ObjContrasenaUsuarioStr.SConfirmeContrasena(pwbConfirmarContrasena.Password, True)
            End If
            .ObjTipoCambioContrasenaByt.ObjValorPro = cboTipoCambio.SelectedIndex
            .ObjFechaExpiracionDtm.ObjValorPro = dtpFechaExpiracion.DisplayDate
            .ObjFechaCambioContrasenaDtm.ObjValorPro = Date.Today
        End With
    End Sub
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region
#Region "Procedimientos sobrescritos"
    Protected Overrides Sub SGuarde()
        MyBase.SGuarde()
        SLevanteEveNoti("La Contraseña fue cambiada exitosamente!", "", 0,
                EnuSeveridadNot.EnuInformacion)
        GblnOK = True
    End Sub
    Protected Overrides Sub SCancele()
        GblnOK = False
        SCerrarClic()
    End Sub
#End Region
#Region "Procedimientos propios"
    Private Sub SHabiliteCtls(ablnHabilte As Boolean)
        pwbContrasenaActual.IsEnabled = Not ablnHabilte
        pwbConfirmarContrasena.IsEnabled = ablnHabilte
        If ablnHabilte Then
            pwbContrasena.Focus()
        End If
    End Sub
    Private Sub SPuebleComboBoxes()
        Dim ldrwDataRow = ClsAdministrador.FdrwConstantesPan(EnuGrupoConstantesPanDef.enuTipoCambioContrasena)
        SPuebleComboBox(ldrwDataRow, cboTipoCambio)
    End Sub
    Private Sub SChequeeDTPEnabled(astrTipoCambioText As String)
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            If (CByte(astrTipoCambioText.Substring(0, 2)) =
                    EnuTipoCambioContrasenaDef.enuAlVencimiento) Then
                dtpFechaExpiracion.Style = FindResource("RecCtlHabilitado")
            Else
                dtpFechaExpiracion.Style = FindResource("RecCtlNoHabilitado")
            End If
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
            Dim lpwbPasswordBox As PasswordBox = lelmElemento
            lpwbPasswordBox.SelectAll()
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Control Then
            Dim lblnMostrarDatos = True
            Dim lstrNombreCtl = lelmElemento.Name
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                With MobjObjetoWin
                    Select Case lstrNombreCtl
                        Case "pwbContrasenaActual"
                            MobjObjetoWin.ObjContrasenaUsuarioStr.SConfirmeContrasena(pwbContrasenaActual.Password)
                            MblnContActualValida = MobjObjetoWin.ObjContrasenaUsuarioStr.BlnEsValido
                            SHabiliteCtls(MblnContActualValida)
                            SRegistre()
                        Case "pwbContrasena"
                            .ObjContrasenaUsuarioStr.ObjValorPro = pwbContrasena.Password
                        Case "pwbConfirmarContrasena"
                            .ObjContrasenaUsuarioStr.SConfirmeContrasena(pwbConfirmarContrasena.Password, True)
                        Case "dtpFechaExpiracion"
                            .ObjFechaExpiracionDtm.ObjValorPro = dtpFechaExpiracion.DisplayDate
                        Case Else
                            lblnMostrarDatos = False
                    End Select
                End With
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
                If lblnMostrarDatos Then
                    SMuestreDatos()
                End If
            End Try
        End If
    End Sub
    Private Sub ClsFormInterface_Closing(sender As Object, e As ComponentModel.CancelEventArgs)
        If Not GblnOK Then
            MobjObjetoWin.SNormaliceEstado(False)
        End If
    End Sub
    Private Sub CboTipoCambio_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cboTipoCambio.SelectionChanged
        If sender.Equals(cboTipoCambio) Then
            If Not HblnCargandoForma Then
                If cboTipoCambio.Items.Count > 0 Then
                    MobjObjetoWin.ObjTipoCambioContrasenaByt.ObjValorPro = cboTipoCambio.SelectedIndex
                    MstrTipoCambio = cboTipoCambio.SelectedItem.ToString
                    If Not String.IsNullOrEmpty(MstrTipoCambio) Then
                        SChequeeDTPEnabled(MstrTipoCambio)
                    End If
                End If
                SMuestreDatos()
            End If
        End If
    End Sub
#End Region
End Class
