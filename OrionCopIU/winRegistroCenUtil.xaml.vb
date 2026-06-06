Imports System.Windows.Controls
Public Class WinRegistroCenUtil
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuLicenciaUso = 0
        enuClaveCompl
    End Enum
#End Region
    ' Variables
    Private ReadOnly MobjObjetoWin As ClsCentroUtilOriCop = GobjParametros

    Private MstrClaveComplementaria As String = String.Empty
    Private ReadOnly MstrNombreVentana As String = "REGISTRO Copropiedad"
    Private MblnEsValidaLicencia As Boolean = True
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuRegistroCenUtil
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 2,
                Nothing, Nothing, True)
        txtIdLicencia.Style = FindResource("RecCtlHabilitado")
        txtClaveCom.Style = FindResource("RecCtlHabilitado")
        HbttCancelar.Content = My.Resources.Cancelar
        txtIdLicencia.Focus()
        EnuOperacionEnWin = EnuOperacionEnVentana.cenuModificando
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
            ObjObjetoWin = MobjObjetoWin
        End If
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        lblAviso.Content = "Para obtener la Clave Complementaria, debe solicitarla a" & vbCrLf &
                "Soporte Técnico de Orión Plus"
        StcValidaControl(EnuValidEntradaDef.enuLicenciaUso) = lblLicencia
        StcValidaControl(EnuValidEntradaDef.enuClaveCompl) = lblClaveCom

        txtPrefijoApp.Content = GobjPanorama.ObjAppActual.ObjDefApp.ObjPrefijoAppStr.ObjValorPro
        txtPrefijoEdi.Content = GobjPanorama.ObjAppActual.ObjDefApp.ObjPrefijoEdicionStr.ObjValorPro
        '
        HbttAceptar.TabIndex = 9
        HbttCancelar.TabIndex = 10
    End Sub
    Protected Overrides Sub SMuestreDatos()
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        StcValidValido(EnuValidEntradaDef.enuLicenciaUso) = MblnEsValidaLicencia
        StcValidValido(EnuValidEntradaDef.enuClaveCompl) = (txtClaveCom.Text = MstrClaveComplementaria)
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        '
    End Sub
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region
#Region "Procedimientos sobrescritos"
    Protected Overrides Sub SCancele()
        SCerrarClic()
    End Sub
    Protected Overrides Sub SGuarde()
        Dim lblnNoHayError = False
        If FblnEstanTodosBien() Then
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty
            Try
                GobjPanorama.ObjAppActual.SRegistreCentroUtilAdi()
                lblnNoHayError = True
            Catch ex As ErrorInesperadoPanDatException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As ErrorInesperadoPanLException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As ArgumentException
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Catch ex As Exception
                lstrMens = ex.Message
                lstrMensEx = ex.ToString
            Finally
                If lblnNoHayError Then
                    lstrMens = "Se modificó la Licencia para un Copropiedad adicional"
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                Else
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
        SFinaliceOperacion()
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
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim lstrNombreControl As String = lelmElemento.Name
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                If Not HbttCancelar.IsFocused Then
                    Select Case lstrNombreControl
                        Case "txtIdLicencia"
                            MblnEsValidaLicencia = False
                            Dim lshrIdLicencia = 0S
                            If IsNumeric(txtIdLicencia.Text) Then
                                lshrIdLicencia = CType(txtIdLicencia.Text, Short)
                                If lshrIdLicencia = ClsAdministrador.FobjAppActual.ObjIdLicenciaShr.ObjValorPro Then
                                    MblnEsValidaLicencia = True
                                    txtIdLicencia.Text = Format(CType(txtIdLicencia.Text, Short), "00000")
                                    SGenereClaveCompl()
                                Else
                                    MblnEsValidaLicencia = False
                                    MstrClaveComplementaria = String.Empty
                                End If
                            Else
                                MblnEsValidaLicencia = False
                            End If
                            If Not MblnEsValidaLicencia Then
                                lstrMens = "La Licencia ingresada, '" & txtIdLicencia.Text & "', no es valida!"
                            End If
                        Case "txtClaveCom"
                            If txtClaveCom.Text <> MstrClaveComplementaria Then
                                lstrMens = "La Clave ingresada, '" & txtClaveCom.Text & "', no es valida!"
                            End If
                    End Select
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
                If lblnNoHayError Then
                    If Not String.IsNullOrEmpty(lstrMens) Then
                        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    End If
                Else
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
            End Try
        End If
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SGenereClaveCompl()
        MstrClaveComplementaria = ClsOrionCop.FstrClaveCenUtil
    End Sub
#End Region
End Class
