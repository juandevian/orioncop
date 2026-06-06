Imports System.IO
Public Class WinRegistroClave
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuLicenciaUso = 0
        enuClaveCompl
        enuIdTitular
    End Enum
#End Region
    ' Variables
    Private ReadOnly MobjObjetoWin As ClsCBObjetoPan = Nothing
    Private MobjTercero As ClsTercero = Nothing
    Private MstrVolName As String = String.Empty
    Private MlngSerial As Long = 0
    Private MstrSysName As String = String.Empty
    Private MstrClaveInstalacion As String = String.Empty
    Private MstrClaveComplementaria As String = String.Empty
    Private MstrIdLicencia As String = String.Empty
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomRegCla
#End Region
#Region "Constructor"
    Public Sub New(aobjObjeto As ClsCBObjetoPan)
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuRegistroClave
        MobjObjetoWin = aobjObjeto
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 3,
                Nothing, Nothing, True)
        txtIdLicencia.Style = FindResource("RecCtlHabilitado")
        txtClaveCom.Style = FindResource("RecCtlHabilitado")
        txtIdTitular.Style = FindResource("RecCtlHabilitado")
        HbttCancelar.Content = My.Resources.Cancelar
        txtIdLicencia.Focus()
        EnuOperacionEnWin = EnuOperacionEnVentana.cenuModificando
        Dim lstrMens = "Recuerde que para registrar Orión Plus debe ejecutarlo como " &
                "Administrador!"
        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuAdvertencia)
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
        MobjTercero = New ClsTercero(EnuModoInstanciaObjDef.enuUnico)
    End Sub
    Protected Overrides Sub SInicialiceControles()
        lblAviso.Content = "Para obtener la Clave Complementaria, debe solicitarla en" & vbCrLf &
                "Tecnología y Softweare S.A.S." & vbCrLf &
                "donde le solicitarán la Clave de Instalación."
        StcValidaControl(EnuValidEntradaDef.enuLicenciaUso) = lblLicencia
        StcValidaControl(EnuValidEntradaDef.enuClaveCompl) = lblClaveCom
        StcValidaControl(EnuValidEntradaDef.enuIdTitular) = lblIdTitular

        txtPrefijoApp.Content = GobjPanorama.ObjAppActual.ObjDefApp.ObjPrefijoAppStr.ObjValorPro
        txtPrefijoEdi.Content = GobjPanorama.ObjAppActual.ObjDefApp.ObjPrefijoEdicionStr.ObjValorPro
        '
        HbttAceptar.TabIndex = 9
        HbttCancelar.TabIndex = 10
    End Sub
    Protected Overrides Sub SMuestreDatos()
        SMuestreDatosTercero()
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        StcValidValido(EnuValidEntradaDef.enuLicenciaUso) = (Not String.IsNullOrEmpty(MstrIdLicencia) AndAlso
                IsNumeric(txtIdLicencia.Text))
        StcValidValido(EnuValidEntradaDef.enuClaveCompl) = (txtClaveCom.Text = MstrClaveComplementaria AndAlso
                Not String.IsNullOrEmpty(MstrClaveComplementaria))
        StcValidValido(EnuValidEntradaDef.enuIdTitular) = MobjTercero.BlnExiste
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
        GblnOK = False
        SCerrarClic()
    End Sub
    Protected Overrides Sub SGuarde()
        Dim fsiArchivo As FileSystemInfo = Nothing
        If FblnEstanTodosBien() Then
            SElimineArchivos()
            Dim lstrTitulo = ClsPanorama.StrAppTitulo
            Dim lstrComa = My.Resources.Coma
            If ClsPanorama.FstrGetSetting(StrConv(StrReverse(lstrTitulo), vbProperCase),
                        "VBACfg") <> "0" Then
                ClsPanorama.SSaveSetting(StrConv(StrReverse(lstrTitulo), vbProperCase), "VBACfg", "0")
            End If
            If Not My.Computer.FileSystem.FileExists(GobjPanorama.StrFTWN) Then
                With My.Computer.FileSystem
                    .WriteAllText(GobjPanorama.StrFTWN, StrReverse(MstrVolName) & lstrComa, False)
                    .WriteAllText(GobjPanorama.StrFTWN, MlngSerial & lstrComa, True)
                    .WriteAllText(GobjPanorama.StrFTWN, StrReverse(MstrSysName) & lstrComa, True)
                    .WriteAllText(GobjPanorama.StrFTWN, StrReverse(GstrOrigenActual) & lstrComa, True)
                    .WriteAllText(GobjPanorama.StrFTWN, StrReverse(MstrClaveInstalacion) & lstrComa, True)
                    .WriteAllText(GobjPanorama.StrFTWN, StrReverse(MstrClaveComplementaria), True)
                End With
                fsiArchivo = New FileInfo(GobjPanorama.StrFTWN) With {
                        .CreationTimeUtc = Now,
                        .Attributes = FileAttributes.Hidden
                    }
            End If
            If Not My.Computer.FileSystem.FileExists(GobjPanorama.StrFTSN) Then
                With My.Computer.FileSystem
                    .WriteAllText(GobjPanorama.StrFTSN, StrReverse(MstrVolName) & lstrComa, False)
                    .WriteAllText(GobjPanorama.StrFTSN, MlngSerial & lstrComa, True)
                    .WriteAllText(GobjPanorama.StrFTSN, StrReverse(MstrSysName) & lstrComa, True)
                    .WriteAllText(GobjPanorama.StrFTSN, StrReverse(GstrOrigenActual) & lstrComa, True)
                    .WriteAllText(GobjPanorama.StrFTSN, StrReverse(MstrClaveInstalacion) & lstrComa, True)
                    .WriteAllText(GobjPanorama.StrFTSN, MstrClaveComplementaria, True)
                End With
                fsiArchivo = New FileInfo(GobjPanorama.StrFTSN) With {
                        .CreationTimeUtc = Now,
                        .Attributes = FileAttributes.Hidden
                    }
            End If
            GobjPanorama.ObjAppActual.SRegistreLicencia(CType(txtIdLicencia.Text, Short),
                                                            MobjTercero.ObjIdTerceroDbl.ObjValorPro)
            ClsPanorama.SSaveSetting(StrConv(StrReverse(lstrTitulo), vbProperCase), "SBCfg", Date.Today.ToBinary)
            ClsPanorama.SSaveSetting(StrConv(StrReverse(lstrTitulo), vbProperCase), "Licencia", MstrIdLicencia)
            ClsPanorama.SSaveSetting(StrConv(StrReverse(lstrTitulo), vbProperCase), "UsuarioLicencia",
                                         MobjTercero.ObjIdTerceroDbl.ToString)
        End If
        SCerrarClic()
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SMuestreDatosTercero()
        With MobjTercero
            If .BlnExiste AndAlso Not String.IsNullOrEmpty(txtClaveCom.Text) Then
                txtIdTitular.Text = .ObjIdTerceroDbl.ToString
                txtNombreTit.Content = .FstrNombreCompleto()
            End If
        End With
    End Sub
    Private Sub SLeaClaves()
        ClsPanorama.SLeaClavesProteccion("C", MstrVolName, MlngSerial, MstrSysName, MstrClaveInstalacion,
                MstrClaveComplementaria, MstrIdLicencia)
        txtClaveIns.Content = MstrClaveInstalacion
    End Sub
    Private Shared Sub SElimineArchivos()
        Dim fsiArchivo As FileSystemInfo = Nothing
        If My.Computer.FileSystem.FileExists(GobjPanorama.StrVION) Then
            fsiArchivo = New FileInfo(GobjPanorama.StrVION) With {
                .Attributes = FileAttributes.Normal
            }
            My.Computer.FileSystem.DeleteFile(GobjPanorama.StrVION)
        End If
        If My.Computer.FileSystem.FileExists(GobjPanorama.StrFTWN) Then
            fsiArchivo = New FileInfo(GobjPanorama.StrFTWN) With {
                .Attributes = FileAttributes.Normal
            }
            My.Computer.FileSystem.DeleteFile(GobjPanorama.StrFTWN)
        End If
        If My.Computer.FileSystem.FileExists(GobjPanorama.StrFTSN) Then
            fsiArchivo = New FileInfo(GobjPanorama.StrFTSN) With {
                .Attributes = FileAttributes.Normal
            }
            My.Computer.FileSystem.DeleteFile(GobjPanorama.StrFTSN)
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
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim lstrNombreControl As String = lelmElemento.Name
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                If Not HbttCancelar.IsFocused Then
                    Select Case lstrNombreControl
                        Case "txtIdLicencia"
                            If Not String.IsNullOrEmpty(txtIdLicencia.Text) AndAlso
                                    IsNumeric(txtIdLicencia.Text) Then
                                txtIdLicencia.Text = Format(CType(txtIdLicencia.Text, Short), "00000")
                                MstrIdLicencia = txtPrefijoApp.Content & "-" & txtPrefijoEdi.Content & "-" &
                                        txtIdLicencia.Text
                                SLeaClaves()
                            Else
                                MstrIdLicencia = String.Empty
                                MstrClaveComplementaria = String.Empty
                            End If
                        Case "txtClaveCom"
                            txtClaveCom.Text = txtClaveCom.Text.Trim.ToUpper
                        Case "txtIdTitular"
                            If Not String.IsNullOrEmpty(txtIdTitular.Text) AndAlso
                                    IsNumeric(txtIdTitular.Text) Then
                                Dim lobjValorLlave As Object() = {txtIdTitular.Text}
                                MobjTercero.SAbra(lobjValorLlave)
                            End If
                    End Select
                    SMuestreDatos()
                    lblnNoHayError = True
                End If
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
    Private Sub TxtClaveCom_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtClaveCom.TextChanged
        txtClaveCom.Text = txtClaveCom.Text.ToUpper()
        txtClaveCom.SelectionStart = txtClaveCom.Text.Length
    End Sub
#End Region
End Class
