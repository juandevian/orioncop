Imports System.Drawing
Imports Microsoft.Win32
Public Class WinRepLegal
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        EnuIdRepLegal
        EnuCargo
    End Enum
#End Region
    ' Variables
    Private MblnFirmaImportada As Boolean = False
    Private MimgFirma As Image = Nothing
    '
    Private MobjObjetoWin As ClsCentroUtilidad = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomTerceroAdmin
#End Region

#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.EnuTerceroAdmin
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolControlesLlave As New Collection
        SCargueForma(EnuElementosAdicionalesDef.None, 2,
                lcolControlesLlave, Nothing, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
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
        If IsNothing(ObjObjetoWin) Then
            ObjObjetoWin = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub

    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.EnuIdRepLegal) = lblIdRepLegal
        StcValidaControl(EnuValidEntrada.EnuCargo) = lblCargo
        '
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub

    Protected Overrides Sub SMuestreDatos()
        With MobjObjetoWin
            txtIdRepLegal.Text = .ObjIdTerceroRepLegalDbl.ObjValorPro
            txtNombreRepLegal.Content = If(.ObjIdTerceroRepLegalDbl.BlnEsValido,
                    .ObjIdTerceroRepLegalDbl.StrNombreRepLegal, String.Empty)
            txtCargo.Text = .ObjCargoStr.ObjValorPro
        End With
        SCargueFirma()
        SValide()
        SAsigneEtiquetaBtt()
    End Sub

    Protected Overrides Sub SValide()
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

    End Sub
#End Region

#Region "Invalida otros metodos de la clase base"
    Protected Overrides Sub SModifique()
        MyBase.SModifique()
        SAsigneEtiquetaBtt()
    End Sub

    Protected Overrides Sub SGuarde()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            SRegistre()
            SValide()
            If FblnEstanTodosBien() Then
                If MblnFirmaImportada Then
                    If MobjObjetoWin.ObjTerRepLegal.EnuEstadoActualizacion =
                            EnuEstadoObjetoDef.EnuConsultando Then
                        MobjObjetoWin.ObjTerRepLegal.EnuEstadoActualizacion =
                            EnuEstadoObjetoDef.EnuModificando
                    End If
                    MobjObjetoWin.ObjTerRepLegal.SAdicioneImagen(MimgFirma,
                            EnuCategoriaImagenDef.EnuFirmas)
                    MblnFirmaImportada = False
                    MobjObjetoWin.ObjTerRepLegal.SActualice(True)
                End If
                If ObjObjetoWin.BlnTengoCambios Then
                    MobjObjetoWin.SActualice(True)
                End If
                SFinaliceOperacion()
                If Not (WinPadre Is Nothing OrElse WinPadre.EnuIdVentana =
                    EnuIdVentanaDef.EnuMWOrionCop) Then
                    If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
                        SCerrarClic()
                    End If
                Else
                    SAsigneEtiquetaBtt()
                End If
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
                If MblnFirmaImportada Then
                    MblnFirmaImportada = False
                    imgFirma.Source = Nothing
                End If
                If FblnEstanTodosBien() AndAlso .BlnTengoCambios Then
                    Dim lstrMensaje As String = "Los datos del Representante han cambiado!" &
                            vbCrLf & "Desea guardar los cambios?"
                    If MsgBox(lstrMensaje, vbYesNo, "Aceptar Cambios") = vbYes Then
                        If .EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando Then
                            .SActualice(True)
                        End If
                    Else
                        If .EnuEstadoActualizacion = EnuEstadoObjetoDef.EnuModificando Then
                            .SNormaliceEstado(True)
                        End If
                    End If
                Else
                    .SNormaliceEstado(True)
                End If
            End With
        End If
        MyBase.SCancele()
        SAsigneEtiquetaBtt()
    End Sub

    Protected Overrides Sub SFinaliceOperacion()
        MyBase.SFinaliceOperacion()
        SAsigneEtiquetaBtt()
    End Sub
#End Region

#Region "Metodos Propios"
    Private Sub SCaptureTercero()
        Dim lstrIdTercero As String = txtIdRepLegal.Text
        If lstrIdTercero <> MobjObjetoWin.ObjIdTerceroRepLegalDbl.ObjValorPro Then
            If Not String.IsNullOrEmpty(lstrIdTercero) Then
                MobjObjetoWin.ObjIdTerceroRepLegalDbl.ObjValorPro = CType(lstrIdTercero, Double)
                If Not MobjObjetoWin.ObjTerRepLegal.BlnExiste Then
                    SCaptureTercero(lstrIdTercero)
                Else
                    SMuestreDatos()
                End If
            Else
                txtNombreRepLegal.Content = String.Empty
                txtCargo.Text = String.Empty
                imgFirma.Source = Nothing
            End If
        End If
    End Sub

    Private Sub SCaptureTercero(astrIdTercero As String)
        Dim lstrMens = String.Empty
        txtNombreRepLegal.Content = String.Empty
        txtCargo.Text = String.Empty
        imgFirma.Source = Nothing
        If ClsPanorama.FblnEsValidoNumero(astrIdTercero, 1, Double.MaxValue, True,
                        EnuTipoValor.EnuDouble) Then
            If MsgBox("El Tercero ingresado no existe. Desea crearlo ahora?",
                      MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Crear Tercero ?") = MsgBoxResult.Yes Then
                Dim lobjTercero As New ClsTercero(EnuModoInstanciaObjDef.EnuNavegable)
                Dim ldblIdTercero = CType(astrIdTercero, Double)
                lobjTercero.SCreeObj({ldblIdTercero})
                lobjTercero.ObjIdTerceroDbl.ObjValorPro = ldblIdTercero
                Dim lwinVentana As New WinTerceros() With {
                    .ObjObjetoWin = lobjTercero,
                    .EnuOperacionEnWin = EnuOperacionEnWin.CenuCreando,
                    .WinPadre = Me
                }
                lwinVentana.ShowDialog()
                lobjTercero.SAbra({ldblIdTercero})
                If lobjTercero.BlnExiste Then
                    MobjObjetoWin.ObjIdTerceroRepLegalDbl.ObjValorPro = ldblIdTercero
                End If
            End If
        Else
            lstrMens = "El Valor ingresado no es válido!"
        End If
        SMuestreDatos()
        If Not String.IsNullOrEmpty(lstrMens) Then
            SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub

    Private Sub SAsigneEtiquetaBtt()
        If EnuOperacionEnWin = EnuOperacionEnVentana.CenuConsultando Then
            If imgFirma.Source Is Nothing Then
                bttFirma.Content = My.Resources.ImporFirma
                bttFirma.IsEnabled = False
            Else
                bttFirma.Content = My.Resources.ElimFirma
                bttFirma.IsEnabled = True
            End If
        Else
            bttFirma.Content = My.Resources.ImporFirma
            bttFirma.IsEnabled = True
        End If
    End Sub

    Private Sub SCargueFirma()
        If MobjObjetoWin.ObjTerRepLegal IsNot Nothing Then
            Dim lmstImagenGuardada As IO.MemoryStream =
                    MobjObjetoWin.ObjTerRepLegal.FmstFirma()
            If lmstImagenGuardada IsNot Nothing Then
                Dim lbimFirma As New BitmapImage
                lbimFirma.BeginInit()
                lbimFirma.StreamSource = lmstImagenGuardada
                lbimFirma.EndInit()
                imgFirma.Source = lbimFirma
            Else
                imgFirma.Source = Nothing
            End If
        End If
    End Sub

    Private Sub SImporteFirma()
        Dim lofdFirma As New OpenFileDialog With {
            .DefaultExt = ".jpg",
            .Filter = My.Resources.TipoArchivoImagen
        }
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Dim lblnOk As Boolean = lofdFirma.ShowDialog
            If lblnOk Then
                ' Obtener el objeto Image de la Firma
                Dim lstmFirma As IO.Stream = lofdFirma.OpenFile
                Dim llngTamano As Long = lstmFirma.Length
                If llngTamano > 70000 Then
                    lstrMens = "El tamaño máximo permitido de la firma es de 68K. " &
                            "Debe importar una firma más pequeña!"
                End If
                MimgFirma = Bitmap.FromStream(lstmFirma)
                ' Asigno la imagen al control imgFirma de la ventana
                Dim lstrTray As String = lofdFirma.FileName
                Dim lbimFirma As New BitmapImage
                lbimFirma.BeginInit()
                lbimFirma.UriSource = New Uri(lstrTray)
                lbimFirma.DecodePixelWidth = 270
                lbimFirma.EndInit()
                imgFirma.Source = lbimFirma
                lblnNoHayError = True
            End If
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
            MblnFirmaImportada = lblnNoHayError
        End Try
    End Sub

    Private Sub SSuprimaFirma()
        Dim lstrMens As String
        Dim lblnSuprimio = MobjObjetoWin.ObjTerRepLegal.FblnSuprimioFirma()
        imgFirma.Source = Nothing
        SCargueFirma()
        SAsigneEtiquetaBtt()
        lstrMens = If(lblnSuprimio, "La firma fue suprimida!",
                "Se presentó un problema. Si se repite, por favor informe a soporte.")
        SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Controls.Button Then
            If lelmElemento.Name = "bttFirma" Then
                If bttFirma.Content = My.Resources.ElimFirma Then
                    SSuprimaFirma()
                Else
                    SImporteFirma()
                End If
            End If
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
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Try
            GobjPanDat.SControleProcesoObj(True)
            With MobjObjetoWin
                If lelmElemento.Name = "txtIdRepLegal" Then
                    SCaptureTercero()
                ElseIf lelmElemento.Name = "txtCargo" Then
                    .ObjCargoStr.ObjValorPro = txtCargo.Text
                End If
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
#End Region
End Class