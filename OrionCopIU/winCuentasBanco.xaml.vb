Imports Microsoft.Win32
Imports System.ComponentModel
Public Class WinCuentasBanco
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntradaDef As Integer
        enuNombreBco = 0
        enuNroCuenta
        enuCtaCont
    End Enum
#End Region
    ' Menu
    Private MmnuAsociarQR As MenuItem = Nothing
    Private MmnuSuprimirQR As MenuItem = Nothing
    Private MimgQR As System.Drawing.Image = Nothing
    Private MblnQRImportado As Boolean = False
    ' Variables
    Private MobjObjetoWin As ClsCuentaBanco = Nothing
    Private MblnDejoUltimoControl As Boolean = False
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomCtaBco
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuCuentasBanco
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolCamposLlaves As New Collection
        SAdicioneCtrlsRestringidos()
        lcolCamposLlaves.Add(txtIdCuenta)
        SCargueForma(EnuElementosAdicionalesDef.None, 3,
                lcolCamposLlaves, txtNombreBancoNuevo, False)
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
            ObjObjetoWin = New ClsCuentaBanco(EnuModoInstanciaObjDef.enuNavegable)
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuNombreBco) = lblNombreBanco
        StcValidaControl(EnuValidEntradaDef.enuNroCuenta) = lblNumeroCuenta
        StcValidaControl(EnuValidEntradaDef.enuCtaCont) = lblIdCtaContab
        '
        SDeshabiliteControlesActuales()
        SModifiqueBarraHerramientas()
        cnvCuentasBanco.DataContext = GobjParametros.FdtbCuentasBanco()
        SEstablezcaToolTipGral()
        ' 
        HbttAceptar.TabIndex = 6
        HbttCancelar.TabIndex = 7
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If Not IsNothing(MobjObjetoWin) Then
            With MobjObjetoWin
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                    If GobjParametros.FdtbCuentasBanco.Rows.Count = 0 Then
                        SLevanteEveNoti("No hay Cuentas Bancarias para ser mostradas!", "", 0,
                                EnuSeveridadNot.EnuInformacion)
                    End If
                    SCargueCodigoQR()
                End If
                txtNombreBancoNuevo.Text = .ObjNombreBancoStr.ObjValorPro
                txtNumeroCuentaNuevo.Text = .ObjNumeroCuentaStr.ObjValorPro
                txtIdCtaContabNueva.Text = .ObjIdCtaContabilidadStr.ObjValorNuevo
                chkActiva.IsChecked = .ObjEstaActivaBln.ObjValorPro
                txtNomCtaContab.Content = .ObjIdCtaContabilidadStr.StrNombreCuenta
            End With
            SHabiliteMenuesQR()
            SValide()
        End If
    End Sub
    Protected Overrides Sub SValide()
        With MobjObjetoWin
            If GobjParametros.FdtbCuentasBanco().Rows.Count = 0 AndAlso EnuOperacionEnWin =
                    EnuOperacionEnVentana.cenuConsultando Then
                SInicialiceValido()
            Else
                StcValidValido(EnuValidEntradaDef.enuNombreBco) = .ObjNombreBancoStr.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuNroCuenta) = .ObjNumeroCuentaStr.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuCtaCont) = .ObjIdCtaContabilidadStr.BlnEsValido
            End If
        End With
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjNombreBancoStr.ObjValorPro = txtNombreBancoNuevo.Text
            .ObjNumeroCuentaStr.ObjValorPro = txtNumeroCuentaNuevo.Text
            .ObjIdCtaContabilidadStr.ObjValorPro = txtIdCtaContabNueva.Text
            .ObjEstaActivaBln.ObjValorPro = chkActiva.IsChecked
        End With
        SValide()
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        Dim lmnuOpcionesQR As MenuItem = FindResource("RecMnuOpcQR")
        HmnuMiMenu.Items.Insert(1, lmnuOpcionesQR)
        MmnuAsociarQR = FmnuiMenuItemPan("MmnuAsociarQR", "A_sociar Código QR", 3, "")
        lmnuOpcionesQR.Items.Insert(0, MmnuAsociarQR)
        MmnuSuprimirQR = FmnuiMenuItemPan("MmnuSuprimirQR", "S_uprimir Código QR", 4, "")
        lmnuOpcionesQR.Items.Insert(1, MmnuSuprimirQR)
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    ''' <summary>
    ''' Sub que prepara a la ventana y a su objeto para crear un nuevo objeto. Invalida el Sub
    ''' "SCree" de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SCree()
        MyBase.SCree()
        SCreeCuentaBanco()
        txtNombreBancoNuevo.Focus()
    End Sub
    ''' <summary>
    ''' Prepara la ventana y su objeto para modificar el objeto. Invalida la función "SModifique"
    ''' de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SModifique()
        MyBase.SModifique()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            SMuestreDatos()
            SModifiqueCuentaBanco()
        End If
        SHabiliteMenuesQR()
    End Sub
    Protected Overrides Sub SGuarde()
        SRegistre()
        If FblnEstanTodosBien() Then
            If MblnQRImportado Then
                MobjObjetoWin.SAdicioneImagen(MimgQR)
            End If
            If ObjObjetoWin.BlnTengoCambios OrElse MblnQRImportado Then
                MobjObjetoWin.SActualice(True)
            End If
            MblnQRImportado = False
            SFinaliceOperacion()
        End If
    End Sub
    Protected Overrides Sub SSuprima()
        MyBase.SSuprima()
        SFinaliceOperacion()
        If IsNothing(MobjObjetoWin) Then
            SInicialiceObjeto()
        End If
    End Sub
    Protected Overrides Sub SRefresqueWin()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Mouse.OverrideCursor = Cursors.Wait
            cnvCuentasBanco.DataContext = GobjParametros.FdtbCuentasBanco()
            SMuestreDatos()
            Mouse.OverrideCursor = Cursors.Arrow
        End If
    End Sub
    Protected Overrides Sub SFinaliceOperacion()
        Dim lblnSelecionarUltimo = (EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando)
        Dim ldtbCtasBanco = GobjParametros.FdtbCuentasBanco()
        MyBase.SFinaliceOperacion()
        SVisibiliceControlesNuevos(False)
        dgrCuentasBanco.IsEnabled = True
        cnvCuentasBanco.DataContext = ldtbCtasBanco
        If dgrCuentasBanco.Items.Count > 0 Then
            If lblnSelecionarUltimo Then
                Dim lentIndice As Integer = ldtbCtasBanco.Rows.Count - 1
                dgrCuentasBanco.SelectedIndex = lentIndice
            Else
                dgrCuentasBanco.SelectedIndex = 0
            End If
        End If
        SOrdeneDataGrid(dgrCuentasBanco, dgrCuentasBanco.Columns(0),
                ClsIdCuentaBancoShr.SstrNombreCampoBd,
                ListSortDirection.Ascending)
        SMuestreDatos()
    End Sub
    Protected Overrides Sub SHabiliteMenues()
        '
    End Sub
#End Region
#Region "Manejo Código QR"
    Private Sub SHabiliteMenuesQR()
        Dim lblnPuedeAsociar = ClsOrionCop.FdtbImagenBancoQR(0).Rows.Count = 0 AndAlso
                chkActiva.IsChecked
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            SHabiliteMnuAsociarQR(lblnPuedeAsociar)
            SHabiliteMnuSuprimirQR(False)
        Else
            SHabiliteMnuAsociarQR(False)
            If Not IsNothing(imgCodigoQR.Source) Then
                SHabiliteMnuSuprimirQR(True)
            Else
                SHabiliteMnuSuprimirQR(False)
            End If
        End If
    End Sub
    Private Sub SHabiliteMnuAsociarQR(ablnHabilte As Boolean)
        If ablnHabilte Then
            SHabiliteMenuItemPan(True, MmnuAsociarQR)
        Else
            If MmnuAsociarQR.IsEnabled Then
                SHabiliteMenuItemPan(False, MmnuAsociarQR)
            End If
        End If
        MnuAsociarQRC.IsEnabled = MmnuAsociarQR.IsEnabled
    End Sub
    Private Sub SHabiliteMnuSuprimirQR(ablnHabilte As Boolean)
        If ablnHabilte Then
            SHabiliteMenuItemPan(True, MmnuSuprimirQR)
        Else
            If MmnuSuprimirQR.IsEnabled Then
                SHabiliteMenuItemPan(False, MmnuSuprimirQR)
            End If
        End If
        MnuSuprimirQRC.IsEnabled = MmnuSuprimirQR.IsEnabled
    End Sub
    Private Sub SImporteCodigoQR()
        Dim lblnOk = False
        Dim lofdFoto As New OpenFileDialog With {
            .DefaultExt = ".jpg",
            .Filter = My.Resources.TipoArchivoImagen
        }
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            lblnOk = lofdFoto.ShowDialog
            If lblnOk Then
                ' Obtener el objeto Image de la Foto
                Dim lstmFoto As Stream = lofdFoto.OpenFile
                Dim llngTamano As Long = lstmFoto.Length
                If llngTamano > 70000 Then
                    lstrMens = "El tamaño máximo permitido de la imagen es de 68K. " &
                            "Debe importar una archivo más pequeño!"
                Else
                    MimgQR = Bitmap.FromStream(lstmFoto)
                    ' Asigno la imagen al control imgFoto de la ventana
                    Dim lstrTray As String = lofdFoto.FileName
                    Dim lbimFoto As New BitmapImage
                    lbimFoto.BeginInit()
                    lbimFoto.UriSource = New Uri(lstrTray)
                    lbimFoto.DecodePixelWidth = 135
                    lbimFoto.EndInit()
                    imgCodigoQR.Source = lbimFoto
                End If
            End If
            lblnNoHayError = True
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                If String.IsNullOrEmpty(lstrMens) AndAlso lblnOk Then
                    lstrMens = "El código QR fue asociado correctamente!"
                    MblnQRImportado = lblnNoHayError
                End If
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Private Sub SCargueCodigoQR()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Dim lmstImagenGuardada As MemoryStream = MobjObjetoWin.FmstImagenQR()
            If Not IsNothing(lmstImagenGuardada) Then
                Dim lbimFoto As New BitmapImage
                lbimFoto.BeginInit()
                lbimFoto.StreamSource = lmstImagenGuardada
                lbimFoto.EndInit()
                imgCodigoQR.Source = lbimFoto
            Else
                imgCodigoQR.Source = Nothing
            End If
            lblnNoHayError = True
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If Not lblnNoHayError Then
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
            MblnQRImportado = lblnNoHayError
        End Try
    End Sub
    Private Sub SSuprimaCodigoQR()
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            Dim lblnSuprimio = MobjObjetoWin.FblnSuprimioQR()
            SCargueCodigoQR()
            lblnNoHayError = True
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                lstrMens = "El código QR fue eliminado correctamente!"
                SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
#End Region
#Region "Busqueda"
    Protected Overrides Sub SBuscar()
        Me.Cursor = Cursors.Wait
        If IsNothing(HwinBusqueda) Then
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

    Private Sub SBusqueCuenta()
        SBuscar()
        If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
            txtIdCtaContabNueva.Text = StrResultadoBusqueda
            MobjObjetoWin.ObjIdCtaContabilidadStr.ObjValorPro = txtIdCtaContabNueva.Text
            SMuestreDatos()
        End If
    End Sub

    Protected Overrides Function FblnDefinioBusqueda() As Boolean
        SDefineCuentaCont()
        Return True
    End Function

    Private Sub SDefineCuentaCont()
        Dim lstrCamposMostrar As String() = {ClsIdCuentaContStr.SstrNombreCampoBd,
                                                 ClsNombreCuentaStr.SstrNombreCampoBd}
        Dim lstrCampoBusqueda As String = ClsNombreCuentaStr.SstrNombreCampoBd
        Dim lstrCampoRetornar As String = ClsIdCuentaContStr.SstrNombreCampoBd
        Dim lstrTabla As String = ClsCuentaContabilidad.SstrNombreTabla
        Dim lstrFiltro As String = ClsIdCarpetaCuentaShr.SstrNombreCampoBd & " = " & GshrIdCarpeta &
                " AND " & lstrCampoBusqueda & "<> ''"
        HwinBusqueda.SDefinaBusqueda("Nombre Cuenta", lstrTabla, lstrCamposMostrar,
                lstrCampoBusqueda, lstrCampoRetornar, lstrFiltro)
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SAdicioneCtrlsRestringidos()
        SAdicioneControlRestringido(dgrCuentasBanco)
        SAdicioneControlRestringido(txtNombreBanco)
        SAdicioneControlRestringido(txtNumeroCuenta)
        SAdicioneControlRestringido(txtIdCtaContab)
        SAdicioneControlRestringido(bttEncontrarCtaBnco)
    End Sub
    Private Sub SModifiqueBarraHerramientas()
        HbttAlPrimero.Visibility = Visibility.Collapsed
        HbttAlAnterior.Visibility = Visibility.Collapsed
        HbttAlSiguiente.Visibility = Visibility.Collapsed
        HbttAlUltimo.Visibility = Visibility.Collapsed
        HbttBuscar.Visibility = Visibility.Collapsed
        Dim ltlbMiToolBar As ToolBar = Nothing
        For Each lobjObjeto As Object In PanelControl.Children
            If TypeOf lobjObjeto Is ToolBar Then
                ltlbMiToolBar = lobjObjeto
                Exit For
            End If
        Next
        If Not IsNothing(ltlbMiToolBar) Then
            For Each lobjObjeto In ltlbMiToolBar.Items
                If TypeOf (lobjObjeto) Is Separator Then
                    Dim lsepSeparador As Separator = lobjObjeto
                    If lsepSeparador.Name = "sepNavegar" Then
                        lsepSeparador.Visibility = Visibility.Collapsed
                        Exit For
                    End If
                End If
            Next
        End If
        HmnuNavegar.Visibility = Visibility.Collapsed
    End Sub
    Private Sub SDeshabiliteControlesActuales()
        txtNombreBanco.Style = FindResource("RecCtlNoHabilitado")
        txtNumeroCuenta.Style = FindResource("RecCtlNoHabilitado")
        txtIdCtaContab.Style = FindResource("RecCtlNoHabilitado")
    End Sub
    ''' <summary>
    ''' Establece el ToolTip de los Controles a un valor constante
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SEstablezcaToolTipGral()
        HbttCrear.ToolTip = My.Resources.TTAdicionaModuloServicio
        HmnuCrear.ToolTip = My.Resources.TTAdicionaModuloServicio
    End Sub
    Private Sub SHabiliteControlesNuevos(ablnHabilite As Boolean)
        SVisibiliceControlesNuevos(ablnHabilite)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            txtNombreBancoNuevo.Style = FindResource("RecCtlHabilitado")
            txtNumeroCuentaNuevo.Style = FindResource("RecCtlHabilitado")
            txtIdCtaContabNueva.Style = FindResource("RecCtlHabilitado")
        End If
    End Sub
    Private Sub SVisibiliceControlesNuevos(ablnHabilite As Boolean)
        Dim lvisVisibilidadNuevos As Visibility
        Dim lvisVisibilidadActual As Visibility
        If ablnHabilite Then
            lvisVisibilidadNuevos = Visibility.Visible
            lvisVisibilidadActual = Visibility.Hidden
        Else
            lvisVisibilidadNuevos = Visibility.Hidden
            lvisVisibilidadActual = Visibility.Visible
        End If
        txtNombreBanco.Visibility = lvisVisibilidadActual
        txtNumeroCuenta.Visibility = lvisVisibilidadActual
        txtIdCtaContab.Visibility = lvisVisibilidadActual

        txtNombreBancoNuevo.Visibility = lvisVisibilidadNuevos
        txtNumeroCuentaNuevo.Visibility = lvisVisibilidadNuevos
        txtIdCtaContabNueva.Visibility = lvisVisibilidadNuevos
        bttEncontrarCtaBnco.Visibility = lvisVisibilidadNuevos
    End Sub
    Private Sub SCreeCuentaBanco()
        SHabiliteControlesNuevos(True)
        dgrCuentasBanco.IsEnabled = False
        SMuestreDatos()
    End Sub
    Private Sub SModifiqueCuentaBanco()
        If Not IsNothing(MobjObjetoWin) AndAlso MobjObjetoWin.BlnExiste Then
            dgrCuentasBanco.IsEnabled = False
            SHabiliteControlesNuevos(True)
            txtNombreBancoNuevo.Focus()
        End If
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Dim lbttBoton As Button = lelmElemento
            If lbttBoton.Name = "bttEncontrarCtaBnco" Then
                SBusqueCuenta()
            End If
        End If
    End Sub
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            Select Case lelmElemento.Name
                Case "MmnuAsociarQR", "MnuAsociarQRC"
                    SImporteCodigoQR()
                Case "MmnuSuprimirQR", "MmnuSuprimirQRC"
                    If MsgBox("Realmente desea suprimir el código QR?", vbYesNo,
                            "Eliminar código QR") = vbYes Then
                        SSuprimaCodigoQR()
                    End If
            End Select
        End If
    End Sub
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        ElseIf lelmElemento.Equals(HbttCancelar) Then
            If MblnDejoUltimoControl Then
                HbttAceptar.Focus()
            End If
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    MblnDejoUltimoControl = False
                    With MobjObjetoWin
                        Select Case lelmElemento.Name
                            Case "txtNombreBancoNuevo"
                                .ObjNombreBancoStr.ObjValorPro = txtNombreBancoNuevo.Text
                            Case "txtNumeroCuentaNuevo"
                                .ObjNumeroCuentaStr.ObjValorPro = txtNumeroCuentaNuevo.Text
                            Case "txtIdCtaContabNueva"
                                .ObjIdCtaContabilidadStr.ObjValorPro = txtIdCtaContabNueva.Text
                                MblnDejoUltimoControl = True
                        End Select
                    End With
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub
    Private Sub Txt_TextChanged(sender As Object, e As TextChangedEventArgs) Handles _
            txtNombreBanco.TextChanged, txtIdCtaContabNueva.TextChanged
        Select Case True
            Case sender.Equals(txtNombreBanco)
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                    If Not String.IsNullOrEmpty(txtNombreBanco.Text) Then
                        Dim lshrIdCtaBanco As Short = CType(txtIdCuenta.Content, Short)
                        Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, lshrIdCtaBanco}
                        MobjObjetoWin.SAbra(lobjValorLlave)
                    Else
                        MobjObjetoWin = Nothing
                    End If
                    SMuestreDatos()
                End If
            Case sender.Equals(txtIdCtaContabNueva)
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    StrResultadoBusqueda = String.Empty
                    If txtIdCtaContabNueva.Text.StartsWith("?") Then
                        SBuscar()
                    End If
                    If BlnBusquedaOk AndAlso Not String.IsNullOrEmpty(StrResultadoBusqueda) Then
                        txtIdCtaContabNueva.Text = StrResultadoBusqueda
                        MobjObjetoWin.ObjIdCtaContabilidadStr.ObjValorPro = txtIdCtaContabNueva.Text
                        SMuestreDatos()
                    End If
                End If
            Case sender.Equals(txtIdCtaContab)
                If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
                    If MobjObjetoWin.ObjIdCtaContabilidadStr.BlnEsValido Then
                        txtNomCtaContab.Content = MobjObjetoWin.ObjIdCtaContabilidadStr.StrNombreCuenta
                    Else
                        txtNomCtaContab.Content = String.Empty
                    End If
                End If
        End Select
    End Sub
    Private Sub ChkActiva_Click(sender As Object, e As RoutedEventArgs) Handles chkActiva.Click
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is CheckBox Then
            If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                MobjObjetoWin.ObjEstaActivaBln.ObjValorPro = chkActiva.IsChecked
            End If
            SHabiliteMenuesQR()
        End If
    End Sub
#End Region
End Class
