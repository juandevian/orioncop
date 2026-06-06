Imports System.ComponentModel
Public Class WinModulosContribucion
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuNombreModulo
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsModuloContribucion = Nothing
    Private MblnDejoUltimoControl As Boolean = False
    Private MnuAbrirSectores As MenuItemPan = Nothing
    Private MnuAbrirSectoresC As MenuItem = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomModCont
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuModulosContr
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolCamposLlave As New Collection From {
            txtIdModulo
        }
        SAdicioneControlRestringido(dgrModulos)
        SAdicioneControlRestringido(bttAbrirSectores)
        SCargueForma(EnuElementosAdicionalesDef.None, 1,
                lcolCamposLlave, txtNombreModuloNuevo, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        SHabiliteControlesNuevos(False)
        dgrModulos.Focus()
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
            ObjObjetoWin = New ClsModuloContribucion(EnuModoInstanciaObjDef.enuNavegable)
        End If
        MobjObjetoWin = ObjObjetoWin
        MobjObjetoWin.SVayaAlPrimero()
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuNombreModulo) = lblNombreModulo
        '
        SEstablezcaToolTipGral()
        SEstablezcaDataContext()
        '
        HbttAceptar.TabIndex = 13
        HbttCancelar.TabIndex = 14
        SModifiqueBarraHerramientas()
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If Not IsNothing(MobjObjetoWin) Then
            With MobjObjetoWin
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    txtIdModuloNuevo.Text = .ObjIdModuloShr.ToString
                    txtNombreModuloNuevo.Text = .ObjNombreModuloStr.ObjValorPro
                    chkContriCuotaAdminNuevo.IsChecked = .ObjContribuyeCuotaAdminBln.ObjValorPro
                Else
                    If GobjParametros.FdtbModulos().Rows.Count = 0 Then
                        SLevanteEveNoti("No hay Modulos de Contribución para ser mostrados!", "", 0,
                                EnuSeveridadNot.EnuInformacion)
                    End If
                End If
            End With
        End If
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        With MobjObjetoWin
            If GobjParametros.FdtbModulos.Rows.Count = 0 AndAlso EnuOperacionEnWin =
                    EnuOperacionEnVentana.cenuConsultando Then
                SInicialiceValido()
            Else
                StcValidValido(EnuValidEntrada.enuNombreModulo) = .ObjNombreModuloStr.BlnEsValido
            End If
        End With
        SHabiliteBotonesTlb()
        If FblnEstanTodosBien() Then
            SHabiliteAccesoSectores(EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando)
        Else
            SHabiliteAccesoSectores(False)
        End If
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjIdModuloShr.ObjValorPro = txtIdModuloNuevo.Text
            .ObjNombreModuloStr.ObjValorPro = txtNombreModuloNuevo.Text
            .ObjContribuyeCuotaAdminBln.ObjValorPro = chkContriCuotaAdminNuevo.IsChecked
        End With
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        MnuAbrirSectores = FmnuiMenuItemPan("MnuAbrirSectores", "_Abrir Sectores Contribuyentes", 1, "")
        Dim lsep As New Separator
        Dim lentIndice = HmnuAcciones.Items.Count - 1
        HmnuAcciones.Items.Insert(lentIndice, lsep)
        HmnuAcciones.Items.Insert(lentIndice, MnuAbrirSectores)
        dgrModulos.ContextMenu = FindResource("RecMnuSectoresMC")
        SAsigneMenuContexModulo()
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SHabiliteMenues()
        MyBase.SHabiliteMenues()
        bttAbrirSectores.IsEnabled = MnuAbrirSectores.IsEnabled
        SHabiliteMenuItem(MnuAbrirSectores.IsEnabled, MnuAbrirSectoresC)
    End Sub
    ''' <summary>
    ''' Sub que prepara a la ventana y a su objeto para crear un nuevo objeto. Invalida el Sub
    ''' "SCree" de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SCree()
        MyBase.SCree()
        SCreeModulo()
    End Sub
    ''' <summary>
    ''' Prepara la ventana y su objeto para modificar el objeto. Invalida la función "SModifique"
    ''' de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SModifique()
        txtIdModuloNuevo.Text = MobjObjetoWin.ObjIdModuloShr.ToString()
        txtNombreModuloNuevo.Text = MobjObjetoWin.ObjNombreModuloStr.ToString()
        chkContriCuotaAdminNuevo.IsChecked = MobjObjetoWin.ObjContribuyeCuotaAdminBln.ObjValorPro
        MyBase.SModifique()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            SModifiqueModulo()
        Else
            SHabiliteAccesoSectores(False)
            SHabiliteControlesNuevos(True)
            dgrModulos.IsEnabled = False
            SMuestreDatos()
        End If
    End Sub
    Protected Overrides Sub SRefresqueWin()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            GobjPanDat.SControleProcesoObj(True)
            txtNombreModulo.Text = String.Empty
            chkContriCuotaAdmin.IsChecked = False
            MyBase.SRefresqueWin()
            SEstablezcaDataContext()
            GobjPanDat.SControleProcesoObj(False)
            Mouse.OverrideCursor = Cursors.Arrow
        End If
    End Sub
    Protected Overrides Sub SEstablezcaWinConsultando()
        MyBase.SEstablezcaWinConsultando()
        SRefresqueWin()
        dgrModulos.IsEnabled = True
        SHabiliteControlesNuevos(False)
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SHabiliteAccesoSectores(ablnHabilite As Boolean)
        Dim lblnHabilite = FblnHabilitarMenuPan(MnuAbrirSectores.EntIdAccion)
        lblnHabilite = lblnHabilite AndAlso ablnHabilite
        SHabiliteMenuItem(lblnHabilite, MnuAbrirSectores)
        bttAbrirSectores.IsEnabled = MnuAbrirSectores.IsEnabled
        SHabiliteMenuItem(MnuAbrirSectores.IsEnabled, MnuAbrirSectoresC)
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
                    If lobjObjeto.Name = "sepNavegar" Then
                        lobjObjeto.Visibility = Visibility.Collapsed
                        Exit For
                    End If
                End If
            Next
        End If
        HmnuNavegar.Visibility = Visibility.Collapsed
    End Sub
    Private Sub SHabiliteControlesNuevos(ablnHabilite As Boolean)
        SVisibiliceControlesNuevos(ablnHabilite)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            txtIdModuloNuevo.Style = FindResource("RecCtlNoHabilitado")
            txtNombreModuloNuevo.Style = FindResource("RecCtlHabilitado")
            chkContriCuotaAdminNuevo.Style = FindResource("RecCtlHabilitado")
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
        txtIdModulo.Visibility = lvisVisibilidadActual
        txtNombreModulo.Visibility = lvisVisibilidadActual
        chkContriCuotaAdmin.Visibility = lvisVisibilidadActual
        txtIdModuloNuevo.Visibility = lvisVisibilidadNuevos
        txtNombreModuloNuevo.Visibility = lvisVisibilidadNuevos
        chkContriCuotaAdminNuevo.Visibility = lvisVisibilidadNuevos
    End Sub
    Private Sub SEstablezcaToolTipGral()
        Dim lstrMens = "Estando ésta Ventana en modo 'Consultando', puede abrir la Ventana de los Sectores " & vbCrLf &
                          "que contribuyen a un Módulo, dando doble click sobre la fila del Módulo deseado!"
        dgrModulos.ToolTip = lstrMens
    End Sub
    Private Sub SCreeModulo()
        SHabiliteControlesNuevos(True)
        dgrModulos.IsEnabled = False
        MobjObjetoWin.ObjContribuyeCuotaAdminBln.SValide()
        SMuestreDatos()
        txtNombreModuloNuevo.Focus()
    End Sub
    Private Sub SModifiqueModulo()
        If Not IsNothing(MobjObjetoWin) AndAlso MobjObjetoWin.BlnExiste Then
            SHabiliteControlesNuevos(True)
            dgrModulos.IsEnabled = False
            txtNombreModuloNuevo.Focus()
        End If
    End Sub
    Private Sub SAbraSectores()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Dim lwinSectores As New WinSectoresModulos(MobjObjetoWin) With {
                .WinPadre = Me
            }
            lwinSectores.ShowDialog()
            MobjObjetoWin.SRefresqueObj()
        End If
    End Sub
    Private Sub SAsigneMenuContexModulo()
        Dim lmnuMenuContextual As ContextMenu = FindResource("RecMnuSectoresMC")
        For Each lobjObjetoMenu As Object In lmnuMenuContextual.Items
            If TypeOf lobjObjetoMenu Is MenuItem Then
                Dim lmnuItemMenuCont As MenuItem = lobjObjetoMenu
                If lmnuItemMenuCont.Name = "MnuAbrirSectoresC" Then
                    MnuAbrirSectoresC = lmnuItemMenuCont
                    Exit For
                End If
            End If
        Next
    End Sub
    Private Sub SEstablezcaDataContext()
        grdModulos.DataContext = GobjParametros.FdtbModulos()
        SOrdeneDataGrid(dgrModulos, dgrModulos.Columns(0), ClsIdModuloShr.SstrNombreCampoBd,
                ListSortDirection.Ascending)
        If dgrModulos.Items.Count > 0 Then
            dgrModulos.SelectedIndex = 0
        End If
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            If lelmElemento.Equals(bttAbrirSectores) Then
                SAbraSectores()
            End If
        End If
    End Sub
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            If lelmElemento.Equals(MnuAbrirSectores) Then
                SAbraSectores()
            End If
        End If
    End Sub
    Private Sub MnuContextual_Click(sender As Object, e As RoutedEventArgs)
        If TypeOf sender Is MenuItem Then
            Dim lmnuConte As MenuItem = sender
            If lmnuConte.Name = "MnuAbrirSectoresC" Then
                SAbraSectores()
            End If
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
                MblnDejoUltimoControl = False
            End If
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando Then
            Dim lelmelemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmelemento Is TextBox OrElse TypeOf lelmelemento Is CheckBox Then
                MblnDejoUltimoControl = False
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    With MobjObjetoWin
                        Select Case lelmelemento.Name
                            Case "txtIdModuloNuevo"
                                .ObjIdModuloShr.ObjValorPro = txtIdModuloNuevo.Text
                            Case "txtNombreModuloNuevo"
                                .ObjNombreModuloStr.ObjValorPro = txtNombreModuloNuevo.Text
                            Case "chkContriCuotaAdminNuevo"
                                .ObjContribuyeCuotaAdminBln.ObjValorPro = chkContriCuotaAdminNuevo.IsChecked
                                MblnDejoUltimoControl = True
                        End Select
                    End With
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub
    Private Sub Dgr_MouseRightButtonUp(sender As Object, e As MouseButtonEventArgs) Handles _
                dgrModulos.MouseRightButtonUp
        If Not IsNothing(dgrModulos.SelectedItem) Then
            SAbraSectores()
        End If
    End Sub
    Private Sub TxtNombre_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtNombreModulo.TextChanged
        Dim lobjValorLlave() As Object = {GshrIdCarpeta, GshrIdCentroUtil, txtIdModulo.Text}
        MobjObjetoWin.SAbra(lobjValorLlave)
        SMuestreDatos()
    End Sub
    Private Sub TxtNombreModuloNuevo_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtNombreModuloNuevo.TextChanged
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            MobjObjetoWin.ObjNombreModuloStr.ObjValorPro = txtNombreModuloNuevo.Text
        End If
    End Sub
    Private Sub Chk_Click(sender As Object, e As RoutedEventArgs) Handles chkContriCuotaAdminNuevo.Click
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            If Not IsNothing(MobjObjetoWin) Then
                MobjObjetoWin.ObjContribuyeCuotaAdminBln.ObjValorPro = chkContriCuotaAdminNuevo.IsChecked
            End If
        End If
    End Sub
#End Region
End Class
