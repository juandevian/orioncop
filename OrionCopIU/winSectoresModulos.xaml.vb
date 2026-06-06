Imports System.ComponentModel
Public Class WinSectoresModulos
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuIdSector = 0
        enuTasaContr
    End Enum
#End Region
    ' Variables
    Private MnuAgregarTodos As MenuItem = Nothing
    Private MobjObjetoWin As ClsSectorModulo = Nothing
    Private ReadOnly MobjModulo As ClsModuloContribucion = Nothing
    Private MblnDejoUltimoControl As Boolean = False
    Private MblnPoblandoCombo As Boolean = False
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomSecMod
#End Region
#Region "Constructor"
    Friend Sub New(aobjModulo As ClsModuloContribucion)
        InitializeComponent()
        MobjModulo = aobjModulo
        HenuIdVentana = EnuIdVentanaDef.enuSectoresModulo
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolCamposLlaves As New Collection From {
            cboSectoresNuevo
        }
        SAdicioneControlRestringido(txtNombreSector)
        SAdicioneControlRestringido(txtTasa)
        SAdicioneControlRestringido(dgrSectoresModulo)
        SCargueForma(EnuElementosAdicionalesDef.None, 2,
                lcolCamposLlaves, txtTasaNueva, False)
        SPuebleBarraEstado(HcolLabelsBarraEstado)
        txtModulo.Content = MobjModulo.ObjNombreModuloStr.ObjValorPro
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
            ObjObjetoWin = New ClsSectorModulo(MobjModulo, EnuModoInstanciaObjDef.enuNavegable)
        End If
        MobjObjetoWin = ObjObjetoWin
        MobjObjetoWin.SVayaAlPrimero()
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuIdSector) = lblSectorModulo
        StcValidaControl(EnuValidEntrada.enuTasaContr) = lblTasa
        '
        SDeshabiliteControlesActuales()
        SModifiqueBarraHerramientas()

        grdSectoresModulo.DataContext = MobjModulo.FdtbSectoresModulo()
        SPuebleComboBoxes()
        ' 
        HbttAceptar.TabIndex = 15
        HbttCancelar.TabIndex = 16
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If dgrSectoresModulo.Items.Count = 0 AndAlso
                EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            SLevanteEveNoti("No hay Sectores del Modulos de Contribución para ser mostrados!", "", 0,
                                EnuSeveridadNot.EnuInformacion)
            Exit Sub
        End If
        If Not IsNothing(MobjObjetoWin) Then
            cboSectoresNuevo.SelectedIndex = MobjObjetoWin.ObjIdSector_SectorModuloShr.ObjValorPro
            txtTasaNueva.Text = Format(MobjObjetoWin.ObjTasaContribucionDbl.ObjValorPro, "p")
        End If
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        If dgrSectoresModulo.Items.Count = 0 AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentana.cenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntrada.enuIdSector) = .ObjIdSector_SectorModuloShr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuTasaContr) = .ObjTasaContribucionDbl.BlnEsValido
            End With
        End If
        SHabiliteBotonesTlb()
        SHabiliteMenuItem(HmnuCrear.IsEnabled, MnuAgregarTodos)
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjIdSector_SectorModuloShr.ObjValorPro = cboSectoresNuevo.SelectedIndex
            .ObjTasaContribucionDbl.ObjValorPro = FdblTasa(txtTasaNueva.Text)
        End With
        SValide()
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        HmnuCrear.Header = My.Resources.MnuAgregarSectorModulo
        HmnuCrear.ToolTip = My.Resources.TTAgreSecMod
        HbttCrear.ToolTip = My.Resources.TTAgreSecMod
        MnuAgregarTodos = FmnuiMenuItem("MnuAgregarTodos", "Agregar _Todos los Sectores", "RecMnuItemSec")
        MnuAgregarTodos.ToolTip = My.Resources.TTAgreTodosSecMod
        HmnuAcciones.Items.Insert(1, MnuAgregarTodos)
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    ''' <summary>
    ''' Sub que prepara a la ventana y a su objeto para crear un nuevo objeto. Invalida el Sub
    ''' "SCree" de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SCree()
        If CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuCrear, Boolean) Then
            MyBase.SCree()
            SAdicioneSector()
        End If
    End Sub
    ''' <summary>
    ''' Prepara la ventana y su objeto para modificar el objeto. Invalida la función "SModifique"
    ''' de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SModifique()
        If CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuModificar, Boolean) Then
            cboSectoresNuevo.SelectedIndex = MobjObjetoWin.ObjIdSector_SectorModuloShr.ObjValorPro
            MyBase.SModifique()
            If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                SModifiqueSectorModulo()
            End If
        End If
    End Sub
    Protected Overrides Sub SSuprima()
        If CType(EnuTipoPermisoObjWin And EnuPermisosDef.enuSuprimir, Boolean) Then
            MyBase.SSuprima()
            SRefresqueWin()
        End If
    End Sub
    Protected Overrides Sub SRefresqueWin()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Mouse.OverrideCursor = Cursors.Wait
            txtNombreSector.Text = String.Empty
            MobjModulo.SRefresqueObj()
            grdSectoresModulo.DataContext = MobjModulo.FdtbSectoresModulo
            SOrdeneDataGrid(dgrSectoresModulo, dgrSectoresModulo.Columns(1),
                            ClsIdSector_SectorModuloShr.SstrNombreCampoBd,
                            ListSortDirection.Ascending)
            If dgrSectoresModulo.Items.Count > 0 Then
                dgrSectoresModulo.SelectedIndex = 0
            End If
            MobjObjetoWin.SRefresqueObj()
            SMuestreDatos()
            Mouse.OverrideCursor = Cursors.Arrow
        End If
    End Sub
    Protected Overrides Sub SEstablezcaWinConsultando()
        MyBase.SEstablezcaWinConsultando()
        SVisibiliceControlesNuevos(False)
        dgrSectoresModulo.IsEnabled = True
        SRefrescarClic()
        If dgrSectoresModulo.Items.Count > 0 Then
            dgrSectoresModulo.SelectedIndex = 0
        End If
    End Sub
#End Region
#Region "Procedimientos Propios"
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
        txtNombreSector.Style = FindResource("RecCtlNoHabilitado")
        txtTasa.Style = FindResource("RecCtlNoHabilitado")
    End Sub
    Private Sub SHabiliteControlesNuevos(ablnHabilite As Boolean)
        SVisibiliceControlesNuevos(ablnHabilite)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            cboSectoresNuevo.Style = FindResource("RecCtlHabilitado")
            txtTasaNueva.Style = FindResource("RecCtlHabilitado")
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
        txtNombreSector.Visibility = lvisVisibilidadActual
        txtTasa.Visibility = lvisVisibilidadActual
        cboSectoresNuevo.Visibility = lvisVisibilidadNuevos
        txtTasaNueva.Visibility = lvisVisibilidadNuevos
    End Sub
    Private Sub SAdicioneSector()
        SHabiliteControlesNuevos(True)
        dgrSectoresModulo.IsEnabled = False
        SMuestreDatos()
        cboSectoresNuevo.Focus()
    End Sub
    Private Sub SModifiqueSectorModulo()
        txtTasaNueva.Text = txtTasa.Text
        If Not IsNothing(MobjObjetoWin) AndAlso MobjObjetoWin.BlnExiste Then
            SHabiliteControlesNuevos(True)
            dgrSectoresModulo.IsEnabled = False
            txtTasaNueva.Focus()
        End If
    End Sub
    Private Sub SPuebleComboBoxes()
        Dim lcolSectores As Collection = GobjParametros.ColSectores
        MblnPoblandoCombo = True
        cboSectoresNuevo.Items.Add("Ninguno")
        For Each lobjSector As ClsSector In lcolSectores
            cboSectoresNuevo.Items.Add(lobjSector.ObjNombreSectorStr.ObjValorPro)
        Next
        MblnPoblandoCombo = False
        cboSectoresNuevo.SelectedIndex = 0
    End Sub
    Private Sub SAgregueTodosSectores()
        Mouse.OverrideCursor = Cursors.Wait
        MobjModulo.SAgregueTodosSectores()
        SFinaliceOperacion()
        SRefrescarClic()
        SValide()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
    Private Sub SAbraSectorModulo(astrNombreSector As String)
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If Not IsNothing(astrNombreSector) AndAlso Not String.IsNullOrEmpty(astrNombreSector) Then
                Dim lshrIdSector As Short = GobjParametros.FshrIdSector(astrNombreSector)
                Dim lobjValorLlave() = {GshrIdCarpeta, GshrIdCentroUtil, MobjModulo.ObjIdModuloShr.ObjValorPro,
                                        lshrIdSector}
                MobjObjetoWin.SAbra(lobjValorLlave)
            Else
                MobjObjetoWin.SVacie()
            End If
            SValide()
        End If
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            If lelmElemento.Equals(MnuAgregarTodos) Then
                SAgregueTodosSectores()
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
            End If
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is ComboBox OrElse TypeOf lelmElemento Is TextBox Then
                MblnDejoUltimoControl = False
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                    Try
                        With MobjObjetoWin
                            Select Case True
                                Case lelmElemento.Equals(cboSectoresNuevo)
                                    .ObjIdSector_SectorModuloShr.ObjValorPro = cboSectoresNuevo.SelectedIndex
                                Case lelmElemento.Equals(txtTasaNueva)
                                    .ObjTasaContribucionDbl.ObjValorPro = FdblTasa(txtTasaNueva.Text)
                                    MblnDejoUltimoControl = True
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
                        If Not lblnNoHayError Then
                            SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                        End If
                    End Try
                End If
            Else
                SMuestreDatos()
            End If
        End If
    End Sub
    Private Sub TxtTasaNueva_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtTasaNueva.TextChanged
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            MobjObjetoWin.ObjTasaContribucionDbl.ObjValorPro = FdblTasa(txtTasaNueva.Text)
        End If
    End Sub
    Private Sub Cbo_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cboSectoresNuevo.SelectionChanged
        If Not MblnPoblandoCombo Then
            If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                MobjObjetoWin.ObjIdSector_SectorModuloShr.ObjValorPro = cboSectoresNuevo.SelectedIndex
                SMuestreDatos()
            End If
        End If
    End Sub
    Private Sub Dgr_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dgrSectoresModulo.SelectionChanged
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            ldrvFilaActual = ldgrActual.SelectedItem
            If Not IsNothing(ldrvFilaActual) Then
                Dim lstrNombreSector = ldrvFilaActual("NombreSector")
                If MobjObjetoWin.BlnExiste AndAlso Not IsNothing(lstrNombreSector) Then
                    SAbraSectorModulo(lstrNombreSector)
                End If
                ObjObjetoWin = MobjObjetoWin
            End If
        End If
    End Sub
#End Region
End Class