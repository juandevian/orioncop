Imports System.Windows.Controls
Imports System.ComponentModel
Public Class WinUbicacion
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuIdPais = 0
        enuNomPais
        enuIdDpto
        enuNomDpto
        enuIdCiudad
        enuNomCiudad
    End Enum

    Private Enum EnuTipoObjeto As Integer
        enuNinguno = 0
        enuBarrio
        enuCiudad
        enuDpto
        enuPais
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsUbicacion = Nothing
    Private MobjPais As ClsPais = Nothing
    Private MobjDpto As ClsDepartamento = Nothing
    Private MobjCiudad As ClsCiudad = Nothing
    Private MnuAccionesPais As MenuItem = Nothing
    Private MnuCrearPais As MenuItem = Nothing
    Private MnuModificarPais As MenuItem = Nothing
    Private MnuEliminarPais As MenuItem = Nothing
    Private MnuAccionesDpto As MenuItem = Nothing
    Private MnuCrearDpto As MenuItem = Nothing
    Private MnuModificarDpto As MenuItem = Nothing
    Private MnuEliminarDpto As MenuItem = Nothing
    Private MnuAccionesCiudad As MenuItem = Nothing
    Private MnuCrearCiudad As MenuItem = Nothing
    Private MnuModificarCiudad As MenuItem = Nothing
    Private MnuEliminarCiudad As MenuItem = Nothing

    Private MstrIdPais As String = String.Empty, mbytIdDpto As Byte = 0, mshrIdCiudad As Short = 0, mshrIdBarrio As Short = 0
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomUbica
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuUbicaciones
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 6,
                Nothing, Nothing, False)
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
            ObjObjetoWin = New ClsUbicacion
        End If
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuIdPais) = lblIdPais
        StcValidaControl(EnuValidEntrada.enuNomPais) = lblNombrePais
        StcValidaControl(EnuValidEntrada.enuIdDpto) = lblIdDpto
        StcValidaControl(EnuValidEntrada.enuNomDpto) = lblNombreDpto
        StcValidaControl(EnuValidEntrada.enuIdCiudad) = lblIdCiudad
        StcValidaControl(EnuValidEntrada.enuNomCiudad) = lblNombreCiudad
        '
        tbiPais.DataContext = MobjObjetoWin.DtbPaises
        SHabiliteDataGrids(True)
        dgrPaises.IsReadOnly = True
        dgrDptos.IsReadOnly = True
        dgrCiudades.IsReadOnly = True

        HbttAceptar.TabIndex = 44
        HbttCancelar.TabIndex = 45
    End Sub
    Protected Overrides Sub SMuestreDatos()
        Select Case tbcUbicacion.SelectedIndex
            Case 0
                If Not IsNothing(MobjPais) Then
                    txtIdPaisNuevo.Text = MobjPais.ObjIdPaisStr.ToString
                    txtNombrePaisNuevo.Text = MobjPais.ObjNombrePaisStr.ObjValorPro
                Else
                    txtIdPaisNuevo.Text = String.Empty
                    txtNombrePaisNuevo.Text = String.Empty
                End If
            Case 1
                If Not IsNothing(MobjDpto) Then
                    txtIdDptoNuevo.Text = MobjDpto.ObjIdDptoByt.ToString
                    txtNombreDptoNuevo.Text = MobjDpto.ObjNombreDptoStr.ObjValorPro
                Else
                    txtIdDptoNuevo.Text = String.Empty
                    txtNombreDptoNuevo.Text = String.Empty
                End If
            Case 2
                If Not IsNothing(MobjCiudad) Then
                    txtIdCiudadNueva.Text = MobjCiudad.ObjIdCiudadShr.ToString
                    txtNombreCiudadNueva.Text = MobjCiudad.ObjNombreCiudadStr.ObjValorPro
                Else
                    txtIdCiudadNueva.Text = String.Empty
                    txtNombreCiudadNueva.Text = String.Empty
                End If
        End Select
        txtPais.Content = txtIdPais.Text & " - " & txtNombrePais.Text
        txtDpto.Content = txtIdDpto.Text & " - " & txtNombreDpto.Text
        txtCiudad.Content = txtIdCiudad.Text & " - " & txtNombreCiudad.Text
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        StcValidValido(EnuValidEntrada.enuIdPais) = True
        StcValidValido(EnuValidEntrada.enuNomPais) = True
        StcValidValido(EnuValidEntrada.enuIdDpto) = True
        StcValidValido(EnuValidEntrada.enuNomDpto) = True
        StcValidValido(EnuValidEntrada.enuIdCiudad) = True
        StcValidValido(EnuValidEntrada.enuNomCiudad) = True
        If MobjObjetoWin.EnuTipoAccion <> EnuTipoAccionDef.None Then
            Select Case tbcUbicacion.SelectedIndex
                Case 0
                    With MobjPais
                        StcValidValido(EnuValidEntrada.enuIdPais) = .ObjIdPaisStr.BlnEsValido
                        StcValidValido(EnuValidEntrada.enuNomPais) = .ObjNombrePaisStr.BlnEsValido
                    End With
                Case 1
                    With MobjDpto
                        StcValidValido(EnuValidEntrada.enuIdDpto) = .ObjIdDptoByt.BlnEsValido
                        StcValidValido(EnuValidEntrada.enuNomDpto) = .ObjNombreDptoStr.BlnEsValido
                    End With
                Case 2
                    With MobjCiudad
                        StcValidValido(EnuValidEntrada.enuIdCiudad) = .ObjIdCiudadShr.BlnEsValido
                        StcValidValido(EnuValidEntrada.enuNomCiudad) = .ObjNombreCiudadStr.BlnEsValido
                    End With
            End Select
        End If
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        Select Case tbcUbicacion.SelectedIndex
            Case 0
                If MobjPais IsNot Nothing Then
                    MobjPais.ObjIdPaisStr.ObjValorPro = txtIdPaisNuevo.Text
                    MobjPais.ObjNombrePaisStr.ObjValorPro = txtNombrePaisNuevo.Text
                End If
            Case 1
                If MobjDpto IsNot Nothing Then
                    MobjDpto.ObjIdDptoByt.ObjValorPro = txtIdDptoNuevo.Text
                    MobjDpto.ObjNombreDptoStr.ObjValorPro = txtNombreDptoNuevo.Text
                End If
            Case 2
                If MobjCiudad IsNot Nothing Then
                    MobjCiudad.ObjIdCiudadShr.ObjValorPro = txtIdCiudadNueva.Text
                    MobjCiudad.ObjNombreCiudadStr.ObjValorPro = txtNombreCiudadNueva.Text
                End If
        End Select
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        For Each lmnuiItem As MenuItem In HmnuMiMenu.Items
            If lmnuiItem.Name = "MnuNavegar" Then
                lmnuiItem.Visibility = Visibility.Collapsed
                Exit For
            End If
        Next
        'Paises
        MnuAccionesPais = FmnuiMenuItem("MnuAccionesPais", "Acciones _Pais", "RecMnuItemPri")
        MnuCrearPais = FmnuiMenuItem("MnuCrearPais", "_Crear País", "RecMnuItemSec")
        MnuModificarPais = FmnuiMenuItem("MnuModificarPais", "_Modificar País", "RecMnuItemSec")
        MnuEliminarPais = FmnuiMenuItem("MnuEliminarPais", "_Eliminiar País", "RecMnuItemSec")
        MnuAccionesPais.Items.Add(MnuCrearPais)
        MnuAccionesPais.Items.Add(MnuModificarPais)
        MnuAccionesPais.Items.Add(MnuEliminarPais)
        MnuAccionesPais.Style = TryFindResource("RecMnuItemPriDes")
        ' Dptos
        MnuAccionesDpto = FmnuiMenuItem("MnuAccionesDpto", "Acciones _Departamento", "RecMnuItemPri")
        MnuCrearDpto = FmnuiMenuItem("MnuCrearDpto", "_Crear Departamento", "RecMnuItemSec")
        MnuModificarDpto = FmnuiMenuItem("MnuModificarDpto", "_Modificar Departamento", "RecMnuItemSec")
        MnuEliminarDpto = FmnuiMenuItem("MnuEliminarDpto", "_Eliminiar Departamento", "RecMnuItemSec")
        MnuAccionesDpto.Items.Add(MnuCrearDpto)
        MnuAccionesDpto.Items.Add(MnuModificarDpto)
        MnuAccionesDpto.Items.Add(MnuEliminarDpto)
        MnuAccionesDpto.Style = TryFindResource("RecMnuItemPriDes")
        ' Ciudades
        MnuAccionesCiudad = FmnuiMenuItem("MnuAccionesCiudad", "Acciones _Ciudad", "RecMnuItemPri")
        MnuCrearCiudad = FmnuiMenuItem("MnuCrearCiudad", "_Crear Ciudad", "RecMnuItemSec")
        MnuModificarCiudad = FmnuiMenuItem("MnuModificarCiudad", "_Modificar Ciudad", "RecMnuItemSec")
        MnuEliminarCiudad = FmnuiMenuItem("MnuEliminarCiudad", "_Eliminiar Ciudad", "RecMnuItemSec")
        MnuAccionesCiudad.Items.Add(MnuCrearCiudad)
        MnuAccionesCiudad.Items.Add(MnuModificarCiudad)
        MnuAccionesCiudad.Items.Add(MnuEliminarCiudad)
        MnuAccionesCiudad.Style = TryFindResource("RecMnuItemPriDes")
        ' Adicionar a mi menu
        HmnuMiMenu.Items.Insert(1, MnuAccionesCiudad)
        HmnuMiMenu.Items.Insert(1, MnuAccionesDpto)
        HmnuMiMenu.Items.Insert(1, MnuAccionesPais)
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SHabiliteMenues()
        Dim lblnTienePermiso = False
        MyBase.SHabiliteMenues()
        With GobjPanorama.ObjUsuarioActual
            ' Tablas Auxiliares
            lblnTienePermiso = .FblnTienePermiso(EnuIdClasesPanDef.enuPais, EnuIdAccionDef.enuCrear)
            SHabiliteMenuItem(lblnTienePermiso, MnuCrearPais)
            lblnTienePermiso = .FblnTienePermiso(EnuIdClasesPanDef.enuPais, EnuIdAccionDef.enuModificar)
            SHabiliteMenuItem(lblnTienePermiso, MnuModificarPais)
            lblnTienePermiso = .FblnTienePermiso(EnuIdClasesPanDef.enuPais, EnuIdAccionDef.enuSuprimir)
            SHabiliteMenuItem(lblnTienePermiso, MnuEliminarPais)

            lblnTienePermiso = .FblnTienePermiso(EnuIdClasesPanDef.enuDepartamento, EnuIdAccionDef.enuCrear)
            SHabiliteMenuItem(lblnTienePermiso, MnuCrearDpto)
            lblnTienePermiso = .FblnTienePermiso(EnuIdClasesPanDef.enuDepartamento, EnuIdAccionDef.enuModificar)
            SHabiliteMenuItem(lblnTienePermiso, MnuModificarDpto)
            lblnTienePermiso = .FblnTienePermiso(EnuIdClasesPanDef.enuDepartamento, EnuIdAccionDef.enuSuprimir)
            SHabiliteMenuItem(lblnTienePermiso, MnuEliminarDpto)

            lblnTienePermiso = .FblnTienePermiso(EnuIdClasesPanDef.enuCiudad, EnuIdAccionDef.enuCrear)
            SHabiliteMenuItem(lblnTienePermiso, MnuCrearCiudad)
            lblnTienePermiso = .FblnTienePermiso(EnuIdClasesPanDef.enuCiudad, EnuIdAccionDef.enuModificar)
            SHabiliteMenuItem(lblnTienePermiso, MnuModificarCiudad)
            lblnTienePermiso = .FblnTienePermiso(EnuIdClasesPanDef.enuCiudad, EnuIdAccionDef.enuSuprimir)
            SHabiliteMenuItem(lblnTienePermiso, MnuEliminarCiudad)
        End With
    End Sub
    ''' <summary>
    ''' Prepara la ventana y su objeto para modificar el objeto. Invalida la función "SModifique"
    ''' de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SModifique()
        MyBase.SModifique()
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            SDeshabiliteControlesActuales()
            SHabiliteMenuesUbicacion()
        End If
    End Sub
    Protected Overrides Sub SFinaliceOperacion()
        MyBase.SFinaliceOperacion()
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            SHabiliteDataGrids(True)
            SHabiliteTabItems(True)
            SHabiliteMenuesUbicacion()
            SAsigneVariablesObjetos()
            txtEstadoWin.Content = "Consultando"
            Select Case tbcUbicacion.SelectedIndex
                Case 0
                    txtIdPais.Visibility = Visibility.Visible
                    txtIdPaisNuevo.Visibility = Visibility.Hidden
                    txtNombrePais.Visibility = Visibility.Visible
                    txtNombrePaisNuevo.Visibility = Visibility.Hidden
                    tbiPais.DataContext = MobjObjetoWin.DtbPaises()
                    If String.IsNullOrEmpty(txtIdPais.Text) AndAlso dgrPaises.Items.Count > 0 Then
                        dgrPaises.SelectedIndex = 0
                    End If
                    SDetermineObjPais()
                Case 1
                    txtIdDpto.Visibility = Visibility.Visible
                    txtIdDptoNuevo.Visibility = Visibility.Hidden
                    txtNombreDpto.Visibility = Visibility.Visible
                    txtNombreDptoNuevo.Visibility = Visibility.Hidden
                    tbiDpto.DataContext = MobjObjetoWin.DtbDptos(MstrIdPais)
                    If String.IsNullOrEmpty(txtIdDpto.Text) AndAlso dgrDptos.Items.Count > 0 Then
                        dgrDptos.SelectedIndex = 0
                    End If
                Case 2
                    txtIdCiudad.Visibility = Visibility.Visible
                    txtIdCiudadNueva.Visibility = Visibility.Hidden
                    txtNombreCiudad.Visibility = Visibility.Visible
                    txtNombreCiudadNueva.Visibility = Visibility.Hidden
                    tbiCiudad.DataContext = MobjObjetoWin.DtbCiudades(MstrIdPais, mbytIdDpto)
                    If String.IsNullOrEmpty(txtIdCiudad.Text) AndAlso dgrCiudades.Items.Count > 0 Then
                        dgrCiudades.SelectedIndex = 0
                    End If
            End Select
            SOrdeneGrid()
            SMuestreDatos()
        End If
    End Sub
    Protected Overrides Sub SRefresqueWin()
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            Mouse.OverrideCursor = Cursors.Wait
            Select Case tbcUbicacion.SelectedIndex
                Case 0
                    txtNombrePais.Text = String.Empty
                    If MobjObjetoWin IsNot Nothing Then
                        MobjObjetoWin.SRefresqueObj()
                    End If
                Case 1
                    txtNombreDpto.Text = String.Empty
                    If MobjPais IsNot Nothing Then
                        MobjPais.SRefresqueObj()
                    End If
                Case 2
                    txtNombreCiudad.Text = String.Empty
                    If MobjDpto IsNot Nothing Then
                        MobjDpto.SRefresqueObj()
                    End If
            End Select
            SFinaliceOperacion()
            SMuestreDatos()
            Mouse.OverrideCursor = Cursors.Arrow
        End If
    End Sub
    Protected Overrides Sub SHabiliteBotonesTlb()
        MyBase.SHabiliteBotonesTlb()
        If CType(MobjObjetoWin.EnuPermisosObj And EnuPermisosDef.enuModificar, Boolean) Then
            SHabiliteBotonTlb(True, HbttModificar)
            SHabiliteMenuItem(True, HmnuModificar)
        End If
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SHabiliteDataGrids(ablnHabilite As Boolean)
        dgrPaises.IsEnabled = ablnHabilite
        dgrDptos.IsEnabled = ablnHabilite
        dgrCiudades.IsEnabled = ablnHabilite
    End Sub
    Private Sub SHabiliteTabItems(ablnHabilite As Boolean)
        If ablnHabilite Then
            For Each objTabItem As Object In tbcUbicacion.Items
                If TypeOf objTabItem Is TabItem Then
                    objTabItem.IsEnabled = True
                End If
            Next
        Else
            Select Case tbcUbicacion.SelectedIndex
                Case 0
                    tbiPais.IsEnabled = True
                    tbiDpto.IsEnabled = False
                    tbiCiudad.IsEnabled = False
                Case 1
                    tbiPais.IsEnabled = False
                    tbiDpto.IsEnabled = True
                    tbiCiudad.IsEnabled = False
                Case 2
                    tbiPais.IsEnabled = False
                    tbiDpto.IsEnabled = False
                    tbiCiudad.IsEnabled = True
            End Select
        End If

    End Sub
    Private Sub SHabiliteMenuesUbicacion()
        Dim lstrRecurso As String
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            lstrRecurso = "RecMnuItemPriHab"
            Select Case tbcUbicacion.SelectedIndex
                Case 0
                    If MnuAccionesPais.Visibility = Visibility.Visible Then
                        MnuAccionesPais.Style = FindResource(lstrRecurso)
                    End If
                Case 1
                    If MnuAccionesDpto.Visibility = Visibility.Visible Then
                        If Not IsNothing(MobjPais) AndAlso
                            Not String.IsNullOrEmpty(MobjPais.ObjIdPaisStr.ToString) Then
                            MnuAccionesDpto.Style = FindResource(lstrRecurso)
                        Else
                            MnuAccionesDpto.Style = FindResource("RecMnuItemPriDes")
                        End If
                    End If
                Case 2
                    If MnuAccionesCiudad.Visibility = Visibility.Visible Then
                        If Not IsNothing(MobjDpto) AndAlso MobjDpto.ObjIdDptoByt.ObjValorPro > 0 Then
                            MnuAccionesCiudad.Style = FindResource(lstrRecurso)
                        Else
                            MnuAccionesCiudad.Style = FindResource("RecMnuItemPriDes")
                        End If
                    End If
            End Select
        Else
            lstrRecurso = "RecMnuItemPriDes"
            MnuAccionesPais.Style = TryFindResource(lstrRecurso)
            MnuAccionesDpto.Style = TryFindResource(lstrRecurso)
            MnuAccionesCiudad.Style = TryFindResource(lstrRecurso)
        End If
    End Sub
    Private Sub SDeshabiliteControlesActuales()
        txtIdPais.Style = FindResource("RecCtlNoHabilitado")
        txtIdDpto.Style = FindResource("RecCtlNoHabilitado")
        txtIdCiudad.Style = FindResource("RecCtlNoHabilitado")
        txtNombrePais.Style = FindResource("RecCtlNoHabilitado")
        txtNombreDpto.Style = FindResource("RecCtlNoHabilitado")
        txtNombreCiudad.Style = FindResource("RecCtlNoHabilitado")
    End Sub
    Private Sub SHabiliteControlesNuevos(ablnHabilite As Boolean)
        SVisibiliceControlesNuevos(ablnHabilite)
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuModificando Then
            Select Case MobjObjetoWin.EnuTipoAccion
                Case EnuTipoAccionDef.enuNuevoPais
                    txtIdPaisNuevo.Style = FindResource("RecCtlHabilitado")
                    txtNombrePaisNuevo.Style = FindResource("RecCtlHabilitado")
                Case EnuTipoAccionDef.enuNuevoDpto
                    txtIdDptoNuevo.Style = FindResource("RecCtlHabilitado")
                    txtNombreDptoNuevo.Style = FindResource("RecCtlHabilitado")
                Case EnuTipoAccionDef.enuNuevoCiud
                    txtIdCiudadNueva.Style = FindResource("RecCtlHabilitado")
                    txtNombreCiudadNueva.Style = FindResource("RecCtlHabilitado")
                Case EnuTipoAccionDef.enuModifPais
                    txtIdPaisNuevo.Style = FindResource("RecCtlNoHabilitado")
                    txtNombrePaisNuevo.Style = FindResource("RecCtlHabilitado")
                Case EnuTipoAccionDef.enuModifDpto
                    txtIdDptoNuevo.Style = FindResource("RecCtlNoHabilitado")
                    txtNombreDptoNuevo.Style = FindResource("RecCtlHabilitado")
                Case EnuTipoAccionDef.enuModifCiud
                    txtIdCiudadNueva.Style = FindResource("RecCtlNoHabilitado")
                    txtNombreCiudadNueva.Style = FindResource("RecCtlHabilitado")
            End Select
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
        Select Case tbcUbicacion.SelectedIndex
            Case 0
                txtIdPais.Visibility = lvisVisibilidadActual
                txtNombrePais.Visibility = lvisVisibilidadActual
                txtIdPaisNuevo.Visibility = lvisVisibilidadNuevos
                txtNombrePaisNuevo.Visibility = lvisVisibilidadNuevos
            Case 1
                txtIdDpto.Visibility = lvisVisibilidadActual
                txtNombreDpto.Visibility = lvisVisibilidadActual
                txtIdDptoNuevo.Visibility = lvisVisibilidadNuevos
                txtNombreDptoNuevo.Visibility = lvisVisibilidadNuevos
            Case 2
                txtIdCiudad.Visibility = lvisVisibilidadActual
                txtNombreCiudad.Visibility = lvisVisibilidadActual
                txtIdCiudadNueva.Visibility = lvisVisibilidadNuevos
                txtNombreCiudadNueva.Visibility = lvisVisibilidadNuevos
        End Select
    End Sub
    Private Sub SOrdeneGrid()
        Select Case tbcUbicacion.SelectedIndex
            Case 0
                SOrdeneDataGrid(dgrPaises, dgrPaises.Columns(1), "Nombre", ListSortDirection.Ascending)
            Case 1
                SOrdeneDataGrid(dgrDptos, dgrDptos.Columns(2), "Nombre", ListSortDirection.Ascending)
            Case 2
                SOrdeneDataGrid(dgrCiudades, dgrCiudades.Columns(3), "Nombre", ListSortDirection.Ascending)
        End Select
    End Sub
    Private Sub SDetermineObjPais()
        If MobjObjetoWin.EnuTipoAccion <> EnuTipoAccionDef.enuNuevoPais Then
            If Not String.IsNullOrEmpty(txtIdPais.Text) Then
                MobjPais = MobjObjetoWin.ObjPais(txtIdPais.Text)
            Else
                MobjPais = Nothing
            End If
        End If
    End Sub
    Private Sub SDetermineObjDpto()
        If MobjObjetoWin.EnuTipoAccion <> EnuTipoAccionDef.enuNuevoDpto Then
            MobjDpto = Nothing
            If Not String.IsNullOrEmpty(txtIdDpto.Text) AndAlso txtIdDpto.Text <> "0" Then
                SDetermineObjPais()
                If Not IsNothing(MobjPais) Then
                    MobjDpto = MobjPais.ObjDpto(CType(txtIdDpto.Text, Byte))
                End If
            End If
        End If
    End Sub
    Private Sub SDetermineObjCiudad()
        If MobjObjetoWin.EnuTipoAccion <> EnuTipoAccionDef.enuNuevoCiud Then
            MobjCiudad = Nothing
            If Not String.IsNullOrEmpty(txtIdCiudad.Text) AndAlso txtIdCiudad.Text <> "0" Then
                SDetermineObjDpto()
                If Not IsNothing(MobjDpto) Then
                    MobjCiudad = MobjDpto.ObjCiudad(CType(txtIdCiudad.Text, Short))
                End If
            End If
        End If
    End Sub
    Private Sub SAsigneVariablesObjetos()
        mshrIdBarrio = 0
        mbytIdDpto = 0
        mshrIdCiudad = 0
        mshrIdBarrio = 0
        If Not String.IsNullOrEmpty(txtIdPais.Text) Then
            MstrIdPais = txtIdPais.Text
        End If
        If Not String.IsNullOrEmpty(txtIdDpto.Text) Then
            mbytIdDpto = CType(txtIdDpto.Text, Byte)
        End If
        If Not String.IsNullOrEmpty(txtIdCiudad.Text) Then
            mshrIdCiudad = CType(txtIdCiudad.Text, Short)
        End If
    End Sub
#End Region
#Region "Acciones de Modificación"
    Private Sub SCreePais()
        MobjPais = MobjObjetoWin.FobjNuevoPais
        MobjObjetoWin.EnuTipoAccion = EnuTipoAccionDef.enuNuevoPais
        SHabiliteControlesNuevos(True)
        SHabiliteTabItems(False)
        SHabiliteDataGrids(False)
        txtEstadoWin.Content = "Creando un nuevo País"
        SMuestreDatos()
        txtIdPaisNuevo.Focus()
    End Sub
    Private Sub SCreeDpto()
        SDetermineObjPais()
        If Not IsNothing(MobjPais) Then
            MobjDpto = MobjPais.FobjNuevoDpto
            MobjObjetoWin.EnuTipoAccion = EnuTipoAccionDef.enuNuevoDpto
            SHabiliteControlesNuevos(True)
            SHabiliteTabItems(False)
            SHabiliteDataGrids(False)
            txtEstadoWin.Content = "Creando un nuevo Departamento"
            SMuestreDatos()
            txtIdDptoNuevo.Focus()
        Else
            SLevanteEveNoti("No hay un País valido seleccionado!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SCreeCiudad()
        SDetermineObjDpto()
        If Not IsNothing(MobjDpto) Then
            MobjCiudad = MobjDpto.FobjNuevoCiudad
            MobjObjetoWin.EnuTipoAccion = EnuTipoAccionDef.enuNuevoCiud
            SHabiliteControlesNuevos(True)
            SHabiliteTabItems(False)
            SHabiliteDataGrids(False)
            txtEstadoWin.Content = "Creando una nueva Ciudad"
            SMuestreDatos()
            txtIdCiudadNueva.Focus()
        Else
            SLevanteEveNoti("No hay un Departamento valido seleccionado!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SModifiquePais()
        SDetermineObjPais()
        If Not IsNothing(MobjPais) Then
            MobjObjetoWin.EnuTipoAccion = EnuTipoAccionDef.enuModifPais
            SHabiliteControlesNuevos(True)
            SHabiliteTabItems(False)
            SHabiliteDataGrids(False)
            txtEstadoWin.Content = "Modificando País"
            txtNombrePaisNuevo.Focus()
        Else
            SLevanteEveNoti("No hay seleccionado un Pais para ser modificado!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SModifiqueDpto()
        SDetermineObjDpto()
        If Not IsNothing(MobjDpto) Then
            MobjObjetoWin.EnuTipoAccion = EnuTipoAccionDef.enuModifDpto
            SHabiliteControlesNuevos(True)
            SHabiliteTabItems(False)
            SHabiliteDataGrids(False)
            txtEstadoWin.Content = "Modificando Departamento"
            txtNombreDptoNuevo.Focus()
        Else
            SLevanteEveNoti("No hay seleccionado un Departamento para ser modificado!",
                    "", 0, EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SModifiqueCiudad()
        SDetermineObjCiudad()
        If Not IsNothing(MobjCiudad) Then
            MobjObjetoWin.EnuTipoAccion = EnuTipoAccionDef.enuModifCiud
            SHabiliteControlesNuevos(True)
            SHabiliteTabItems(False)
            SHabiliteDataGrids(False)
            txtEstadoWin.Content = "Modificando Ciudad"
            txtNombreCiudadNueva.Focus()
        Else
            SLevanteEveNoti("No hay seleccionado una Ciudad para ser modificada!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
        End If
    End Sub
    Private Sub SEliminePais()
        SLevanteEveNoti("Esta operación solo puede ser llevada a cabo desde el Administrador!",
                "", 0, EnuSeveridadNot.EnuInformacion)
    End Sub
    Private Sub SElimineDpto()
        SLevanteEveNoti("Esta operación solo puede ser llevada a cabo desde el Administrador!",
                "", 0, EnuSeveridadNot.EnuInformacion)
    End Sub
    Private Sub SElimineCiudad()
        SLevanteEveNoti("Esta operación solo puede ser llevada a cabo desde el Administrador!",
                "", 0, EnuSeveridadNot.EnuInformacion)
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        Dim lstrNombreAccion As String = lelmElemento.Name.Substring(3)
        If Not FblnEjecutoOpcion(lstrNombreAccion) Then
            Select Case lstrNombreAccion
                Case "EliminarDpto"
                    SElimineDpto()
                Case "CrearCiudad"
                    SCreeCiudad()
                Case "ModificarCiudad"
                    SModifiqueCiudad()
                Case "EliminarCiudad"
                    SElimineCiudad()
            End Select
        End If
    End Sub
    Private Function FblnEjecutoOpcion(astrNombreAccion As String) As Boolean
        Dim lblnEje = True
        Select Case astrNombreAccion
            Case "CrearPais"
                SCreePais()
            Case "ModificarPais"
                SModifiquePais()
            Case "EliminarPais"
                SEliminePais()
            Case "CrearDpto"
                SCreeDpto()
            Case "ModificarDpto"
                SModifiqueDpto()
            Case Else
                lblnEje = False
        End Select
        Return lblnEje
    End Function
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        End If
    End Sub
    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
            If TypeOf lelmElemento Is TextBox Then
                Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                Try
                    Select Case True
                        Case lelmElemento.Name = "txtIdPaisNuevo"
                            MobjPais.ObjIdPaisStr.ObjValorPro = txtIdPaisNuevo.Text
                        Case lelmElemento.Name = "txtNombrePaisNuevo"
                            MobjPais.ObjNombrePaisStr.ObjValorPro = txtNombrePaisNuevo.Text
                        Case lelmElemento.Name = "txtIdDptoNuevo"
                            MobjDpto.ObjIdDptoByt.ObjValorPro = txtIdDptoNuevo.Text
                        Case lelmElemento.Name = "txtNombreDptoNuevo"
                            MobjDpto.ObjNombreDptoStr.ObjValorPro = txtNombreDptoNuevo.Text
                        Case lelmElemento.Name = "txtIdCiudadNueva"
                            MobjCiudad.ObjIdCiudadShr.ObjValorPro = txtIdCiudadNueva.Text
                        Case lelmElemento.Name = "txtNombreCiudadNueva"
                            MobjCiudad.ObjNombreCiudadStr.ObjValorPro = txtNombreCiudadNueva.Text
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
    End Sub
    Private Sub TxtNombre_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtNombrePais.TextChanged,
            txtNombreDpto.TextChanged, txtNombreCiudad.TextChanged
        SAsigneVariablesObjetos()
        Select Case True
            Case sender.Equals(txtNombrePais)
                txtPais.Content = txtIdPais.Text & " - " & txtNombrePais.Text
                txtIdPaisNuevo.Text = txtIdPais.Text
                txtNombrePaisNuevo.Text = txtNombrePais.Text
                If Not String.IsNullOrEmpty(MstrIdPais) Then
                    MobjPais = MobjObjetoWin.ObjPais(MstrIdPais)
                Else
                    MobjPais = Nothing
                End If
                tbiDpto.DataContext = MobjObjetoWin.DtbDptos(MstrIdPais)
            Case sender.Equals(txtNombreDpto)
                txtDpto.Content = txtIdDpto.Text & " - " & txtNombreDpto.Text
                txtIdDptoNuevo.Text = txtIdDpto.Text
                txtNombreDptoNuevo.Text = txtNombreDpto.Text
                If mbytIdDpto > 0 Then
                    MobjDpto = MobjPais.ObjDpto(mbytIdDpto)
                Else
                    MobjDpto = Nothing
                End If
                tbiCiudad.DataContext = MobjObjetoWin.DtbCiudades(MstrIdPais, mbytIdDpto)
            Case sender.Equals(txtNombreCiudad)
                txtCiudad.Content = txtIdCiudad.Text & " - " & txtNombreCiudad.Text
                txtIdCiudadNueva.Text = txtIdCiudad.Text
                txtNombreCiudadNueva.Text = txtNombreCiudad.Text
                If mshrIdCiudad > 0 Then
                    MobjCiudad = MobjDpto.ObjCiudad(mshrIdCiudad)
                Else
                    MobjCiudad = Nothing
                End If
        End Select
    End Sub
    Private Sub TbcUbicacion_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles tbcUbicacion.SelectionChanged
        If Not IsNothing(MnuAccionesPais) Then
            MnuAccionesPais.Visibility = Visibility.Collapsed
            MnuAccionesDpto.Visibility = Visibility.Collapsed
            MnuAccionesCiudad.Visibility = Visibility.Collapsed
            Select Case tbcUbicacion.SelectedIndex
                Case 0
                    MnuAccionesPais.Visibility = Visibility.Visible
                Case 1
                    MnuAccionesDpto.Visibility = Visibility.Visible
                Case 2
                    MnuAccionesCiudad.Visibility = Visibility.Visible
            End Select
            SHabiliteMenuesUbicacion()
        End If
    End Sub
#Region "Eventros Datagrid"
    Private Sub DgrPaises_AutoGeneratedColumns(sender As Object, e As EventArgs) Handles dgrPaises.AutoGeneratedColumns
        dgrPaises.Columns(0).Width = 65
        dgrPaises.Columns(1).Width = 265
        dgrPaises.Columns(0).Header = "Id. Pais"
        dgrPaises.Columns(1).Header = "Nombre Pais"
    End Sub
    Private Sub DgrDptos_AutoGeneratedColumns(sender As Object, e As EventArgs) Handles dgrDptos.AutoGeneratedColumns
        dgrDptos.Columns(0).Width = 65
        dgrDptos.Columns(1).Visibility = Visibility.Collapsed
        dgrDptos.Columns(2).Width = 265
        dgrDptos.Columns(0).Header = "Id. Dpto."
        dgrDptos.Columns(2).Header = "Nombre Departamento"
    End Sub
    Private Sub DgrCiudades_AutoGeneratedColumns(sender As Object, e As EventArgs) Handles dgrCiudades.AutoGeneratedColumns
        dgrCiudades.Columns(0).Width = 65
        dgrCiudades.Columns(1).Visibility = Visibility.Collapsed
        dgrCiudades.Columns(2).Visibility = Visibility.Collapsed
        dgrCiudades.Columns(3).Width = 265
        dgrCiudades.Columns(0).Header = "Id. Ciudad"
        dgrCiudades.Columns(3).Header = "Nombre Ciudad"
    End Sub
#End Region
#End Region
End Class