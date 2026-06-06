Imports System.Windows
Imports System.Data
Imports System.Windows.Controls
Imports System.ComponentModel
Public Class WinCuentasContabilidad
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuIdCuenta = 0
        enuNombre
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsCuentaContabilidad = Nothing
    Private MobjMiCarpeta As ClsCarpeta = Nothing
    Private MblnDejoUltimoControl As Boolean = False
    Private MblnCreo As Boolean = False
    Dim MnuImportar As MenuItemPan = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomCtaCon
#End Region

#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuCuentasContabilidad
    End Sub
#End Region

#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        Dim lcolCamposLlave As New Collection From {
            txtIdCuentaNueva
        }
        SAdicioneControlRestringido(txtIdCuenta)
        SAdicioneControlRestringido(txtNombre)
        SAdicioneControlRestringido(dgrCuentasCont)
        SAdicioneControlRestringido(txtBuscar)
        SAdicioneControlRestringido(bttIr)
        SCargueForma(EnuElementosAdicionalesDef.None, 2,
                lcolCamposLlave, txtNombreNuevo, False)
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
        MobjMiCarpeta = GobjPanorama.ObjCarpetaActual
        If IsNothing(ObjObjetoWin) Then
            ObjObjetoWin = New ClsCuentaContabilidad(MobjMiCarpeta,
                    EnuModoInstanciaObjDef.EnuNavegable)
        End If
        MobjObjetoWin = ObjObjetoWin
        If Not MobjObjetoWin.FblnEstaVacioOrigenDatos Then
            MobjObjetoWin.SVayaAlPrimero()
        End If
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub

    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuIdCuenta) = lblIdCuenta
        StcValidaControl(EnuValidEntrada.enuNombre) = lblNombreCuenta
        '
        SDeshabiliteControlesActuales()
        SModifiqueBarraHerramientas()
        SMuestreUbicacion()
        Dim ldtbCtasCont = MobjMiCarpeta.FdtbCuentasCont()
        dgrCuentasCont.DataContext = ldtbCtasCont
        grdCuentasContabilidad.DataContext = ldtbCtasCont
        HbttAceptar.TabIndex = 26
        HbttCancelar.TabIndex = 27
    End Sub

    Protected Overrides Sub SMuestreDatos()
        If Not IsNothing(MobjObjetoWin) Then
            With MobjObjetoWin
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    txtIdCuentaNueva.Text = .ObjIdCuentaContStr.ObjValorPro
                    txtNombreNuevo.Text = .ObjNombreCuentaStr.ObjValorPro
                Else
                    If MobjMiCarpeta.FdtbCuentasCont().Rows.Count = 0 Then
                        SLevanteEveNoti("No hay Cuentas de Contabilidad para ser mostradas!", "", 0,
                                EnuSeveridadNot.EnuInformacion)
                    End If
                End If
            End With
        End If
        SValide()
    End Sub

    Protected Overrides Sub SValide()
        With MobjObjetoWin
            If MobjMiCarpeta.FdtbCuentasCont().Rows.Count = 0 AndAlso EnuOperacionEnWin =
                    EnuOperacionEnVentana.cenuConsultando Then
                SInicialiceValido()
            Else
                StcValidValido(EnuValidEntrada.enuIdCuenta) = .ObjIdCuentaContStr.BlnEsValido
                StcValidValido(EnuValidEntrada.enuNombre) = .ObjNombreCuentaStr.BlnEsValido
            End If
        End With
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub

    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjIdCarpetaCuentaShr.ObjValorPro = GshrIdCarpeta
            .ObjIdCuentaContStr.ObjValorPro = txtIdCuentaNueva.Text
            .ObjNombreCuentaStr.ObjValorPro = txtNombreNuevo.Text
        End With
    End Sub

    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        HbttCrear.Visibility = Visibility.Visible
        HbttSuprimir.Visibility = Visibility.Visible
        Dim lmnuCrear As MenuItem = Nothing
        Dim lmnuSuprimir As MenuItem = Nothing
        For Each lobjMenu As Object In HmnuAcciones.Items
            If TypeOf lobjMenu Is MenuItem Then
                If lobjMenu.Name = "MnuCrear" Then
                    lmnuCrear = lobjMenu
                ElseIf lobjMenu.Name = "MnuSuprimir" Then
                    lmnuSuprimir = lobjMenu
                End If
            End If
        Next
        lmnuCrear.Visibility = Visibility.Visible
        lmnuSuprimir.Visibility = Visibility.Visible
        ' Adicionar menú Importar
        MnuImportar = FmnuiMenuItemPan("MnuImportar", "_Importar Cuentas Contables", 1, "")
        Dim lsepSeparador As New Separator
        HmnuAcciones.Items.Insert(7, MnuImportar)
        HmnuAcciones.Items.Insert(8, lsepSeparador)
    End Sub
#End Region

#Region "Procedimientos invalidantes"
    Protected Overrides Sub SHabiliteMenues()
        MyBase.SHabiliteMenues()
    End Sub

    Protected Overrides Sub SCree()
        txtBuscar.Text = ""
        ObjObjetoWin.SCreeObj(ObjValorLlave)
        EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando
        SFrmAdicione()
        SCreeCuentaCont()
    End Sub

    Protected Overrides Sub SModifique()
        txtIdCuentaNueva.Text = txtIdCuenta.Text
        txtNombreNuevo.Text = txtNombre.Text
        MyBase.SModifique()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            SModifiqueCuentaCont()
        End If
    End Sub

    Protected Overrides Sub SGuarde()
        MyBase.SGuarde()
        SRefresqueWin()
    End Sub

    Protected Overrides Sub SRefresqueWin()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            GobjPanDat.SControleProcesoObj(True)
            MyBase.SRefresqueWin()
            Dim ldtbCtasCont = MobjMiCarpeta.FdtbCuentasCont()
            dgrCuentasCont.DataContext = ldtbCtasCont
            grdCuentasContabilidad.DataContext = ldtbCtasCont
            SOrdeneDataGrid(dgrCuentasCont, dgrCuentasCont.Columns(0), "IdCuentaCont", ListSortDirection.Ascending)
            SValide()
            GobjPanDat.SControleProcesoObj(False)
            Mouse.OverrideCursor = Cursors.Arrow
        End If
    End Sub

    Protected Overrides Sub SSuprima()
        Dim lstrMens = String.Empty, lblnNoHayError As Boolean
        Dim lstrMensEx = String.Empty, lblnSuprimio = False
        Try
            GobjPanDat.SControleProcesoObj(True)
            Dim lblnEsSuprimible = MobjObjetoWin.FblnEsSuprimible()
            If lblnEsSuprimible Then
                Dim lstrIdCtaCon = MobjObjetoWin.ObjIdCuentaContStr.ToString
                lblnEsSuprimible = ClsOrionCop.FblnCtaConEsEliminables(lstrIdCtaCon)
                If lblnEsSuprimible Then
                    If MsgBox("Esta seguro de suprimir la presente " & ObjObjetoWin.StrNombreClase & "?",
                        MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Supresión") = MsgBoxResult.Yes Then
                        lblnSuprimio = MobjObjetoWin.FblnSuprimio()
                    End If
                    SFinaliceOperacion()
                End If
            End If
            If Not lblnEsSuprimible Then
                lstrMens = "La Cuenta no puede ser suprimida porque ya esta en uso!"
            End If
            If lblnSuprimio Then
                lstrMens = "La Cuenta fue suprimida exitosamente!"
            End If
            lblnNoHayError = True
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As ArgumentNullException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If lblnNoHayError Then
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SRefresqueWin()
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                    GobjPanDat.SControleProcesoObj(False)
                End If
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                GobjPanDat.SControleProcesoObj(False, True)
            End If
        End Try
    End Sub

    Protected Overrides Sub SEstablezcaWinConsultando()
        MyBase.SEstablezcaWinConsultando()
        SVisibiliceControlesNuevos(False)
        Dim ldtbCtasCont As DataTable = MobjMiCarpeta.FdtbCuentasCont()
        dgrCuentasCont.DataContext = ldtbCtasCont
        grdCuentasContabilidad.DataContext = ldtbCtasCont
        dgrCuentasCont.IsEnabled = True
        txtBuscar.Style = FindResource("RecCtlHabilitado")
    End Sub
#End Region

#Region "Procedimientos Propios"
    Private Sub SMuestreUbicacion()
        txtNombreCarpeta.Content = MobjMiCarpeta.ObjIdCarpetaShr.ToString & " - " &
                MobjMiCarpeta.ObjNombreStr.ToString
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
        If Not IsNothing(HmnuNavegar) Then
            HmnuNavegar.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub SDeshabiliteControlesActuales()
        Dim lstyEstiloNoHabilitado As Style = FindResource("RecCtlNoHabilitado")
        txtIdCuenta.Style = lstyEstiloNoHabilitado
        txtNombre.Style = lstyEstiloNoHabilitado
        bttIr.IsEnabled = True
    End Sub

    Private Sub SHabiliteControlesNuevos(ablnHabilite As Boolean)
        SVisibiliceControlesNuevos(ablnHabilite)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            txtIdCuentaNueva.Style = FindResource("RecCtlHabilitado")
            txtNombreNuevo.Style = FindResource("RecCtlHabilitado")
            txtBuscar.Style = FindResource("RecCtlNoHabilitado")
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
        txtIdCuenta.Visibility = lvisVisibilidadActual
        txtIdCuentaNueva.Visibility = lvisVisibilidadNuevos
        txtNombre.Visibility = lvisVisibilidadActual
        txtNombreNuevo.Visibility = lvisVisibilidadNuevos
    End Sub

    Private Sub SCreeCuentaCont()
        SHabiliteControlesNuevos(True)
        dgrCuentasCont.IsEnabled = False
        SMuestreDatos()
        txtIdCuentaNueva.Focus()
    End Sub

    Private Sub SModifiqueCuentaCont()
        If Not IsNothing(MobjObjetoWin) AndAlso MobjObjetoWin.BlnExiste Then
            SHabiliteControlesNuevos(True)
            dgrCuentasCont.IsEnabled = False
            txtIdCuentaNueva.Style = FindResource("RecCtlNoHabilitado")
            txtNombreNuevo.Focus()
            SMuestreDatos()
        End If
    End Sub

    Private Sub SUbiqueCuenta()
        If Not String.IsNullOrEmpty(txtBuscar.Text) Then
            Dim ldtbCtasCont As DataTable = MobjMiCarpeta.FdtbCuentasCont()
            If MblnCreo Then
                dgrCuentasCont.DataContext = ldtbCtasCont
                grdCuentasContabilidad.DataContext = ldtbCtasCont
                MblnCreo = False
            End If
            Static lstrNomCta As String
            Static lentIdBus As Integer
            If lstrNomCta <> txtBuscar.Text Then
                lstrNomCta = txtBuscar.Text
                lentIdBus = 1
            Else
                lentIdBus += 1
            End If
            Dim lentIndice = MobjObjetoWin.FentIndiceCta(ldtbCtasCont, txtBuscar.Text, lentIdBus)
            If lentIndice > -1 Then
                dgrCuentasCont.UpdateLayout()
                dgrCuentasCont.SelectedIndex = lentIndice
                dgrCuentasCont.ScrollIntoView(dgrCuentasCont.SelectedItem)
            Else
                lentIdBus = 0
            End If
        End If
    End Sub

    Private Sub SAbraImportar()
        Dim lwinImportar As New WinImportar(MobjObjetoWin) With {
            .WinPadre = Me
        }
        lwinImportar.ShowDialog()
        SRefrescarClic()
    End Sub
#End Region

#Region "Eventos en la Ventana"
    Private Sub BttIr_Click(sender As Object, e As RoutedEventArgs) Handles bttIr.Click
        SUbiqueCuenta()
    End Sub

    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            If lelmElemento.Name = "MnuImportar" Then
                SAbraImportar()
            End If
        End If
    End Sub

    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        ElseIf lelmElemento.Equals(HbttCancelar) Then
            If MblnDejoUltimoControl AndAlso FblnEstanTodosBien() Then
                HbttAceptar.Focus()
            End If
        End If
    End Sub

    Private Sub OnPierdeFoco(sender As Object, e As RoutedEventArgs)
        If Not HblnSeEstaCerrando Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    Dim lblnNoHayError = False, lstrMens = String.Empty, lstrMensEx = String.Empty
                    MblnDejoUltimoControl = False
                    Try
                        Select Case True
                            Case lelmElemento.Name = "txtIdCuentaNueva"
                                MobjObjetoWin.ObjIdCuentaContStr.ObjValorPro = txtIdCuentaNueva.Text
                            Case lelmElemento.Name = "txtNombreNuevo"
                                MobjObjetoWin.ObjNombreCuentaStr.ObjValorPro = txtNombreNuevo.Text
                                MblnDejoUltimoControl = True
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

    Private Sub DgrCuentasCont_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles dgrCuentasCont.SelectionChanged
        Dim ldrvFilaActual As DataRowView
        If TypeOf sender Is DataGrid Then
            Dim ldgrActual As DataGrid = sender
            ldrvFilaActual = ldgrActual.SelectedItem
            If Not IsNothing(ldrvFilaActual) Then
                Dim lstrIdCta = ldrvFilaActual("IdCuentaCont")
                Dim lobjvalorllave() As Object = {GshrIdCarpeta, lstrIdCta}
                MobjObjetoWin.SAbra(lobjvalorllave)
                ObjObjetoWin = MobjObjetoWin
            End If
        End If
    End Sub

    Private Sub TxtBuscar_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtBuscar.TextChanged
        SLevanteEveNoti("", "", 0, EnuSeveridadNot.EnuOk)
    End Sub
#End Region
End Class