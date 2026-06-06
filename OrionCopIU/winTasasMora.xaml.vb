Imports System.ComponentModel
Public Class WinTasasMora
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    Private Enum EnuValidEntrada As Integer
        enuFechaDesde = 0
        enuTasaMora
    End Enum
#End Region
    ' Variables
    Private MobjTasaMora As ClsTasaMora = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomTasMor
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuTasasMora
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneControlRestringido(dgrTasasMora)
        SAdicioneControlRestringido(txtFechaHasta)
        SCargueForma(EnuElementosAdicionalesDef.None, 2,
                Nothing, dtpFechaDesde, False)
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
            ObjObjetoWin = New ClsTasaMora(EnuModoInstanciaObjDef.enuNavegable)
        End If
        MobjTasaMora = ObjObjetoWin
        MobjTasaMora.SVayaAlPrimero()
        EnuTipoPermisoObjWin = MobjTasaMora.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntrada.enuFechaDesde) = lblFechaDesde
        StcValidaControl(EnuValidEntrada.enuTasaMora) = lblTasaMora
        '
        grdTasasMora.DataContext = GobjParametros.FdtbTasasMora()
        '
        txtTasaMoraNueva.ToolTip = "Tasa Mensual"
        SModifiqueBarraHerramientas()
        HbttAceptar.TabIndex = 19
        HbttCancelar.TabIndex = 20
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando AndAlso
                MobjTasaMora IsNot Nothing Then
            dtpFechaDesdeNueva.Text = MobjTasaMora.ObjFechaDesdeTasaMoraDtm.ObjValorPro.ToString
            txtTasaMoraNueva.Text = Format(MobjTasaMora.ObjTasaMoraDbl.DblTasaMensual, "#0.000%")
            txtTasaEA.Content = Format(MobjTasaMora.ObjTasaMoraDbl.ObjValorPro, "#0.000%")
        ElseIf Not IsNothing(MobjTasaMora) Then
            If GobjParametros.FdtbTasasMora().Rows.Count = 0 Then
                SLevanteEveNoti("No hay Tasas de Mora para ser mostradas!", "", 0,
                        EnuSeveridadNot.EnuInformacion)
            Else
                txtTasaEA.Content = Format(MobjTasaMora.ObjTasaMoraDbl.ObjValorPro, "#0.000%")
                txtTasaMora.Text = Format(MobjTasaMora.ObjTasaMoraDbl.DblTasaMensual, "#0.000%")
            End If
        End If
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        With MobjTasaMora
            If GobjParametros.FdtbTasasMora.Rows.Count = 0 AndAlso EnuOperacionEnWin =
                    EnuOperacionEnVentana.cenuConsultando Then
                SInicialiceValido()
            Else
                StcValidValido(EnuValidEntrada.enuFechaDesde) = .ObjFechaDesdeTasaMoraDtm.BlnEsValido
                StcValidValido(EnuValidEntrada.enuTasaMora) = .ObjTasaMoraDbl.BlnEsValido
            End If
        End With
        SHabiliteBotonesTlb()
        FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjTasaMora
            .ObjFechaDesdeTasaMoraDtm.ObjValorPro = dtpFechaDesdeNueva.SelectedDate
            .ObjFechaHastaTasaMoraDtm.ObjValorPro = Date.Today
            .ObjTasaMoraDbl.ObjValorPro = FdblTasaMora()
        End With
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        '
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    ''' <summary>
    ''' Sub que prepara a la ventana y a su objeto para crear un nuevo objeto. Invalida el Sub
    ''' "SCree" de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SCree()
        Dim ldtmFechaIni As Date = GCDTMFECHANULA
        If Not MobjTasaMora.BlnExiste Then
            ldtmFechaIni = GobjParametros.ObjAnoActual.ObjPeriodoActual.DtmFechaInicioPeriodo
        End If
        MyBase.SCree()
        MobjTasaMora.ObjFechaDesdeTasaMoraDtm.ObjValorPro = ldtmFechaIni
        SDeshabiliteControlesActuales()
        SCreeTasaMora()
    End Sub
    ''' <summary>
    ''' Prepara la ventana y su objeto para modificar el objeto. Invalida la función "SModifique"
    ''' de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SModifique()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            If dgrTasasMora.Items.Count > 0 Then
                dgrTasasMora.SelectedIndex = dgrTasasMora.Items.Count - 1
            End If
        End If
        MyBase.SModifique()
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            SDeshabiliteControlesActuales()
            SModifiqueTasaMora()
        End If
    End Sub
    Protected Overrides Sub SRefresqueWin()
        If EnuOperacionEnWin = EnuOperacionEnVentana.cenuConsultando Then
            Mouse.OverrideCursor = Cursors.Wait
            MobjTasaMora.SRefresqueObj()
            grdTasasMora.DataContext = GobjParametros.FdtbTasasMora()
            SOrdeneDataGrid(dgrTasasMora, dgrTasasMora.Columns(0), ClsOrdinalTasaMoraEnt.SstrNombreCampoBd,
                            ListSortDirection.Ascending)
            SMuestreDatos()
            Mouse.OverrideCursor = Cursors.Arrow
        End If
    End Sub
    Protected Overrides Sub SSuprima()
        MyBase.SSuprima()
        SRefresqueWin()
    End Sub
    Protected Overrides Sub SFinaliceOperacion()
        Dim lstrMens As String = "Los datos del objeto " & ObjObjetoWin.StrNombreClase &
                " han cambiado!" & vbCrLf & "Desea guardar los cambios?"
        Dim lblnCreanado = (EnuOperacionEnWin = EnuOperacionEnVentana.cenuCreando)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            With MobjTasaMora
                If .EnuEstadoActualizacion <> EnuEstadoObjetoDef.enuConsultando Then
                    If FblnEstanTodosBien() AndAlso .BlnTengoCambios AndAlso
                            .ObjTasaMoraDbl.ObjValorPro > 0 Then
                        If MsgBox(lstrMens, vbYesNo, "Aceptar Cambios") = vbYes Then
                            .SActualice(True)
                            If lblnCreanado Then
                                .SNormaliceEstado(True)
                                SEstablezcaWinConsultando()
                            End If
                        Else
                            .SNormaliceEstado(True)
                            SEstablezcaWinConsultando()
                            If BlnCanceleCierra Then
                                SCerrarClic()
                            End If
                        End If
                    Else
                        .SNormaliceEstado(True)
                        SEstablezcaWinConsultando()
                        If BlnCanceleCierra Then
                            SCerrarClic()
                        End If
                    End If
                Else
                    .SNormaliceEstado(True)
                    SEstablezcaWinConsultando()
                End If
            End With
        End If
        SRefresqueWin()
        If String.IsNullOrEmpty(dtpFechaDesde.Text) AndAlso dgrTasasMora.Items.Count > 0 Then
            dgrTasasMora.SelectedIndex = 0
        End If
        SVisibiliceControlesNuevos(False)
        dgrTasasMora.IsEnabled = True
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
        txtOrdinal.Style = FindResource("RecCtlNoHabilitado")
        dtpFechaDesde.Style = FindResource("RecCtlNoHabilitado")
        txtTasaMora.Style = FindResource("RecCtlNoHabilitado")
    End Sub
    Private Sub SHabiliteControlesNuevos(ablnHabilite As Boolean)
        SVisibiliceControlesNuevos(ablnHabilite)
        If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
            dtpFechaDesdeNueva.Style = FindResource("RecCtlHabilitado")
            txtTasaMoraNueva.Style = FindResource("RecCtlHabilitado")
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
        dtpFechaDesde.Visibility = lvisVisibilidadActual
        txtTasaMora.Visibility = lvisVisibilidadActual
        dtpFechaDesdeNueva.Visibility = lvisVisibilidadNuevos
        txtTasaMoraNueva.Visibility = lvisVisibilidadNuevos
    End Sub
    Private Sub SCreeTasaMora()
        SHabiliteControlesNuevos(True)
        dgrTasasMora.IsEnabled = False
        SMuestreDatos()
        dtpFechaDesdeNueva.Focus()
    End Sub
    Private Sub SModifiqueTasaMora()
        If Not IsNothing(MobjTasaMora) AndAlso MobjTasaMora.BlnExiste Then
            SHabiliteControlesNuevos(True)
            dgrTasasMora.IsEnabled = False
            txtTasaMora.Focus()
        End If
    End Sub
    Private Function FdblTasaMora() As Double
        Dim lstrTasaMora = txtTasaMoraNueva.Text, ldblTasa As Double
        If String.IsNullOrEmpty(txtTasaMoraNueva.Text) Then
            lstrTasaMora = "0"
        End If
        If lstrTasaMora.EndsWith("%") Then
            ldblTasa = lstrTasaMora.Substring(0, lstrTasaMora.Length - 1).Trim / 100
        Else
            ldblTasa = lstrTasaMora.Trim / 100
        End If
        Return Math.Round(ldblTasa, 6)
    End Function
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
        If Not HblnSeEstaCerrando Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox OrElse TypeOf lelmElemento Is DatePicker Then
                If EnuOperacionEnWin <> EnuOperacionEnVentana.cenuConsultando Then
                    Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
                    Try
                        Select Case lelmElemento.Name
                            Case "dtpFechaDesdeNueva"
                                MobjTasaMora.ObjFechaDesdeTasaMoraDtm.ObjValorPro = dtpFechaDesdeNueva.SelectedDate
                            Case "txtTasaMoraNueva"
                                MobjTasaMora.ObjTasaMoraDbl.ObjValorPro = FdblTasaMora()
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
    Private Sub TxtOrdinal_TextChanged(sender As Object, e As TextChangedEventArgs) _
            Handles txtOrdinal.TextChanged
        If Not String.IsNullOrEmpty(txtOrdinal.Text) Then
            Dim lshrIdCarpeta As Short = GobjPanorama.ObjCarpetaActual.ObjIdCarpetaShr.ObjValorPro
            Dim lshrIdCenUtil As Short = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.ObjIdCentroUtilShr.ObjValorPro
            Dim lobjValorLlave() As Object = {lshrIdCarpeta, lshrIdCenUtil, CType(txtOrdinal.Text, Integer)}
            MobjTasaMora.SAbra(lobjValorLlave)
        End If
        SMuestreDatos()
    End Sub
#End Region
End Class