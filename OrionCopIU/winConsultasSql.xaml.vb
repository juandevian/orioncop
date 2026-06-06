Imports System.Windows.Controls
Imports System.Data
Public Class WinConsultasSql
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
#Region "Enumeradores"
    'El siguiente enumerador se utiliza para el mecanismo de validacion. Debe haber un elemento por campo a validar.
    Private Enum EnuValidEntradaDef As Integer
        enuNombreConsulta
        enuConsulta
    End Enum
#End Region
    ' Variables
    Private MobjObjetoWin As ClsConsultaSql = Nothing
    Private MdtbResultado As DataTable = Nothing
    'Variables Base de datos
    Private MobjBaseDatos As ClsBaseDatos = Nothing
    Private MobjBaseDatosPan As ClsBaseDatos = Nothing
    Private MobjTabla As ClsTabla = Nothing
    Private MtviBaseDatos As TreeViewItem = Nothing
    Private MtviBaseDatosPan As TreeViewItem = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomConSql
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuConsultasSql
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneControlRestringido(cboNombreConsulta)
        SCargueForma(EnuElementosAdicionalesDef.None, 2,
                Nothing, Nothing, False)
        cboNombreConsulta.Focus()
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
        ObjObjetoWin = New ClsConsultaSql(GstrIdUsuario = GCSTRUSUARIOU)
        MobjObjetoWin = ObjObjetoWin
        EnuTipoPermisoObjWin = MobjObjetoWin.EnuPermisosObj
    End Sub
    Protected Overrides Sub SInicialiceControles()
        StcValidaControl(EnuValidEntradaDef.enuNombreConsulta) = lblNombreConsulta
        StcValidaControl(EnuValidEntradaDef.enuConsulta) = lblConsulta
        SInvisibiliceCtls()
        SPuebleComboConsultas()
        SPuebleArbol()
        HbttAceptar.TabIndex = 10
        HbttCancelar.TabIndex = 11
    End Sub
    Protected Overrides Sub SMuestreDatos()
        If Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos Then
            SLevanteEveNoti("No hay Consultas para ser mostradas!", "", 0,
                    EnuSeveridadNot.EnuInformacion)
            cboNombreConsulta.IsEnabled = False
        End If
        With MobjObjetoWin
            txtIdConsulta.Content = .ObjIdConsultaShr.ToString
            txtConsulta.Text = .ObjExpresionSqlStr.ObjValorPro
            If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuCreando Then
                txtNombreConsulta.Text = .ObjNombreConsultaStr.ObjValorPro
            End If
        End With
        SValide()
    End Sub
    Protected Overrides Sub SValide()
        Dim lblnNoHayDatos = Not BlnVentanaAux AndAlso ObjObjetoWin.FblnEstaVacioOrigenDatos
        If lblnNoHayDatos AndAlso EnuOperacionEnWin =
                EnuOperacionEnVentanaDef.cenuConsultando Then
            SInicialiceValido()
        Else
            With MobjObjetoWin
                StcValidValido(EnuValidEntradaDef.enuNombreConsulta) = .ObjNombreConsultaStr.BlnEsValido
                StcValidValido(EnuValidEntradaDef.enuConsulta) = .ObjExpresionSqlStr.BlnEsValido
            End With
        End If
        '
        SHabiliteBotonesTlb()
        bttEjecutar.IsEnabled = FblnEstanTodosBien()
    End Sub
    Protected Overrides Sub SRegistre()
        With MobjObjetoWin
            .ObjCamposAgrupamientoStr.ObjValorPro = Nothing
            .ObjCamposIndiceStr.ObjValorPro = Nothing
            .ObjCamposInsertarStr.ObjValorPro = Nothing
            .ObjCamposPrimariosStr.ObjValorPro = Nothing
            .ObjCamposRefActualizacionStr.ObjValorPro = Nothing
            .ObjCamposRefEliminacionStr.ObjValorPro = Nothing
            .ObjCamposRelPrimariosStr.ObjValorPro = Nothing
            .ObjCamposRelSecundariosStr.ObjValorPro = Nothing
            .ObjCamposSecundariosStr.ObjValorPro = Nothing
            .ObjFiltroStr.ObjValorPro = String.Empty
            .ObjEsIndiceUnicoBln.ObjValorPro = True
            .ObjTablaSecundariaStr.ObjValorPro = Nothing
            .ObjTablaPrimariaStr.ObjValorPro = Nothing
            .ObjIdAppConsultaShr.ObjValorPro = GenuIdAplicacion
            .ObjOrigenInstanciaStr.ObjValorPro = GstrOrigenActual
            If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuCreando Then
                .ObjNombreConsultaStr.ObjValorPro = txtNombreConsulta.Text
            Else
                If Not IsNothing(cboNombreConsulta.SelectedItem) Then
                    .ObjNombreConsultaStr.ObjValorPro = cboNombreConsulta.SelectedItem
                Else
                    .ObjNombreConsultaStr.ObjValorPro = String.Empty
                End If
            End If
            .ObjExpresionSqlStr.ObjValorPro = txtConsulta.Text
        End With
        SValide()
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
        MyBase.SCree()
        cboNombreConsulta.Visibility = Visibility.Collapsed
        txtNombreConsulta.Visibility = Visibility.Visible
        txtNombreConsulta.IsEnabled = True
        MdtbResultado = Nothing
        dgrResultado.DataContext = MdtbResultado
        bttExportar.IsEnabled = False
        MobjObjetoWin.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuExpresionSql
        MobjObjetoWin.ObjNombreConsultaStr.ObjValorPro = txtNombreConsulta.Text
        MobjObjetoWin.ObjExpresionSqlStr.ObjValorPro = txtConsulta.Text
        SValide()
    End Sub
    ''' <summary>
    ''' Prepara la ventana y su objeto para modificar el objeto. Invalida la función "SModifique"
    ''' de la clase base.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SModifique()
        MyBase.SModifique()
        cboNombreConsulta.IsEnabled = False
        txtNombreConsulta.IsEnabled = False
    End Sub
    Protected Overrides Sub SFinaliceOperacion()
        MyBase.SFinaliceOperacion()
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            txtNombreConsulta.Visibility = Visibility.Collapsed
            cboNombreConsulta.Visibility = Visibility.Visible
            MdtbResultado = Nothing
            dgrResultado.DataContext = MdtbResultado
            bttExportar.IsEnabled = False
            SPuebleComboConsultas()
            cboNombreConsulta.IsEnabled = True
        End If
    End Sub
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnBotonClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Dim lbttBoton As Button = lelmElemento
            Select Case lbttBoton.Name
                Case "bttEjecutar"
                    SEjecuteConsulta()
                Case "bttExportar"
                    SExporteAExcel()
            End Select
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
        If Not HblnSeEstaCerrando Then
            Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
            If TypeOf lelmElemento Is TextBox Then
                If EnuOperacionEnWin <> EnuOperacionEnVentanaDef.cenuConsultando Then
                    With MobjObjetoWin
                        Select Case lelmElemento.Name
                            Case "txtNombreConsulta"
                                .ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuExpresionSql
                                .ObjNombreConsultaStr.ObjValorPro = txtNombreConsulta.Text
                            Case "txtConsulta"
                                .ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuExpresionSql
                                .ObjExpresionSqlStr.ObjValorPro = txtConsulta.Text.Trim
                        End Select
                    End With
                    SMuestreDatos()
                End If
            End If
        End If
    End Sub
    Private Sub Ecbo_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cboNombreConsulta.SelectionChanged
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            MobjObjetoWin.ObjTipoConsultaByt.ObjValorPro = EnuTipoConsultaDef.enuExpresionSql
            MobjObjetoWin.ObjNombreConsultaStr.ObjValorPro = cboNombreConsulta.SelectedItem
            MdtbResultado = Nothing
            dgrResultado.DataContext = MdtbResultado
            bttExportar.IsEnabled = False
            SMuestreDatos()
        End If
    End Sub
    Private Sub GridSplitter_MouseEnter(sender As Object, e As MouseEventArgs)
        Mouse.OverrideCursor = Cursors.ScrollWE
    End Sub
    Private Sub GridSplitter_MouseLeave(sender As Object, e As MouseEventArgs)
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SPuebleComboConsultas()
        cboNombreConsulta.Items.Clear()
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            Dim lalsNomCol As ArrayList = MobjObjetoWin.FalsNombresConsultas
            For Each lstrNombre As String In lalsNomCol
                cboNombreConsulta.Items.Add(lstrNombre)
            Next
        End If
        If EnuOperacionEnWin = EnuOperacionEnVentanaDef.cenuConsultando Then
            cboNombreConsulta.SelectedIndex = 0
        End If
    End Sub
    Private Sub SInvisibiliceCtls()
        HmnuNavegar.Visibility = Visibility.Collapsed
        HbttAlAnterior.Visibility = Visibility.Collapsed
        HbttAlPrimero.Visibility = Visibility.Collapsed
        HbttAlSiguiente.Visibility = Visibility.Collapsed
        HbttAlUltimo.Visibility = Visibility.Collapsed
        HbttBuscar.Visibility = Visibility.Collapsed
    End Sub
    Private Sub SEjecuteConsulta()
        Dim lstrConsulta = txtConsulta.Text.Trim
        Dim lentRegistrosAfectados As Integer
        Dim lstrTipo = lstrConsulta.Substring(0, lstrConsulta.IndexOf(" "))
        Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
        Try
            MdtbResultado = Nothing
            dgrResultado.DataContext = MdtbResultado
            Select Case lstrTipo.ToUpper
                Case "SELECT"
                    MdtbResultado = ClsPanorama.FdtbDataTable(lstrConsulta)
                    dgrResultado.DataContext = MdtbResultado
                Case "UPDATE", "INSERT", "DELETE"
                    lentRegistrosAfectados = GobjPanDat.SEjecuteSentenciaSql(lstrConsulta)
                    lstrMens = "Se afectaron " & lentRegistrosAfectados.ToString &
                            " registros!"
                Case Else
                    lstrMens = "No se puede ejecutar porque no es una expresión válida!"
            End Select
            lblnNoHayError = True
        Catch ex As ProveedorBdPanException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanDatException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As PanLException
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Catch ex As Exception
            lstrMens = ex.Message
            lstrMensEx = ex.ToString
        Finally
            If Not IsNothing(MdtbResultado) Then
                bttExportar.IsEnabled = (MdtbResultado.Rows.Count > 0)
            Else
                bttExportar.IsEnabled = False
            End If
            If lblnNoHayError Then
                If Not String.IsNullOrEmpty(lstrMens) Then
                    SLevanteEveNoti(lstrMens, String.Empty, 0, EnuSeveridadNot.EnuInformacion)
                End If
            Else
                SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
            End If
        End Try
    End Sub
    Private Sub SExporteAExcel()
        If FblnEstanTodosBien() Then
            Mouse.OverrideCursor = Cursors.Wait
            Dim lstrMens = String.Empty, lstrMensEx = String.Empty, lblnNoHayError = False
            Try
                Dim lobjExp As New ClsExportToExcell
                Dim lstrFecha = Today.Year.ToString & Format(Today.Month, "00") &
                        Format(Today.Day, "00")
                Dim lstrNomArchivo = lstrFecha & "_" & MobjObjetoWin.ObjNombreConsultaStr.ObjValorPro
                If lstrNomArchivo.Length > 31 Then
                    lstrNomArchivo = lstrNomArchivo.Substring(0, 31)
                End If
                lstrNomArchivo = GstrTrayReportes & "\" & lstrNomArchivo
                lobjExp.SExporteToExcel(lstrNomArchivo, MdtbResultado)
                If MsgBox("Desea abrir el Archivo de Excel?", MsgBoxStyle.Question + vbYesNo,
                        "Abrir Archivo") = MsgBoxResult.Yes Then
                    lobjExp.SAbraExcel()
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
                If Not lblnNoHayError Then
                    SLevanteEveNoti(lstrMens, lstrMensEx, 0, EnuSeveridadNot.EnuExcep)
                End If
                Mouse.OverrideCursor = Cursors.Arrow
            End Try
        End If
    End Sub
#End Region
#Region "Procedimientos Pueblan Arbol"
    Private Sub SPuebleArbol()
        Dim lshrIdAdmin = EnuListaAplicaciones.EnuAdministrador
        Dim lshrIdOrion = EnuListaAplicaciones.EnuOrionCop
        MobjBaseDatosPan = ClsPanoramaDat.FobjEstructuraBD(lshrIdAdmin)
        MobjBaseDatos = ClsPanoramaDat.FobjEstructuraBD(lshrIdOrion)
        ' Nodo Raiz (Base de Datos)
        With MobjBaseDatos
            MtviBaseDatos = FtviTviPan(.StrNombreBD, "RecImagenes/database.png")
            MtviBaseDatos.Name = .StrNombreBD
            MtviBaseDatos.Tag = "BD"
            trvOrionPlus.Items.Add(MtviBaseDatos)
            For Each lobjTabla As ClsTabla In .ColTablas
                MobjTabla = lobjTabla
                SAdicioneTabla(lobjTabla, False)
            Next
        End With
        With MobjBaseDatosPan
            MtviBaseDatosPan = FtviTviPan(.StrNombreBD, "RecImagenes/database.png")
            MtviBaseDatosPan.Name = .StrNombreBD
            MtviBaseDatosPan.Tag = "BD"
            trvOrionPlus.Items.Add(MtviBaseDatosPan)
            For Each lobjTabla As ClsTabla In .ColTablas
                MobjTabla = lobjTabla
                SAdicioneTabla(lobjTabla, True)
            Next
        End With
    End Sub
    Private Sub SAdicioneTabla(aobjTabla As ClsTabla, ablnPanorama As Boolean)
        Dim ltviNodoTabla = FtviTviPan(aobjTabla.StrNombre, "RecImagenes/DataTable.png")
        ltviNodoTabla.Tag = "TBL"
        Dim ltviNodoColumnas = FtviTviPan("Columnas", "RecImagenes/DataColumns.png")
        ltviNodoColumnas.Tag = "COLS"
        Dim ltviNodoIndices = FtviTviPan("Indices", "RecImagenes/DataIndex.png")
        ltviNodoIndices.Tag = "INDS"
        ltviNodoTabla.Name = aobjTabla.StrNombre
        ltviNodoTabla.Items.Add(ltviNodoColumnas)
        ltviNodoTabla.Items.Add(ltviNodoIndices)
        ltviNodoColumnas.Name = "Columnas"
        ltviNodoIndices.Name = "Indices"
        For Each lobjColumna As ClsColumna In aobjTabla.ColColumnas
            SAdicioneColumna(ltviNodoColumnas, lobjColumna)
        Next
        For Each lobjIndice As ClsIndice In aobjTabla.ColIndices
            SAdicioneIndice(ltviNodoIndices, lobjIndice)
        Next
        If ablnPanorama Then
            MtviBaseDatosPan.Items.Add(ltviNodoTabla)
        Else
            MtviBaseDatos.Items.Add(ltviNodoTabla)
        End If
    End Sub
    Private Sub SAdicioneColumna(atviNodoPadre As TreeViewItem, aobjColumna As ClsColumna)
        Dim ltviNodoColumna As TreeViewItem
        If FblnColumnaEsDelIndice(aobjColumna.StrNombre) Then
            ltviNodoColumna = FtviTviPan(aobjColumna.StrNombre, "RecImagenes/DataIndex.png", 18)
        Else
            ltviNodoColumna = FtviTviPan(aobjColumna.StrNombre, "RecImagenes/DataColumns.png", 18)
        End If
        ltviNodoColumna.Name = aobjColumna.StrNombre
        ltviNodoColumna.Tag = "COL"
        atviNodoPadre.Items.Add(ltviNodoColumna)
    End Sub
    Private Shared Sub SAdicioneIndice(atviNodoPadre As TreeViewItem, aobjIndice As ClsIndice)
        Dim ltviNodoIndice = FtviTviPan(aobjIndice.StrNombre, "RecImagenes/DataIndex.png", 18)
        ltviNodoIndice.Name = aobjIndice.StrNombre
        ltviNodoIndice.Tag = "IND"
        atviNodoPadre.Items.Add(ltviNodoIndice)
    End Sub
    Private Function FblnColumnaEsDelIndice(astrNombreCol As String) As Boolean
        Dim lblnEsDelIndice = False
        For Each lobjIndice As ClsIndice In MobjTabla.ColIndices
            If lobjIndice.BlnPrincipal Then
                Dim lobjColumIndi As ClsColumnaIndice = Nothing
                For i = 1 To lobjIndice.ColColumnasIndice.Count
                    lobjColumIndi = lobjIndice.ColColumnasIndice(i)
                    If lobjColumIndi.StrNombre = astrNombreCol Then
                        lblnEsDelIndice = True
                        Exit For
                    End If
                Next
                If lblnEsDelIndice Then Exit For
            End If
        Next
        Return lblnEsDelIndice
    End Function
#End Region
End Class
