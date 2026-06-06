Imports System.Windows.Controls
Imports OPT.OrionP.GesDat.ActualizaBd
Public Class WinBaseDatos
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
    'Variables
    Private MobjBaseDatos As ClsBaseDatos = Nothing
    Private MobjBaseDatosPan As ClsBaseDatos = Nothing
    Private MobjTabla As ClsTabla = Nothing
    Private MobjColumna As ClsColumna = Nothing
    Private MobjIndice As ClsIndice = Nothing
    Private MtviBaseDatos As TreeViewItem = Nothing
    Private MtviBaseDatosPan As TreeViewItem = Nothing
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomBasDat
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuBaseDatos
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SCargueForma(EnuElementosAdicionalesDef.None, 0, Nothing, Nothing, True)
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
        ObjObjetoWin = GobjPanorama.ObjCarpetaActual
        EnuTipoPermisoObjWin = GobjPanorama.ObjCarpetaActual.EnuPermisosObj
    End Sub

    Protected Overrides Sub SInicialiceControles()
        Dim lshrIdAdmin = EnuListaAplicaciones.EnuAdministrador
        Dim lshrIdOrion = EnuListaAplicaciones.EnuOrionCop
        MobjBaseDatosPan = ClsPanoramaDat.FobjEstructuraBD(lshrIdAdmin)
        MobjBaseDatos = ClsPanoramaDat.FobjEstructuraBD(lshrIdOrion)
        SPuebleArbol()
    End Sub

    Protected Overrides Sub SMuestreDatos()
        trvOrionPlus.Focus()
    End Sub
    Protected Overrides Sub SValide()
        '
    End Sub
    Protected Overrides Sub SRegistre()
        '
    End Sub
    ''' <summary>
    ''' Adiciona al menu de la ventana (hmnuMiMenu) los items de acuerdo al tipo de ventana y al objeto de la
    ''' ventana "objObjetoWin". 
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overrides Sub SConfigureMenuesPropios()
        Dim lmnuItem As MenuItem
        Dim lsepSrparador As New Separator
        HmnuAcciones.Items.Insert(5, lsepSrparador)
        lmnuItem = FmnuiMenuItem("MnuContraerRama", "Contraer _Rama", "RecMnuItemSec")
        HmnuAcciones.Items.Insert(5, lmnuItem)
        lmnuItem = FmnuiMenuItem("MnuContraerTodo", "_Contraer Todo", "RecMnuItemSec", "Ctrl+C")
        HmnuAcciones.Items.Insert(5, lmnuItem)
        lmnuItem = FmnuiMenuItem("MnuExpandirRama", "Expandir _Rama", "RecMnuItemSec")
        HmnuAcciones.Items.Insert(5, lmnuItem)
        lmnuItem = FmnuiMenuItem("MnuExpandirTodo", "_Expandir Todo", "RecMnuItemSec", "Ctrl+E")
        HmnuAcciones.Items.Insert(5, lmnuItem)
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    '
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnMenuClic(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is MenuItem Then
            Select Case lelmElemento.Name
                Case "MnuExpandirTodo"
                    SExpandaTodo(MtviBaseDatos)
                    SExpandaTodo(MtviBaseDatosPan)
                Case "MnuExpandirRama"
                    SExpandaRama(trvOrionPlus.SelectedItem)
                Case "MnuContraerTodo"
                    SContraigaTodo(MtviBaseDatos)
                    SContraigaTodo(MtviBaseDatosPan)
                Case "MnuContraerRama"
                    SContraigaRama(trvOrionPlus.SelectedItem)
            End Select
        End If
    End Sub

    Private Sub GridSplitter_MouseEnter(sender As Object, e As MouseEventArgs)
        Mouse.OverrideCursor = Cursors.ScrollWE
    End Sub

    Private Sub GridSplitter_MouseLeave(sender As Object, e As MouseEventArgs)
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub

    Private Sub TrvOrionPlus_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of System.Object)) Handles trvOrionPlus.SelectedItemChanged
        Dim ltviOri As TreeViewItem = CType(e.NewValue, TreeViewItem)
        Dim lstrNombreTabla As String
        SExpandaRama(ltviOri)
        lvOrionCop.Items.Clear()
        Select Case ltviOri.Tag
            Case "BD"
                SMuestreBD()
            Case "TBL"
                lstrNombreTabla = ltviOri.Name
                If lstrNombreTabla.ToUpper.StartsWith("PAN") Then
                    MobjTabla = MobjBaseDatosPan.ColTablas(lstrNombreTabla)
                Else
                    MobjTabla = MobjBaseDatos.ColTablas(lstrNombreTabla)
                End If
                SMuestreTabla()
            Case "COLS"
                Dim ltviTabla As TreeViewItem = ltviOri.Parent
                If ltviTabla.Name.ToUpper.StartsWith("PAN") Then
                    MobjTabla = MobjBaseDatosPan.ColTablas(ltviTabla.Name)
                Else
                    lstrNombreTabla = ltviTabla.Name
                    MobjTabla = MobjBaseDatos.ColTablas(lstrNombreTabla)
                End If
                SMuestreCols()
            Case "COL"
                Dim ltviCols As TreeViewItem = ltviOri.Parent
                Dim ltviTabla As TreeViewItem = ltviCols.Parent
                If ltviTabla.Name.ToUpper.StartsWith("PAN") Then
                    MobjTabla = MobjBaseDatosPan.ColTablas(ltviTabla.Name)
                Else
                    lstrNombreTabla = ltviTabla.Name
                    MobjTabla = MobjBaseDatos.ColTablas(lstrNombreTabla)
                End If
                MobjColumna = MobjTabla.ColColumnas(ltviOri.Name)
                SMuestreColumna()
            Case "INDS"
                Dim ltviTabla As TreeViewItem = ltviOri.Parent
                If ltviTabla.Name.ToUpper.StartsWith("PAN") Then
                    MobjTabla = MobjBaseDatosPan.ColTablas(ltviTabla.Name)
                Else
                    lstrNombreTabla = ltviTabla.Name
                    MobjTabla = MobjBaseDatos.ColTablas(lstrNombreTabla)
                End If
                SMuestreInds()
            Case "IND"
                Dim ltviInds As TreeViewItem = ltviOri.Parent
                Dim ltviTabla As TreeViewItem = ltviInds.Parent
                lstrNombreTabla = ltviTabla.Name
                If lstrNombreTabla.ToUpper.StartsWith("PAN") Then
                    MobjTabla = MobjBaseDatosPan.ColTablas(ltviTabla.Name)
                Else
                    MobjTabla = MobjBaseDatos.ColTablas(lstrNombreTabla)
                End If
                MobjIndice = MobjTabla.ColIndices(ltviOri.Name)
                SMuestreIndice()
        End Select
    End Sub

    Private Sub ClsFormInterface_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyboardDevice.Modifiers = ModifierKeys.Control Then
            If e.Key = Key.E Then
                SExpandaTodo(MtviBaseDatos)
                SExpandaTodo(MtviBaseDatosPan)
            End If
            If e.Key = Key.C Then
                SContraigaTodo(MtviBaseDatos)
                SContraigaTodo(MtviBaseDatosPan)
            End If
        End If
    End Sub

#End Region
#Region "Procedimientos Pueblan Arbol"
    Private Sub SPuebleArbol()
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
#Region "Manejo Arbol"
    Private Sub SExpandaTodo(atviNodo As TreeViewItem)
        atviNodo.IsExpanded = True
        If atviNodo.Items.Count > 0 Then
            For Each ltviNodo As TreeViewItem In atviNodo.Items
                SExpandaTodo(ltviNodo)
            Next
        End If
        MtviBaseDatos.IsSelected = True
    End Sub

    Private Shared Sub SExpandaRama(atviNodo As TreeViewItem)
        atviNodo.IsExpanded = True
    End Sub

    Private Shared Sub SContraigaRama(atviNodo As TreeViewItem)
        atviNodo.IsExpanded = False
    End Sub

    Private Sub SContraigaTodo(atviNodo As TreeViewItem)
        If atviNodo.Items.Count > 0 Then
            For Each ltviNodo As TreeViewItem In atviNodo.Items
                SContraigaTodo(ltviNodo)
            Next
        End If
        atviNodo.IsExpanded = False
    End Sub
#End Region
#Region "Mostrar Objetos"
    Private Sub SMuestreBD()
        With MobjBaseDatos
            lvOrionCop.Items.Add(New ClsViewItemPar("Nombre Base de Datos", .StrNombreBD))
            lvOrionCop.Items.Add(New ClsViewItemPar("Versión", .EntVersion))
            lvOrionCop.Items.Add(New ClsViewItemPar("Conjunto de Caracteres", .StrCharacterSet))
            lvOrionCop.Items.Add(New ClsViewItemPar("Tipo comparación", .StrCollationName))
            lvOrionCop.Items.Add(New ClsViewItemPar("Cantidad de Tablas", .ColTablas.Count))
            lvOrionCop.Items.Add(New ClsViewItemPar("Cantidad de Comandos", .ColComandos.Count))
        End With
    End Sub

    Private Sub SMuestreTabla()
        With MobjTabla
            lvOrionCop.Items.Add(New ClsViewItemPar("Nombre de la Tabla", .StrNombre))
            lvOrionCop.Items.Add(New ClsViewItemPar("Cantidad de Columnas", .ColColumnas.Count))
            lvOrionCop.Items.Add(New ClsViewItemPar("Cantidad de Indices", .ColIndices.Count))
            If .ColComandos.Count > 0 Then
                lvOrionCop.Items.Add(New ClsViewItemPar("Cantidad de Comandos", .ColComandos.Count))
            End If
        End With
    End Sub

    Private Sub SMuestreCols()
        With MobjTabla
            lvOrionCop.Items.Add(New ClsViewItemPar("Cantidad de Columnas", .ColColumnas.Count))
        End With
    End Sub

    Private Sub SMuestreColumna()
        With MobjColumna
            lvOrionCop.Items.Add(New ClsViewItemPar("Nombre de la Columna", .StrNombre))
            lvOrionCop.Items.Add(New ClsViewItemPar("Tipo de Datos", .StrTipoDatos))
            lvOrionCop.Items.Add(New ClsViewItemPar("Adminte Null", IIf(.BlnRequerido, "No", "Si")))
            lvOrionCop.Items.Add(New ClsViewItemPar("Es Autonumérico", IIf(.BlnAutoNumerico, "Si", "No")))
            If Not String.IsNullOrEmpty(.StrLongitud) Then
                lvOrionCop.Items.Add(New ClsViewItemPar("Longitud del Campo", .StrLongitud))
            End If
            If Not String.IsNullOrEmpty(.StrComentario) Then
                lvOrionCop.Items.Add(New ClsViewItemPar("Comentario", .StrComentario))
            End If
        End With
    End Sub

    Private Sub SMuestreInds()
        With MobjTabla
            lvOrionCop.Items.Add(New ClsViewItemPar("Cantidad de Indices", .ColIndices.Count))
        End With
    End Sub

    Private Sub SMuestreIndice()
        With MobjIndice
            lvOrionCop.Items.Add(New ClsViewItemPar("Nombre del Indice", .StrNombre))
            lvOrionCop.Items.Add(New ClsViewItemPar("Es Principal", IIf(.BlnPrincipal, "Si", "No")))
            lvOrionCop.Items.Add(New ClsViewItemPar("Es Unico", IIf(.BlnUnico, "Si", "No")))
            lvOrionCop.Items.Add(New ClsViewItemPar("", ""))
            lvOrionCop.Items.Add(New ClsViewItemPar("Columnas que conforman el Indice", ""))
            For Each lobjColInd As ClsColumnaIndice In .ColColumnasIndice
                Dim lstrOrden As String = IIf(lobjColInd.BlnAscendente, "Ascendente", "Descendente")
                Dim lstrCol = lobjColInd.StrNombre & (" orden ") & lstrOrden
                lvOrionCop.Items.Add(New ClsViewItemPar("Columna", lstrCol))
            Next
        End With
    End Sub
#End Region
End Class
