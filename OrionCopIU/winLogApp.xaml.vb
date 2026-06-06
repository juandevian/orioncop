Imports System.Windows.Controls
Imports System.Data
Public Class WinLogApp
#Region "Definiciones"
    ' Herencia e Interfaz
    Inherits ClsFormInterface
    Implements IDisposable
    ' Variables
    Private MdtbLogApp As DataTable = Nothing
    Private MdvwLogApp As DataView = Nothing
    Private MblnPoblando As Boolean = False
    Private ReadOnly MstrNombreVentana As String = My.Resources.NomLogApp
#End Region
#Region "Constructor"
    Public Sub New()
        InitializeComponent()
        HenuIdVentana = EnuIdVentanaDef.enuLogApp
    End Sub
#End Region
#Region "Invalida metodos en la clase base que implementan la Interfaz"
    Protected Overrides Sub SLoad()
        SAdicioneCtlsRestringidos()
        SCargueForma(EnuElementosAdicionalesDef.None, 0,
                Nothing, Nothing, True)
        SEstablezcaFiltro()
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
        ObjObjetoWin = GobjParametros
        '
        MdtbLogApp = ClsPanorama.FdtbLogApp
        MdvwLogApp = New DataView(MdtbLogApp) With {
            .Sort = "IdLogApp ASC"
        }
        EnuTipoPermisoObjWin = GobjParametros.EnuPermisosObj
    End Sub

    Protected Overrides Sub SInicialiceControles()
        MblnPoblando = True
        SPuebleComboBoxes()
        chkFechas.IsChecked = True
        dtpDesde.SelectedDate = Date.Today.AddMonths(-1)
        dtpHasta.SelectedDate = Date.Today
        MblnPoblando = False
    End Sub

    Protected Overrides Sub SMuestreDatos()
        '
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
        '
    End Sub
#End Region
#Region "Procedimientos invalidantes"
    Protected Overrides Sub SRefresqueWin()
        SPuebleComboBoxes()
        MdtbLogApp = ClsPanorama.FdtbLogApp
        MdvwLogApp = New DataView(MdtbLogApp)
        SEstablezcaFiltro()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
#End Region
#Region "Busqueda"
    ' 
#End Region
#Region "Eventos en la Ventana"
    Private Sub OnCogerFoco(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is TextBox Then
            Dim ltxtTextBox As TextBox = lelmElemento
            ltxtTextBox.SelectAll()
        End If
    End Sub

    Private Sub TxtIdObjeto_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtIdObjeto.TextChanged
        SEstablezcaFiltro()
    End Sub

    Private Sub Echk_Click(sender As Object, e As RoutedEventArgs) Handles chkTipo.Click, chkClase.Click,
            chkDato.Click, chkFechas.Click, chkIdObjeto.Click, chkUsuario.Click
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is CheckBox Then
            Select Case lelmElemento.Name
                Case "chkFechas"
                    MblnPoblando = True
                    If chkFechas.IsChecked Then
                        dtpDesde.IsEnabled = True
                        dtpHasta.IsEnabled = True
                    Else
                        dtpDesde.IsEnabled = False
                        dtpHasta.IsEnabled = False
                    End If
                    MblnPoblando = False
            End Select
        End If
        SEstablezcaFiltro()
    End Sub

    Private Sub Edtp_SelectedDateChanged(sender As Object, e As SelectionChangedEventArgs) Handles dtpDesde.SelectedDateChanged,
            dtpHasta.SelectedDateChanged
        If Not MblnPoblando Then
            SEstablezcaFiltro()
        End If
    End Sub

    Private Sub Ecbo_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cboTipoLog.SelectionChanged,
            cboNombreClase.SelectionChanged, cboNombreDato.SelectionChanged, cboUsuario.SelectionChanged
        If Not MblnPoblando Then
            SEstablezcaFiltro()
        End If
    End Sub
#End Region
#Region "Procedimientos Propios"
    Private Sub SAdicioneCtlsRestringidos()
        SAdicioneControlRestringido(chkClase)
        SAdicioneControlRestringido(chkDato)
        SAdicioneControlRestringido(chkFechas)
        SAdicioneControlRestringido(chkIdObjeto)
        SAdicioneControlRestringido(chkTipo)
        SAdicioneControlRestringido(chkUsuario)
        SAdicioneControlRestringido(dtpDesde)
        SAdicioneControlRestringido(dtpHasta)
        SAdicioneControlRestringido(cboNombreClase)
        SAdicioneControlRestringido(cboNombreDato)
        SAdicioneControlRestringido(cboTipoLog)
        SAdicioneControlRestringido(cboUsuario)
        SAdicioneControlRestringido(txtIdObjeto)
    End Sub

    Private Sub SPuebleComboBoxes()
        Dim ldrwTiposLog = ClsPanorama.FdrwTiposLog
        Dim ldrwNombresClases As DataRow() = ClsPanorama.FdrwNombresClases
        Dim ldrwNombresDatos As DataRow() = ClsPanorama.FdrwNombresDatos
        Dim ldrwUsuarios As DataRow() = ClsPanorama.FdrwUsuarios
        SPuebleComboBox(ldrwTiposLog, cboTipoLog)
        SPuebleCbo(ldrwNombresClases, cboNombreClase)
        SPuebleCbo(ldrwNombresDatos, cboNombreDato)
        SPuebleCbo(ldrwUsuarios, cboUsuario)
    End Sub

    Private Shared Sub SPuebleCbo(adrwDatos As DataRow(), acboCombo As ComboBox)
        acboCombo.Items.Clear()
        If adrwDatos.Length > 0 Then
            For Each ldrwDat As DataRow In adrwDatos
                acboCombo.Items.Add(ldrwDat(0))
            Next
        End If
    End Sub

    Private Sub SEstablezcaFiltro()
        If Not MblnPoblando Then
            Dim lstrFiltro As String = String.Empty
            If chkFechas.IsChecked Then
                Dim lstrFechaIni As String = dtpDesde.SelectedDate.ToString
                Dim lstrFechaFin As String = dtpHasta.SelectedDate.ToString
                lstrFechaIni = ClsPanoramaDat.FstrFechaHoraInicioDia(lstrFechaIni)
                lstrFechaFin = ClsPanoramaDat.FstrFechaHoraFinDia(lstrFechaFin)
                lstrFiltro = "FechaCreacion >= '" & lstrFechaIni & "' AND FechaCreacion <= '" & lstrFechaFin & "' AND "
            End If
            If chkTipo.IsChecked Then
                If Not IsNothing(cboTipoLog.SelectedItem) Then
                    lstrFiltro &= "Dato = '" & cboTipoLog.SelectedItem & "' AND "
                End If
            End If
            If chkClase.IsChecked Then
                If Not IsNothing(cboNombreClase.SelectedItem) Then
                    lstrFiltro &= "NombreClase = '" & cboNombreClase.SelectedItem & "' AND "
                End If
            End If
            If chkDato.IsChecked Then
                If Not IsNothing(cboNombreDato.SelectedItem) Then
                    lstrFiltro &= "Nombre = '" & cboNombreDato.SelectedItem & "' AND "
                End If
            End If
            If chkIdObjeto.IsChecked Then
                lstrFiltro &= "IdObjeto = '" & txtIdObjeto.Text & "' AND "
            End If
            If chkUsuario.IsChecked Then
                If Not IsNothing(cboUsuario.SelectedItem) Then
                    lstrFiltro &= "IdUsuario = '" & cboUsuario.SelectedItem & "' AND "
                End If
            End If
            If Not String.IsNullOrEmpty(lstrFiltro) Then
                lstrFiltro = lstrFiltro.Substring(0, lstrFiltro.Length - 5)
                MdvwLogApp.RowFilter = lstrFiltro
            Else
                MdvwLogApp.RowFilter = Nothing
            End If
        End If
        dgrLogApp.DataContext = MdvwLogApp
    End Sub
#End Region
#Region "Implementa IDisposable"
    Protected Overridable Overloads Sub Dispose(disposing As Boolean)
        If disposing Then
            MdvwLogApp.Dispose()
            MdtbLogApp.Dispose()
        End If
    End Sub

    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
