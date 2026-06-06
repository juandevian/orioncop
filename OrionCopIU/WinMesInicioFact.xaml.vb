Public Class WinMesInicioFact
    Inherits Window
    Friend BlnInicioFacMesActual As Boolean = False
    Public Sub New()
        InitializeComponent()
        GblnOK = False
    End Sub
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        rdbMesAct.IsChecked = True
        Dim NoUsado = rdbMesAct.Focus()
    End Sub
    Private Sub BttCancelar_Click(sender As Object, e As RoutedEventArgs) Handles bttCancelar.Click
        GblnOK = False
        Close()
    End Sub
    Private Sub BttAceptar_Click(sender As Object, e As RoutedEventArgs) Handles bttAceptar.Click
        GblnOK = True
        BlnInicioFacMesActual = rdbMesAct.IsChecked
        Close()
    End Sub
End Class
