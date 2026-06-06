Class Application
#Region "Eventos de la Aplicación"
    Private Sub Application_Startup(sender As Object, e As StartupEventArgs)
        GenuTamanoIcono = EnuTamanoIconos.enuGrande
        Mouse.OverrideCursor = Cursors.AppStarting
        SInicialiceApp()
        Mouse.OverrideCursor = Cursors.Arrow
    End Sub
#End Region
End Class
