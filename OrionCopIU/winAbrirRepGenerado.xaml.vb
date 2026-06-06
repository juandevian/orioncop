Imports Microsoft.Win32
Public Class WinAbrirRepGenerado
#Region "Definiciones"
    'Variables
    Private MstrMens = String.Empty, MstrMensEx = String.Empty
    Private MblnEsValido As Boolean = False
    Private Sub SValide()
        MblnEsValido = My.Computer.FileSystem.FileExists(txtArchivoSele.Content)
        If Not MblnEsValido Then
            If String.IsNullOrEmpty(txtArchivoSele.Content) Then
                MstrMens = "Aún no se ha seleccionado un Reporte!"
            Else
                MstrMens = "El archivo no existe en este Equipo!."
            End If
        Else
            MblnEsValido = txtArchivoSele.Content.ToString.EndsWith(".rpt")
            If Not MblnEsValido Then
                MstrMens = "El archivo seleccionado no es a un Reporte valido!"
            End If
        End If
        SMuestreMensaje()
    End Sub
#End Region
#Region "Eventos de la Ventana"
    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        Dim lelmElemento As FrameworkElement = CType(e.Source, FrameworkElement)
        If TypeOf lelmElemento Is Button Then
            Mouse.OverrideCursor = Cursors.Wait
            Select Case lelmElemento.Name
                Case "bttSeleccionar"
                    SSeleccioneArchivo()
                Case "bttAceptar"
                    SAbraReporte()
                Case "bttCancelar"
                    Me.Close()
            End Select
            Mouse.OverrideCursor = Cursors.Arrow
        End If
    End Sub
#End Region
#Region "Procedimientos"
    Private Sub SSeleccioneArchivo()
        Dim lblnNohayError = False
        Dim lofdArchivo As New OpenFileDialog With {
            .DefaultExt = ".rpt",
            .Filter = My.Resources.TipoArchivoRep,
            .InitialDirectory = GstrTrayReportes
        }
        Try
            Dim lblnOk As Boolean = lofdArchivo.ShowDialog
            If lblnOk Then
                txtArchivoSele.Content = lofdArchivo.FileName
            End If
            lblnNohayError = True
        Catch ex As Exception
            MstrMens = ex.Message
            MstrMensEx = ex.ToString
        Finally
            If lblnNohayError Then
                SValide()
            Else
                SMuestreMensaje()
            End If
        End Try
    End Sub
    Private Sub SAbraReporte()
        If MblnEsValido Then
            Dim lobjRep = New ClsRepOrionCop(GCOBJREGISTRO)
            MstrMens = lobjRep.SAbraReporte(txtArchivoSele.Content)
        Else
            MstrMens = "No ha sido seleccionado aún un Reporte!"
        End If
        SMuestreMensaje()
    End Sub
    Private Sub SMuestreMensaje()
        If Not String.IsNullOrEmpty(MstrMens) Then
            MsgBox(MstrMens, vbOKOnly, "Información")
        End If
        If Not String.IsNullOrEmpty(MstrMensEx) Then
            ClsPanorama.SEscribaArchivoError(MstrMensEx)
        End If
    End Sub
#End Region
End Class