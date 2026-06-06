Imports System.Diagnostics
Public Class WinRegistroCopropiedad
    Private MstrMensaje As String = String.Empty
    Private MblnAceptoContrato As Boolean
    Friend DblNit As Long = 0
    Friend Sub New()
        InitializeComponent()
        GblnOK = False
    End Sub
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        txtNit.IsEnabled = False
        bttAceptar.IsEnabled = False
    End Sub
    Private Sub SMuestreMensaje()
        If Not IsNothing(lblNotifica) Then
            If String.IsNullOrEmpty(MstrMensaje) Then
                lblNotifica.Background = System.Windows.Media.Brushes.Transparent
                lblNotifica.Content = String.Empty
            Else
                lblNotifica.Background = System.Windows.Media.Brushes.LightGray
                lblNotifica.Foreground = System.Windows.Media.Brushes.Blue
                lblNotifica.Content = MstrMensaje
                txtNit.SelectAll()
            End If
        End If
    End Sub
    Private Sub SHabiliteCtls()
        bttAceptar.IsEnabled = False
        bttCancelar.IsEnabled = False
        If MblnAceptoContrato Then
            txtNit.IsEnabled = True
            txtNit.Focus()
        Else
            txtNit.Text = String.Empty
            txtNit.IsEnabled = False
            bttAceptar.IsEnabled = False
            bttCancelar.IsEnabled = True
            bttCancelar.Focus()
        End If
    End Sub
    Private Sub BttCancelar_Click(sender As Object, e As RoutedEventArgs) Handles bttCancelar.Click
        GblnOK = False
        Close()
    End Sub
    Private Sub BttAceptar_Click(sender As Object, e As RoutedEventArgs) Handles bttAceptar.Click
        DblNit = txtNit.Text
        GblnOK = True
        Close()
    End Sub
    Private Sub BttAbrirContrato_Click(sender As Object, e As RoutedEventArgs) Handles _
            bttAbrirContrato.Click
        Dim lstrMens = String.Empty
        Dim lstrArchivo = GstrTrayDat & "ContratoDelClienteOrionPlus.pdf"
        If My.Computer.FileSystem.FileExists(lstrArchivo) Then
            Dim lpdfProcess As New Process
            lpdfProcess.StartInfo.UseShellExecute = True
            lpdfProcess.StartInfo.FileName = lstrArchivo
            Try
                lpdfProcess.Start()
                MstrMensaje = "Archivo abierto exitosamente!"
            Catch ex As Exception
                MstrMensaje = "Error al abrir el archivo " & ex.Message
            End Try
        Else
            lstrMens = "Parece que hubo un problema en la instalación." & vbCrLf &
                    "No existe el archivo que contiene el Contrato! Por favor comuniquese con " &
                    "soporte de OPTIMUSFT S.A.S. al teléfono " & Chr(34) & "311 630 0406" & Chr(34)
        End If
        If Not String.IsNullOrEmpty(lstrMens) Then
            MsgBox(lstrMens, vbOKOnly, "Mensaje")
        End If
    End Sub
    Private Sub Chk_Click(sender As Object, e As RoutedEventArgs) Handles _
            chkAceptarCon.Click
        MblnAceptoContrato = chkAceptarCon.IsChecked
        SHabiliteCtls()
    End Sub
    Private Sub Txt_TextChanged(sender As Object, e As TextChangedEventArgs) Handles _
            txtNit.TextChanged
        Dim lstrValor = txtNit.Text
        MstrMensaje = String.Empty
        If Not IsNumeric(lstrValor) Then
            MstrMensaje = "INFORMACION: El Nit debe ser numérico!"
            bttAceptar.IsEnabled = False
        ElseIf txtNit.Text.Length >= 9 Then
            bttAceptar.IsEnabled = True
        End If
        SMuestreMensaje()
    End Sub
    Private Sub TxtNit_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtNit.LostFocus
        MblnAceptoContrato = False
        If IsNumeric(txtNit.Text) Then
            Dim ldblNit As Double = txtNit.Text
            If ldblNit = GobjPanorama.ObjCarpetaActual.ObjCentroUtilidadActual.
                        ObjIdTerceroCentroUtilDbl.ObjValorPro Then
                MstrMensaje = String.Empty
                bttAceptar.IsEnabled = True
                bttAceptar.Focus()
            Else
                MstrMensaje = "El NIT registrado no coincide con el registrado en nuestra " &
                        "base de datos!"
            End If
        Else
            MstrMensaje = "INFORMACION: El Nit debe ser numérico!"
        End If
        If Not String.IsNullOrEmpty(MstrMensaje) Then
            SMuestreMensaje()
        End If
    End Sub
End Class
