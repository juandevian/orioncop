Public Class WinIncrementoCA
    Private MstrMensaje As String = String.Empty
    Private MblnPoblandoCbo As Boolean = False
    Friend DblIncCA As Double = 0
    Friend BlnProcesar As Boolean = False
    Friend EnuTipoCalculoCuotaAno As EnuTipoBaseCalculo = EnuTipoBaseCalculo.None
    Public Sub New()
        InitializeComponent()
        GblnOK = False
    End Sub
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        If GobjParametros.ObjAnoActual IsNot Nothing Then
            EnuTipoCalculoCuotaAno =
                    GobjParametros.ObjAnoActual.ObjTipoCalculoCuotaByt.ObjValorPro
            If EnuTipoCalculoCuotaAno = EnuTipoBaseCalculo.EnuImportadas Then
                EnuTipoCalculoCuotaAno = EnuTipoBaseCalculo.EnuCoeficientePro
            End If
        End If
        SPuebleComboBoxes()
        CboBaseCalcCA.SelectedIndex = EnuTipoCalculoCuotaAno
        txtIncremento.Focus()
        txtIncremento.SelectAll()
    End Sub
    Private Sub Txt_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtIncremento.LostFocus
        Dim lstrValor As String
        If Not String.IsNullOrEmpty(MstrMensaje) Then
            DblIncCA = 0
        Else
            DblIncCA = FdblIncrementoCA() / 100
        End If
        lstrValor = Format(DblIncCA, "p")
        txtIncremento.Text = lstrValor
    End Sub
    Private Sub Txt_TextChanged(sender As Object, e As TextChangedEventArgs) Handles txtIncremento.TextChanged
        Dim lstrValor = txtIncremento.Text
        MstrMensaje = String.Empty
        If String.IsNullOrEmpty(lstrValor) Then
            lstrValor = "0"
        End If
        If lstrValor.EndsWith("%") Then
            lstrValor = lstrValor.Substring(0, txtIncremento.Text.Length - 1)
        End If
        If Not IsNumeric(lstrValor) Then
            MstrMensaje = "INFORMACION: El dato debe ser numérico!"
            bttAceptar.IsEnabled = False
        Else
            bttAceptar.IsEnabled = True
        End If
        SMuestreMensaje()
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
                txtIncremento.SelectAll()
            End If
        End If
    End Sub
    Private Function FdblIncrementoCA() As Double
        Dim lstrValor = txtIncremento.Text
        If String.IsNullOrEmpty(lstrValor) Then
            lstrValor = "0"
        End If
        If lstrValor.EndsWith("%") Then
            lstrValor = lstrValor.Substring(0, txtIncremento.Text.Length - 1)
        End If
        Return CType(lstrValor, Double)
    End Function
    Private Sub BttCancelar_Click(sender As Object, e As RoutedEventArgs) Handles bttCancelar.Click
        BlnProcesar = False
        Close()
    End Sub
    Private Sub BttAceptar_Click(sender As Object, e As RoutedEventArgs) Handles bttAceptar.Click
        If Not GobjParametros.ObjAnoActual.ObjModuloPorServicioBln.ObjValorPro Then
            Dim lstrmens = "Está de acuerdo con que los valores de los Modulos de Contribución se " &
                    "incrementen en el " & txtIncremento.Text & "?" & vbCrLf & "Si no está de " &
                    "acuerdo, debe ingresar los valores con los cuales contribuyen los Módulos a " &
                    "las Cuotas de Administración!"
            GblnOK = MsgBox(lstrmens, vbYesNo, "Confirmar incremento") = MsgBoxResult.Yes
        Else
            GblnOK = True
        End If
        BlnProcesar = True
        Close()
    End Sub
    Private Sub CboBaseCalcCA_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) _
            Handles CboBaseCalcCA.SelectionChanged
        If TypeOf sender Is ComboBox Then
            If Not MblnPoblandoCbo Then
                MstrMensaje = String.Empty
                EnuTipoCalculoCuotaAno = CboBaseCalcCA.SelectedIndex
                If EnuTipoCalculoCuotaAno = EnuTipoBaseCalculo.EnuImportadas Then
                    MstrMensaje = "No es posible importar las cuotas!"
                    bttAceptar.IsEnabled = False
                ElseIf EnuTipoCalculoCuotaAno = EnuTipoBaseCalculo.None Then
                    MstrMensaje = "INFORMACION: Dato no valido!"
                    bttAceptar.IsEnabled = False
                Else
                    bttAceptar.IsEnabled = True
                End If
                SMuestreMensaje()
            End If
        End If
    End Sub
    Private Sub SPuebleComboBoxes()
        Dim ldrwTiposBaseCalc = ClsOrionCop.FdrwConstantesOri(
                EnuGrupoConstantesOriDef.EnuTipoBaseCalculo)
        MblnPoblandoCbo = True
        SPuebleComboBox(ldrwTiposBaseCalc, CboBaseCalcCA)
        MblnPoblandoCbo = False
    End Sub
End Class
