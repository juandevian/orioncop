Public Class WinParametrizacionInteres
    Inherits Window
#Region "Definiciones"
    Private Const MCSTRTITULO As String = "Interés de mora (IBC)"
    Private MlstCertificados As New List(Of StcIbcCertificado)
    Private MstcVigente As StcIbcCertificado
    Private MblnHayIbc As Boolean = False
    Private MblnCargando As Boolean = False
#End Region
#Region "Constructores"
    Public Sub New()
        InitializeComponent()
        GblnOK = False
    End Sub
#End Region
#Region "Carga y presentación"
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        MblnCargando = True
        Try
            SCargueCertificadoVigente()
            Dim lenuModo = CType(CInt(GobjParametros.ObjModoInteresByt.ObjValorPro), EnuModoInteres)
            optNinguno.IsChecked = (lenuModo = EnuModoInteres.None)
            optFijo.IsChecked = (lenuModo = EnuModoInteres.EnuFijo)
            optVariable.IsChecked = (lenuModo = EnuModoInteres.EnuVariable)
            txtTasaFija.Text = FstrPorcentajeDigitable(CDbl(GobjParametros.ObjTasaFijaDeseadaDbl.ObjValorPro))
            txtFactor.Text = CDbl(GobjParametros.ObjFactorVariableDbl.ObjValorPro).ToString("0.####")
        Finally
            MblnCargando = False
        End Try
        SRefresque()
    End Sub

    Private Sub SCargueCertificadoVigente()
        MlstCertificados = New ClsIbcAlmacenBd().FlstCertificados()
        MblnHayIbc = ClsIbcCalculo.FblnCertificadoVigente(MlstCertificados, Date.Today, MstcVigente)
        If MblnHayIbc Then
            lblIbcVigente.Text = "IBC vigente: " & Format(MstcVigente.DblIbc, "#0.00%") & " E.A. (certificado " &
                    MstcVigente.StrIdfile & ", del " & Format(MstcVigente.DtmFechaDesde, "yyyy-MM-dd") &
                    " al " & Format(MstcVigente.DtmFechaHasta, "yyyy-MM-dd") & ")"
            lblTope.Text = "Tope legal (1,5 × IBC): " & Format(ClsIbcCalculo.FdblTopeMaximo(MstcVigente.DblIbc),
                    "#0.00%") & " anual. No se puede cobrar un interés superior."
        Else
            lblIbcVigente.Text = "No hay un certificado de IBC vigente guardado en este equipo. " &
                    "Use «Sincronizar ahora» o Herramientas > Consulta IBC."
            lblTope.Text = "El tope legal (1,5 × IBC) se validará al sincronizar."
        End If
    End Sub

    Private Shared Function FstrPorcentajeDigitable(adblFraccion As Double) As String
        Return (adblFraccion * 100).ToString("0.####")
    End Function

    Private Function FenuModoSeleccionado() As EnuModoInteres
        If optFijo.IsChecked = True Then Return EnuModoInteres.EnuFijo
        If optVariable.IsChecked = True Then Return EnuModoInteres.EnuVariable
        Return EnuModoInteres.None
    End Function

    ''' <summary>Muestra la tasa anual y mensual que resultan de lo digitado, sin mensajes ni cambios.</summary>
    Private Sub SRefresque()
        If MblnCargando Then Return
        txtTasaFija.IsEnabled = (optFijo.IsChecked = True)
        txtFactor.IsEnabled = (optVariable.IsChecked = True)
        lblAnual.Text = String.Empty
        lblMensual.Text = String.Empty
        lblAviso.Text = String.Empty
        Dim ldblTasa As Double = 0
        Select Case FenuModoSeleccionado()
            Case EnuModoInteres.None
                lblAviso.Text = "Sin parametrizar: el Cierre de mes usa la tasa registrada manualmente en Tasas de Mora."
                Return
            Case EnuModoInteres.EnuFijo
                Dim ldblDeseada As Double
                If Not ClsIbcCalculo.FblnTryParsePorcentaje(txtTasaFija.Text, ldblDeseada) Then Return
                ldblTasa = ldblDeseada
                If MblnHayIbc Then
                    ldblTasa = ClsIbcCalculo.FdblTasaFijaEfectiva(ldblDeseada, MstcVigente.DblIbc)
                    If ldblTasa < ldblDeseada Then
                        lblAviso.Text = "La tasa deseada (" & Format(ldblDeseada, "#0.00%") & ") supera el tope; " &
                                "se cobrará el tope vigente. La tasa deseada se conserva y se volverá a aplicar " &
                                "cuando el tope lo permita."
                    End If
                End If
            Case EnuModoInteres.EnuVariable
                Dim ldblFactor As Double
                If Not ClsIbcCalculo.FblnTryParseFactor(txtFactor.Text, ldblFactor) OrElse
                        Not ClsIbcCalculo.FblnFactorValido(ldblFactor) Then Return
                If Not MblnHayIbc Then
                    lblAviso.Text = "La tasa se calculará con el IBC vigente al sincronizar."
                    Return
                End If
                ldblTasa = ClsIbcCalculo.FdblTasaVariable(MstcVigente.DblIbc, ldblFactor)
        End Select
        lblAnual.Text = "Tasa anual: " & Format(ldblTasa, "#0.00%")
        lblMensual.Text = "Equivalente mensual (anual / 12): " &
                Format(ClsIbcCalculo.FdblMensualParaMostrar(ldblTasa), "#0.00%")
    End Sub
#End Region
#Region "Validación de lo digitado"
    ''' <summary>
    ''' Valida la tasa fija digitada. Si supera el tope legal, avisa y la deja en el máximo permitido.
    ''' </summary>
    Private Function FblnTasaFijaValida(ByRef adblTasa As Double) As Boolean
        If Not ClsIbcCalculo.FblnTryParsePorcentaje(txtTasaFija.Text, adblTasa) Then
            MsgBox("Digite la tasa fija anual como un porcentaje, por ejemplo 24 o 24,5.",
                    MsgBoxStyle.Exclamation, MCSTRTITULO)
            Return False
        End If
        If adblTasa <= 0 Then
            MsgBox("La tasa fija debe ser mayor que 0.", MsgBoxStyle.Exclamation, MCSTRTITULO)
            Return False
        End If
        If MblnHayIbc Then
            Dim lblnSeAjusto As Boolean
            Dim ldblPermitida = ClsIbcCalculo.FdblTasaFijaPermitida(adblTasa, MstcVigente.DblIbc, lblnSeAjusto)
            If lblnSeAjusto Then
                MsgBox("No es posible cobrar un interés superior a 1,5 veces el interés bancario corriente (" &
                        Format(ClsIbcCalculo.FdblTopeMaximo(MstcVigente.DblIbc), "#0.00%") &
                        "). Se ajustó al máximo permitido.", MsgBoxStyle.Exclamation, MCSTRTITULO)
                adblTasa = ldblPermitida
                txtTasaFija.Text = FstrPorcentajeDigitable(adblTasa)
            End If
        ElseIf adblTasa > 1 Then
            MsgBox("La tasa fija anual no puede ser mayor que 100 %.", MsgBoxStyle.Exclamation, MCSTRTITULO)
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' Valida el factor digitado. Si supera el máximo legal, avisa y lo deja en el máximo (1,5).
    ''' </summary>
    Private Function FblnFactorValido(ByRef adblFactor As Double) As Boolean
        If Not ClsIbcCalculo.FblnTryParseFactor(txtFactor.Text, adblFactor) Then
            MsgBox("Digite el factor como un número, por ejemplo 1,3.", MsgBoxStyle.Exclamation, MCSTRTITULO)
            Return False
        End If
        Dim lblnSeAjusto As Boolean
        Dim ldblPermitido = ClsIbcCalculo.FdblFactorPermitido(adblFactor, lblnSeAjusto)
        If lblnSeAjusto Then
            MsgBox("El factor no puede superar " & ClsIbcCalculo.CDBLFACTORMAXIMO.ToString("0.0#") &
                    " (Ley 675 de 2001). Se ajustó al máximo permitido.", MsgBoxStyle.Exclamation, MCSTRTITULO)
            adblFactor = ldblPermitido
            txtFactor.Text = adblFactor.ToString("0.####")
        End If
        If Not ClsIbcCalculo.FblnFactorValido(adblFactor) Then
            MsgBox("El factor debe ser mayor que 0.", MsgBoxStyle.Exclamation, MCSTRTITULO)
            Return False
        End If
        Return True
    End Function

    ''' <summary>Valida y guarda la parametrización en el centro de utilidad.</summary>
    Private Function FblnGuarde() As Boolean
        Dim lenuModo = FenuModoSeleccionado()
        Dim ldblTasaFija As Double = 0, ldblFactor As Double = 0
        If lenuModo = EnuModoInteres.EnuFijo AndAlso Not FblnTasaFijaValida(ldblTasaFija) Then Return False
        If lenuModo = EnuModoInteres.EnuVariable AndAlso Not FblnFactorValido(ldblFactor) Then Return False
        Try
            GobjParametros.SRegistreParametrosInteres(lenuModo, ldblTasaFija, ldblFactor)
        Catch ex As Exception
            MsgBox("No se pudo guardar la parametrización del interés de mora: " & ex.Message,
                    MsgBoxStyle.Critical, MCSTRTITULO)
            Return False
        End Try
        SRefresque()
        Return True
    End Function
#End Region
#Region "Eventos"
    Private Sub Modo_Checked(sender As Object, e As RoutedEventArgs) Handles optNinguno.Checked,
            optFijo.Checked, optVariable.Checked
        SRefresque()
    End Sub

    Private Sub TxtTasaFija_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtTasaFija.LostFocus
        If MblnCargando OrElse optFijo.IsChecked <> True Then Return
        Dim ldblTasa As Double
        Dim lblnOk = FblnTasaFijaValida(ldblTasa)
        SRefresque()
    End Sub

    Private Sub TxtFactor_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtFactor.LostFocus
        If MblnCargando OrElse optVariable.IsChecked <> True Then Return
        Dim ldblFactor As Double
        Dim lblnOk = FblnFactorValido(ldblFactor)
        SRefresque()
    End Sub

    Private Sub BttGuardar_Click(sender As Object, e As RoutedEventArgs) Handles bttGuardar.Click
        If FblnGuarde() Then
            GblnOK = True
            Close()
        End If
    End Sub

    Private Sub BttCancelar_Click(sender As Object, e As RoutedEventArgs) Handles bttCancelar.Click
        GblnOK = False
        Close()
    End Sub

    ''' <summary>Guarda y actualiza OriTasasMora ahora, consultando la API del IBC.</summary>
    Private Sub BttSincronizar_Click(sender As Object, e As RoutedEventArgs) Handles bttSincronizar.Click
        If Not FblnGuarde() Then Return
        Dim lstrMens = String.Empty
        Dim lblnOk As Boolean
        Mouse.OverrideCursor = Cursors.Wait
        Try
            lblnOk = ClsOrionCop.FblnSincronizaIbcManual(lstrMens)
        Catch ex As Exception
            lblnOk = False
            lstrMens = ex.Message
        Finally
            Mouse.OverrideCursor = Nothing
        End Try
        If lblnOk Then
            SCargueCertificadoVigente()
            SRefresque()
            MsgBox("La tasa de mora quedó al día. Tasa anual vigente: " &
                    Format(GobjParametros.FdblTasaMoraFecha(Date.Today.AddDays(1)), "#0.00%") & ".",
                    MsgBoxStyle.Information, MCSTRTITULO)
        Else
            MsgBox(lstrMens, MsgBoxStyle.Exclamation, MCSTRTITULO)
        End If
    End Sub
#End Region
End Class
