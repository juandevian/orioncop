Friend Class FrmReportes
#Region "Definiciones"
    '
#End Region
    Friend Property RcReporte As ReportClass = Nothing
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub FrmReportes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        crvReportes.ReportSource = RcReporte
    End Sub
    Private Sub FrmReportes_Closed(sender As Object, e As EventArgs) Handles MyBase.Closed
        If RcReporte IsNot Nothing Then
            RcReporte.Close()
            RcReporte.Dispose()
        End If
    End Sub
End Class