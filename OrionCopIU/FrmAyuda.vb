Public Class FrmAyuda
#Region "Constructor"
    Private ReadOnly MWinPadre As ClsFormInterface = Nothing
    Private MblnUsuarioCerro As Boolean = True
    Sub New(awinPadre As ClsFormInterface)
        InitializeComponent()
        MWinPadre = awinPadre
    End Sub
#End Region
    Friend WriteOnly Property StrTitulo As String
        Set(value As String)
            txtProcedimiento.Text = value
        End Set
    End Property

    Friend WriteOnly Property StrMensaje As String
        Set(value As String)
            txtAyuda.Text = value
        End Set
    End Property

    Friend Sub SCierre()
        txtAyuda.Text = String.Empty
        MblnUsuarioCerro = False
        Close()
    End Sub

    Private Sub FrmAyuda_FormClosing(sender As Object, e As Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        If MWinPadre IsNot Nothing Then
            If MblnUsuarioCerro Then
                EnuEstadoAyuda = EnuEstadoAyudaDef.EnuOff
                If MWinPadre.EnuIdVentana = EnuIdVentanaDef.EnuParametrizacion Then
                    Dim lwinPara As WinParametrizacion = MWinPadre
                    lwinPara.SCerroAyuda()
                End If
            End If
        End If
    End Sub

    Private Sub FrmAyuda_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim scr As Forms.Screen = Forms.Screen.FromPoint(Location)
        Location = New System.Drawing.Point(scr.WorkingArea.Right - Width, scr.WorkingArea.Top)
    End Sub
End Class