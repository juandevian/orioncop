<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmReportes
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmReportes))
        Me.crvReportes = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.repFactura1 = New OPT.OrionP.RepOriCop.repFacturaLogo()
        Me.repFactura2 = New OPT.OrionP.RepOriCop.repFacturaLogo()
        Me.SuspendLayout()
        '
        'crvReportes
        '
        Me.crvReportes.ActiveViewIndex = -1
        Me.crvReportes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvReportes.Cursor = System.Windows.Forms.Cursors.Default
        Me.crvReportes.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvReportes.Location = New System.Drawing.Point(0, 0)
        Me.crvReportes.Name = "crvReportes"
        Me.crvReportes.Size = New System.Drawing.Size(1144, 587)
        Me.crvReportes.TabIndex = 0
        Me.crvReportes.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'frmReportes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1144, 587)
        Me.Controls.Add(Me.crvReportes)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmReportes"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents crvReportes As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents repFactura1 As OPT.OrionP.RepOriCop.repFacturaLogo
    Friend WithEvents repFactura2 As OPT.OrionP.RepOriCop.repFacturaLogo
End Class
