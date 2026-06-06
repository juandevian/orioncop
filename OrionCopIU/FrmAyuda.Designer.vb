<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmAyuda
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
        Me.txtAyuda = New System.Windows.Forms.RichTextBox()
        Me.lblProcedimiento = New System.Windows.Forms.Label()
        Me.txtProcedimiento = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtAyuda
        '
        Me.txtAyuda.BackColor = System.Drawing.SystemColors.ControlLight
        Me.txtAyuda.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtAyuda.CausesValidation = False
        Me.txtAyuda.Cursor = System.Windows.Forms.Cursors.Default
        Me.txtAyuda.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAyuda.Location = New System.Drawing.Point(10, 50)
        Me.txtAyuda.Margin = New System.Windows.Forms.Padding(4)
        Me.txtAyuda.Name = "txtAyuda"
        Me.txtAyuda.ReadOnly = True
        Me.txtAyuda.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.txtAyuda.Size = New System.Drawing.Size(900, 500)
        Me.txtAyuda.TabIndex = 0
        Me.txtAyuda.TabStop = False
        Me.txtAyuda.Text = ""
        '
        'lblProcedimiento
        '
        Me.lblProcedimiento.AutoSize = True
        Me.lblProcedimiento.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProcedimiento.Location = New System.Drawing.Point(9, 8)
        Me.lblProcedimiento.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblProcedimiento.Name = "lblProcedimiento"
        Me.lblProcedimiento.Size = New System.Drawing.Size(227, 32)
        Me.lblProcedimiento.TabIndex = 1
        Me.lblProcedimiento.Text = "Procedimiento: "
        '
        'txtProcedimiento
        '
        Me.txtProcedimiento.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProcedimiento.Location = New System.Drawing.Point(231, 9)
        Me.txtProcedimiento.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.txtProcedimiento.Name = "txtProcedimiento"
        Me.txtProcedimiento.Size = New System.Drawing.Size(679, 31)
        Me.txtProcedimiento.TabIndex = 2
        '
        'FrmAyuda
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(914, 553)
        Me.Controls.Add(Me.txtProcedimiento)
        Me.Controls.Add(Me.lblProcedimiento)
        Me.Controls.Add(Me.txtAyuda)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmAyuda"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ayuda"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtAyuda As Forms.RichTextBox
    Friend WithEvents lblProcedimiento As Forms.Label
    Friend WithEvents txtProcedimiento As Forms.Label
End Class
