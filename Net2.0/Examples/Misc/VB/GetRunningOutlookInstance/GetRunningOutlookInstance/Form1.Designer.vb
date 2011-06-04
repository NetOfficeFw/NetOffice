<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.textBoxLog = New System.Windows.Forms.TextBox
        Me.buttonStart = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'textBoxLog
        '
        Me.textBoxLog.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.textBoxLog.Location = New System.Drawing.Point(23, 50)
        Me.textBoxLog.Multiline = True
        Me.textBoxLog.Name = "textBoxLog"
        Me.textBoxLog.Size = New System.Drawing.Size(227, 99)
        Me.textBoxLog.TabIndex = 3
        '
        'buttonStart
        '
        Me.buttonStart.Location = New System.Drawing.Point(23, 21)
        Me.buttonStart.Name = "buttonStart"
        Me.buttonStart.Size = New System.Drawing.Size(227, 23)
        Me.buttonStart.TabIndex = 2
        Me.buttonStart.Text = "Start"
        Me.buttonStart.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(273, 171)
        Me.Controls.Add(Me.textBoxLog)
        Me.Controls.Add(Me.buttonStart)
        Me.Name = "Form1"
        Me.Text = "Get running Outlook Application"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents textBoxLog As System.Windows.Forms.TextBox
    Private WithEvents buttonStart As System.Windows.Forms.Button

End Class
