<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SamplePane
    Inherits System.Windows.Forms.UserControl

    'UserControl überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.UsageTimer = New System.Windows.Forms.Timer(Me.components)
        Me.UsageLabel = New System.Windows.Forms.Label()
        Me.UsageBar = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'UsageTimer
        '
        Me.UsageTimer.Interval = 400
        '
        'UsageLabel
        '
        Me.UsageLabel.AutoSize = True
        Me.UsageLabel.BackColor = System.Drawing.Color.Transparent
        Me.UsageLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.UsageLabel.ForeColor = System.Drawing.Color.Blue
        Me.UsageLabel.Location = New System.Drawing.Point(126, 8)
        Me.UsageLabel.Name = "UsageLabel"
        Me.UsageLabel.Size = New System.Drawing.Size(48, 13)
        Me.UsageLabel.TabIndex = 9
        Me.UsageLabel.Text = "<Empty>"
        Me.UsageLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'UsageBar
        '
        Me.UsageBar.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UsageBar.Location = New System.Drawing.Point(0, 0)
        Me.UsageBar.Name = "UsageBar"
        Me.UsageBar.Size = New System.Drawing.Size(300, 30)
        Me.UsageBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.UsageBar.TabIndex = 10
        '
        'SamplePane
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.UsageLabel)
        Me.Controls.Add(Me.UsageBar)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "SamplePane"
        Me.Size = New System.Drawing.Size(300, 30)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private WithEvents UsageTimer As Windows.Forms.Timer
    Private WithEvents UsageLabel As Windows.Forms.Label
    Private WithEvents UsageBar As Windows.Forms.ProgressBar
End Class
