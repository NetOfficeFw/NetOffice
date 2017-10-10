<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OptionPage
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.SettingsGrid = New System.Windows.Forms.PropertyGrid()
        Me.SuspendLayout()
        '
        'SettingsGrid
        '
        Me.SettingsGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SettingsGrid.Location = New System.Drawing.Point(0, 0)
        Me.SettingsGrid.Name = "SettingsGrid"
        Me.SettingsGrid.PropertySort = System.Windows.Forms.PropertySort.Alphabetical
        Me.SettingsGrid.Size = New System.Drawing.Size(300, 300)
        Me.SettingsGrid.TabIndex = 1
        '
        'OptionPage
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.SettingsGrid)
        Me.Name = "OptionPage"
        Me.Size = New System.Drawing.Size(300, 300)
        Me.ResumeLayout(False)

    End Sub

    Private WithEvents SettingsGrid As Windows.Forms.PropertyGrid
End Class
