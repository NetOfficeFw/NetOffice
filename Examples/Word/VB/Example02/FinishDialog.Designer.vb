<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FinishDialog
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
        Me.labelDocumentPath = New System.Windows.Forms.Label
        Me.labelMessage = New System.Windows.Forms.Label
        Me.buttonOpenDocument = New System.Windows.Forms.Button
        Me.buttonClose = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'labelDocumentPath
        '
        Me.labelDocumentPath.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.labelDocumentPath.Location = New System.Drawing.Point(12, 40)
        Me.labelDocumentPath.Name = "labelDocumentPath"
        Me.labelDocumentPath.Size = New System.Drawing.Size(321, 41)
        Me.labelDocumentPath.TabIndex = 9
        Me.labelDocumentPath.Text = "labelDocumentPath"
        '
        'labelMessage
        '
        Me.labelMessage.AutoSize = True
        Me.labelMessage.Location = New System.Drawing.Point(12, 18)
        Me.labelMessage.Name = "labelMessage"
        Me.labelMessage.Size = New System.Drawing.Size(72, 13)
        Me.labelMessage.TabIndex = 8
        Me.labelMessage.Text = "labelMessage"
        '
        'buttonOpenDocument
        '
        Me.buttonOpenDocument.Location = New System.Drawing.Point(13, 84)
        Me.buttonOpenDocument.Name = "buttonOpenDocument"
        Me.buttonOpenDocument.Size = New System.Drawing.Size(102, 22)
        Me.buttonOpenDocument.TabIndex = 7
        Me.buttonOpenDocument.Text = "Open Document"
        Me.buttonOpenDocument.UseVisualStyleBackColor = True
        '
        'buttonClose
        '
        Me.buttonClose.Location = New System.Drawing.Point(231, 84)
        Me.buttonClose.Name = "buttonClose"
        Me.buttonClose.Size = New System.Drawing.Size(102, 22)
        Me.buttonClose.TabIndex = 6
        Me.buttonClose.Text = "Ok"
        Me.buttonClose.UseVisualStyleBackColor = True
        '
        'FinishDialog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(352, 127)
        Me.Controls.Add(Me.labelDocumentPath)
        Me.Controls.Add(Me.labelMessage)
        Me.Controls.Add(Me.buttonOpenDocument)
        Me.Controls.Add(Me.buttonClose)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FinishDialog"
        Me.Padding = New System.Windows.Forms.Padding(9)
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Example02"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents labelDocumentPath As System.Windows.Forms.Label
    Private WithEvents labelMessage As System.Windows.Forms.Label
    Private WithEvents buttonOpenDocument As System.Windows.Forms.Button
    Private WithEvents buttonClose As System.Windows.Forms.Button

End Class
