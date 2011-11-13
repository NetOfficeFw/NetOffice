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
        Me.richTextBoxResult = New System.Windows.Forms.RichTextBox
        Me.richTextBoxInfo = New System.Windows.Forms.RichTextBox
        Me.buttonStartExample = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'richTextBoxResult
        '
        Me.richTextBoxResult.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.richTextBoxResult.Font = New System.Drawing.Font("Lucida Console", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.richTextBoxResult.Location = New System.Drawing.Point(24, 171)
        Me.richTextBoxResult.Name = "richTextBoxResult"
        Me.richTextBoxResult.Size = New System.Drawing.Size(546, 118)
        Me.richTextBoxResult.TabIndex = 15
        Me.richTextBoxResult.Text = "" & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'richTextBoxInfo
        '
        Me.richTextBoxInfo.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.richTextBoxInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.richTextBoxInfo.Location = New System.Drawing.Point(24, 64)
        Me.richTextBoxInfo.Name = "richTextBoxInfo"
        Me.richTextBoxInfo.Size = New System.Drawing.Size(546, 83)
        Me.richTextBoxInfo.TabIndex = 14
        Me.richTextBoxInfo.Text = "This tutorial shows how to use the EntityIsAvailable feature." & Global.Microsoft.VisualBasic.ChrW(10) & "you can check at ru" & _
            "ntime for specific entity support." & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'buttonStartExample
        '
        Me.buttonStartExample.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.buttonStartExample.Location = New System.Drawing.Point(24, 19)
        Me.buttonStartExample.Name = "buttonStartExample"
        Me.buttonStartExample.Size = New System.Drawing.Size(546, 30)
        Me.buttonStartExample.TabIndex = 13
        Me.buttonStartExample.Text = "Start example"
        Me.buttonStartExample.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(604, 316)
        Me.Controls.Add(Me.richTextBoxResult)
        Me.Controls.Add(Me.richTextBoxInfo)
        Me.Controls.Add(Me.buttonStartExample)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.Text = "Tutorial08"
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents richTextBoxResult As System.Windows.Forms.RichTextBox
    Private WithEvents richTextBoxInfo As System.Windows.Forms.RichTextBox
    Private WithEvents buttonStartExample As System.Windows.Forms.Button

End Class
