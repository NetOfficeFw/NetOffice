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
        Me.panelSelection = New System.Windows.Forms.Panel
        Me.radioButton8 = New System.Windows.Forms.RadioButton
        Me.radioButton7 = New System.Windows.Forms.RadioButton
        Me.radioButton6 = New System.Windows.Forms.RadioButton
        Me.radioButton5 = New System.Windows.Forms.RadioButton
        Me.radioButton4 = New System.Windows.Forms.RadioButton
        Me.radioButton3 = New System.Windows.Forms.RadioButton
        Me.radioButton2 = New System.Windows.Forms.RadioButton
        Me.radioButton1 = New System.Windows.Forms.RadioButton
        Me.textBox1 = New System.Windows.Forms.TextBox
        Me.button1 = New System.Windows.Forms.Button
        Me.panelSelection.SuspendLayout()
        Me.SuspendLayout()
        '
        'panelSelection
        '
        Me.panelSelection.Controls.Add(Me.radioButton8)
        Me.panelSelection.Controls.Add(Me.radioButton7)
        Me.panelSelection.Controls.Add(Me.radioButton6)
        Me.panelSelection.Controls.Add(Me.radioButton5)
        Me.panelSelection.Controls.Add(Me.radioButton4)
        Me.panelSelection.Controls.Add(Me.radioButton3)
        Me.panelSelection.Controls.Add(Me.radioButton2)
        Me.panelSelection.Controls.Add(Me.radioButton1)
        Me.panelSelection.Location = New System.Drawing.Point(404, 19)
        Me.panelSelection.Name = "panelSelection"
        Me.panelSelection.Size = New System.Drawing.Size(154, 202)
        Me.panelSelection.TabIndex = 18
        '
        'radioButton8
        '
        Me.radioButton8.AutoSize = True
        Me.radioButton8.Location = New System.Drawing.Point(3, 180)
        Me.radioButton8.Name = "radioButton8"
        Me.radioButton8.Size = New System.Drawing.Size(111, 17)
        Me.radioButton8.TabIndex = 15
        Me.radioButton8.Text = "xlDialogApplyStyle"
        Me.radioButton8.UseVisualStyleBackColor = True
        '
        'radioButton7
        '
        Me.radioButton7.AutoSize = True
        Me.radioButton7.Location = New System.Drawing.Point(3, 157)
        Me.radioButton7.Name = "radioButton7"
        Me.radioButton7.Size = New System.Drawing.Size(131, 17)
        Me.radioButton7.TabIndex = 14
        Me.radioButton7.Text = "xlDialogFormatNumber"
        Me.radioButton7.UseVisualStyleBackColor = True
        '
        'radioButton6
        '
        Me.radioButton6.AutoSize = True
        Me.radioButton6.Location = New System.Drawing.Point(3, 134)
        Me.radioButton6.Name = "radioButton6"
        Me.radioButton6.Size = New System.Drawing.Size(120, 17)
        Me.radioButton6.TabIndex = 13
        Me.radioButton6.Text = "xlDialogPrinterSetup"
        Me.radioButton6.UseVisualStyleBackColor = True
        '
        'radioButton5
        '
        Me.radioButton5.AutoSize = True
        Me.radioButton5.Location = New System.Drawing.Point(3, 111)
        Me.radioButton5.Name = "radioButton5"
        Me.radioButton5.Size = New System.Drawing.Size(96, 17)
        Me.radioButton5.TabIndex = 12
        Me.radioButton5.Text = "xlDialogSearch"
        Me.radioButton5.UseVisualStyleBackColor = True
        '
        'radioButton4
        '
        Me.radioButton4.AutoSize = True
        Me.radioButton4.Location = New System.Drawing.Point(3, 88)
        Me.radioButton4.Name = "radioButton4"
        Me.radioButton4.Size = New System.Drawing.Size(122, 17)
        Me.radioButton4.TabIndex = 11
        Me.radioButton4.Text = "xlDialogGallery3dBar"
        Me.radioButton4.UseVisualStyleBackColor = True
        '
        'radioButton3
        '
        Me.radioButton3.AutoSize = True
        Me.radioButton3.Location = New System.Drawing.Point(3, 65)
        Me.radioButton3.Name = "radioButton3"
        Me.radioButton3.Size = New System.Drawing.Size(104, 17)
        Me.radioButton3.TabIndex = 10
        Me.radioButton3.Text = "xlDialogEditColor"
        Me.radioButton3.UseVisualStyleBackColor = True
        '
        'radioButton2
        '
        Me.radioButton2.AutoSize = True
        Me.radioButton2.Location = New System.Drawing.Point(3, 42)
        Me.radioButton2.Name = "radioButton2"
        Me.radioButton2.Size = New System.Drawing.Size(83, 17)
        Me.radioButton2.TabIndex = 9
        Me.radioButton2.Text = "xlDialogFont"
        Me.radioButton2.UseVisualStyleBackColor = True
        '
        'radioButton1
        '
        Me.radioButton1.AutoSize = True
        Me.radioButton1.Checked = True
        Me.radioButton1.Location = New System.Drawing.Point(3, 19)
        Me.radioButton1.Name = "radioButton1"
        Me.radioButton1.Size = New System.Drawing.Size(131, 17)
        Me.radioButton1.TabIndex = 8
        Me.radioButton1.TabStop = True
        Me.radioButton1.Text = "xlDialogAddinManager"
        Me.radioButton1.UseVisualStyleBackColor = True
        '
        'textBox1
        '
        Me.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.textBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox1.Location = New System.Drawing.Point(26, 55)
        Me.textBox1.Multiline = True
        Me.textBox1.Name = "textBox1"
        Me.textBox1.Size = New System.Drawing.Size(369, 167)
        Me.textBox1.TabIndex = 17
        Me.textBox1.Text = "This example contains code to work with dialogs." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Start excel and it shows the se" & _
            "lected Dialog" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "and waits for user input. Excel has more than 50 different dialog" & _
            "s, this is only a sample selection."
        '
        'button1
        '
        Me.button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.button1.Location = New System.Drawing.Point(26, 21)
        Me.button1.Name = "button1"
        Me.button1.Size = New System.Drawing.Size(369, 28)
        Me.button1.TabIndex = 16
        Me.button1.Text = "Start example"
        Me.button1.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(587, 242)
        Me.Controls.Add(Me.panelSelection)
        Me.Controls.Add(Me.textBox1)
        Me.Controls.Add(Me.button1)
        Me.Name = "Form1"
        Me.Text = "Example6 - Dialogs"
        Me.panelSelection.ResumeLayout(False)
        Me.panelSelection.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents panelSelection As System.Windows.Forms.Panel
    Private WithEvents radioButton8 As System.Windows.Forms.RadioButton
    Private WithEvents radioButton7 As System.Windows.Forms.RadioButton
    Private WithEvents radioButton6 As System.Windows.Forms.RadioButton
    Private WithEvents radioButton5 As System.Windows.Forms.RadioButton
    Private WithEvents radioButton4 As System.Windows.Forms.RadioButton
    Private WithEvents radioButton3 As System.Windows.Forms.RadioButton
    Private WithEvents radioButton2 As System.Windows.Forms.RadioButton
    Private WithEvents radioButton1 As System.Windows.Forms.RadioButton
    Private WithEvents textBox1 As System.Windows.Forms.TextBox
    Private WithEvents button1 As System.Windows.Forms.Button

End Class
