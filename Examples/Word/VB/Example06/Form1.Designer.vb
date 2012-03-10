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
        Me.label1 = New System.Windows.Forms.Label
        Me.textBoxEvents = New System.Windows.Forms.TextBox
        Me.textBox1 = New System.Windows.Forms.TextBox
        Me.button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'label1
        '
        Me.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.label1.Location = New System.Drawing.Point(12, 183)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(506, 22)
        Me.label1.TabIndex = 25
        Me.label1.Text = "EventLog"
        '
        'textBoxEvents
        '
        Me.textBoxEvents.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.textBoxEvents.Location = New System.Drawing.Point(12, 208)
        Me.textBoxEvents.Multiline = True
        Me.textBoxEvents.Name = "textBoxEvents"
        Me.textBoxEvents.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.textBoxEvents.Size = New System.Drawing.Size(506, 209)
        Me.textBoxEvents.TabIndex = 24
        '
        'textBox1
        '
        Me.textBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox1.Location = New System.Drawing.Point(12, 50)
        Me.textBox1.Multiline = True
        Me.textBox1.Name = "textBox1"
        Me.textBox1.Size = New System.Drawing.Size(507, 124)
        Me.textBox1.TabIndex = 23
        Me.textBox1.Text = "This example contains code to catch events in word." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "example 7 shows you how to c" & _
            "atch click events from gui." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'button1
        '
        Me.button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.button1.Location = New System.Drawing.Point(12, 16)
        Me.button1.Name = "button1"
        Me.button1.Size = New System.Drawing.Size(507, 28)
        Me.button1.TabIndex = 22
        Me.button1.Text = "Start example"
        Me.button1.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(529, 427)
        Me.Controls.Add(Me.label1)
        Me.Controls.Add(Me.textBoxEvents)
        Me.Controls.Add(Me.textBox1)
        Me.Controls.Add(Me.button1)
        Me.Name = "Form1"
        Me.Text = "Example06 - Events"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents label1 As System.Windows.Forms.Label
    Private WithEvents textBoxEvents As System.Windows.Forms.TextBox
    Private WithEvents textBox1 As System.Windows.Forms.TextBox
    Private WithEvents button1 As System.Windows.Forms.Button

End Class
