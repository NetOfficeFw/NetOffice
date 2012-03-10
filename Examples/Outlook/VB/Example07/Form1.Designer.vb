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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.button2 = New System.Windows.Forms.Button
        Me.button1 = New System.Windows.Forms.Button
        Me.label1 = New System.Windows.Forms.Label
        Me.textBoxEvents = New System.Windows.Forms.TextBox
        Me.textBox1 = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'button2
        '
        Me.button2.Enabled = False
        Me.button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.button2.Location = New System.Drawing.Point(10, 51)
        Me.button2.Name = "button2"
        Me.button2.Size = New System.Drawing.Size(507, 28)
        Me.button2.TabIndex = 34
        Me.button2.Text = "Quit Outlook"
        Me.button2.UseVisualStyleBackColor = True
        '
        'button1
        '
        Me.button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.button1.Location = New System.Drawing.Point(11, 17)
        Me.button1.Name = "button1"
        Me.button1.Size = New System.Drawing.Size(507, 28)
        Me.button1.TabIndex = 30
        Me.button1.Text = "Start example"
        Me.button1.UseVisualStyleBackColor = True
        '
        'label1
        '
        Me.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.label1.Location = New System.Drawing.Point(12, 228)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(506, 22)
        Me.label1.TabIndex = 33
        Me.label1.Text = "EventLog"
        '
        'textBoxEvents
        '
        Me.textBoxEvents.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.textBoxEvents.Location = New System.Drawing.Point(12, 253)
        Me.textBoxEvents.Multiline = True
        Me.textBoxEvents.Name = "textBoxEvents"
        Me.textBoxEvents.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.textBoxEvents.Size = New System.Drawing.Size(506, 209)
        Me.textBoxEvents.TabIndex = 32
        '
        'textBox1
        '
        Me.textBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox1.Location = New System.Drawing.Point(11, 85)
        Me.textBox1.Multiline = True
        Me.textBox1.Name = "textBox1"
        Me.textBox1.Size = New System.Drawing.Size(506, 132)
        Me.textBox1.TabIndex = 31
        Me.textBox1.Text = resources.GetString("textBox1.Text")
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(534, 474)
        Me.Controls.Add(Me.button2)
        Me.Controls.Add(Me.button1)
        Me.Controls.Add(Me.label1)
        Me.Controls.Add(Me.textBoxEvents)
        Me.Controls.Add(Me.textBox1)
        Me.Name = "Form1"
        Me.Text = "Example07 - Customize UI"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents button2 As System.Windows.Forms.Button
    Private WithEvents button1 As System.Windows.Forms.Button
    Private WithEvents label1 As System.Windows.Forms.Label
    Private WithEvents textBoxEvents As System.Windows.Forms.TextBox
    Private WithEvents textBox1 As System.Windows.Forms.TextBox

End Class
