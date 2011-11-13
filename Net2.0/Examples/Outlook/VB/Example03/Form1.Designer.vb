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
        Me.textBoxBody = New System.Windows.Forms.TextBox
        Me.textBoxSubject = New System.Windows.Forms.TextBox
        Me.textBoxReciever = New System.Windows.Forms.TextBox
        Me.label3 = New System.Windows.Forms.Label
        Me.label2 = New System.Windows.Forms.Label
        Me.label1 = New System.Windows.Forms.Label
        Me.textBox1 = New System.Windows.Forms.TextBox
        Me.button1 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'textBoxBody
        '
        Me.textBoxBody.Location = New System.Drawing.Point(78, 206)
        Me.textBoxBody.Multiline = True
        Me.textBoxBody.Name = "textBoxBody"
        Me.textBoxBody.Size = New System.Drawing.Size(316, 121)
        Me.textBoxBody.TabIndex = 19
        Me.textBoxBody.Text = "This is a test mail."
        Me.textBoxBody.WordWrap = False
        '
        'textBoxSubject
        '
        Me.textBoxSubject.Location = New System.Drawing.Point(78, 180)
        Me.textBoxSubject.Name = "textBoxSubject"
        Me.textBoxSubject.Size = New System.Drawing.Size(316, 20)
        Me.textBoxSubject.TabIndex = 18
        Me.textBoxSubject.Text = "NetOffice Example Mail"
        '
        'textBoxReciever
        '
        Me.textBoxReciever.Location = New System.Drawing.Point(78, 154)
        Me.textBoxReciever.Name = "textBoxReciever"
        Me.textBoxReciever.Size = New System.Drawing.Size(316, 20)
        Me.textBoxReciever.TabIndex = 17
        Me.textBoxReciever.Text = "public.sebastian@web.de"
        '
        'label3
        '
        Me.label3.AutoSize = True
        Me.label3.Location = New System.Drawing.Point(22, 209)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(31, 13)
        Me.label3.TabIndex = 16
        Me.label3.Text = "Body"
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Location = New System.Drawing.Point(22, 187)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(43, 13)
        Me.label2.TabIndex = 15
        Me.label2.Text = "Subject"
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(22, 157)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(50, 13)
        Me.label1.TabIndex = 14
        Me.label1.Text = "Reciever"
        '
        'textBox1
        '
        Me.textBox1.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.textBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.textBox1.Location = New System.Drawing.Point(12, 57)
        Me.textBox1.Multiline = True
        Me.textBox1.Name = "textBox1"
        Me.textBox1.Size = New System.Drawing.Size(382, 77)
        Me.textBox1.TabIndex = 13
        Me.textBox1.Text = "this example shows you how to send a mail."
        '
        'button1
        '
        Me.button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.button1.Location = New System.Drawing.Point(12, 11)
        Me.button1.Name = "button1"
        Me.button1.Size = New System.Drawing.Size(382, 30)
        Me.button1.TabIndex = 12
        Me.button1.Text = "Start example"
        Me.button1.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(410, 343)
        Me.Controls.Add(Me.textBoxBody)
        Me.Controls.Add(Me.textBoxSubject)
        Me.Controls.Add(Me.textBoxReciever)
        Me.Controls.Add(Me.label3)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.label1)
        Me.Controls.Add(Me.textBox1)
        Me.Controls.Add(Me.button1)
        Me.Name = "Form1"
        Me.Text = "Example03 - Send a mail"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents textBoxBody As System.Windows.Forms.TextBox
    Private WithEvents textBoxSubject As System.Windows.Forms.TextBox
    Private WithEvents textBoxReciever As System.Windows.Forms.TextBox
    Private WithEvents label3 As System.Windows.Forms.Label
    Private WithEvents label2 As System.Windows.Forms.Label
    Private WithEvents label1 As System.Windows.Forms.Label
    Private WithEvents textBox1 As System.Windows.Forms.TextBox
    Private WithEvents button1 As System.Windows.Forms.Button

End Class
