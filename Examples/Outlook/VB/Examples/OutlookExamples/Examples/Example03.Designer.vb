<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Example03
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Example03))
        Me.textBoxBody = New System.Windows.Forms.TextBox()
        Me.textBoxSubject = New System.Windows.Forms.TextBox()
        Me.textBoxReciever = New System.Windows.Forms.TextBox()
        Me.label3 = New System.Windows.Forms.Label()
        Me.label2 = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.textBox1 = New System.Windows.Forms.TextBox()
        Me.buttonStartExample = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'textBoxBody
        '
        Me.textBoxBody.Location = New System.Drawing.Point(91, 162)
        Me.textBoxBody.Multiline = True
        Me.textBoxBody.Name = "textBoxBody"
        Me.textBoxBody.Size = New System.Drawing.Size(610, 114)
        Me.textBoxBody.TabIndex = 27
        Me.textBoxBody.Text = "This is a mail from NetOffice example."
        Me.textBoxBody.WordWrap = False
        '
        'textBoxSubject
        '
        Me.textBoxSubject.Location = New System.Drawing.Point(91, 136)
        Me.textBoxSubject.Name = "textBoxSubject"
        Me.textBoxSubject.Size = New System.Drawing.Size(610, 20)
        Me.textBoxSubject.TabIndex = 26
        Me.textBoxSubject.Text = "NetOffice Example Mail"
        '
        'textBoxReciever
        '
        Me.textBoxReciever.Location = New System.Drawing.Point(91, 110)
        Me.textBoxReciever.Name = "textBoxReciever"
        Me.textBoxReciever.Size = New System.Drawing.Size(610, 20)
        Me.textBoxReciever.TabIndex = 25
        Me.textBoxReciever.Text = "public.sebastian@web.de"
        '
        'label3
        '
        Me.label3.AutoSize = True
        Me.label3.Location = New System.Drawing.Point(36, 165)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(31, 13)
        Me.label3.TabIndex = 24
        Me.label3.Text = "Body"
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Location = New System.Drawing.Point(35, 140)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(43, 13)
        Me.label2.TabIndex = 23
        Me.label2.Text = "Subject"
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(35, 113)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(50, 13)
        Me.label1.TabIndex = 22
        Me.label1.Text = "Reciever"
        '
        'textBox1
        '
        Me.textBox1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.textBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.textBox1.Location = New System.Drawing.Point(91, 68)
        Me.textBox1.Multiline = True
        Me.textBox1.Name = "textBox1"
        Me.textBox1.Size = New System.Drawing.Size(610, 25)
        Me.textBox1.TabIndex = 21
        Me.textBox1.Text = "this example shows you how to send a mail."
        '
        'buttonStartExample
        '
        Me.buttonStartExample.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.buttonStartExample.Image = CType(resources.GetObject("buttonStartExample.Image"), System.Drawing.Image)
        Me.buttonStartExample.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.buttonStartExample.Location = New System.Drawing.Point(36, 22)
        Me.buttonStartExample.Name = "buttonStartExample"
        Me.buttonStartExample.Size = New System.Drawing.Size(665, 30)
        Me.buttonStartExample.TabIndex = 20
        Me.buttonStartExample.Text = "Start example"
        Me.buttonStartExample.UseVisualStyleBackColor = True
        '
        'Example03
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.textBoxBody)
        Me.Controls.Add(Me.textBoxSubject)
        Me.Controls.Add(Me.textBoxReciever)
        Me.Controls.Add(Me.label3)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.label1)
        Me.Controls.Add(Me.textBox1)
        Me.Controls.Add(Me.buttonStartExample)
        Me.Name = "Example03"
        Me.Size = New System.Drawing.Size(739, 304)
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
    Private WithEvents buttonStartExample As System.Windows.Forms.Button

End Class
