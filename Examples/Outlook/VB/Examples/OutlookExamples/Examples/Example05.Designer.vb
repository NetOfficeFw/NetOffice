<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Example05
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Example05))
        Me.listViewContacts = New System.Windows.Forms.ListView()
        Me.columnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.columnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.textBox1 = New System.Windows.Forms.TextBox()
        Me.buttonStartExample = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'listViewContacts
        '
        Me.listViewContacts.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.columnHeader1, Me.columnHeader2})
        Me.listViewContacts.Location = New System.Drawing.Point(35, 111)
        Me.listViewContacts.Name = "listViewContacts"
        Me.listViewContacts.Size = New System.Drawing.Size(665, 168)
        Me.listViewContacts.TabIndex = 16
        Me.listViewContacts.UseCompatibleStateImageBehavior = False
        Me.listViewContacts.View = System.Windows.Forms.View.Details
        '
        'columnHeader1
        '
        Me.columnHeader1.Text = "Nr."
        Me.columnHeader1.Width = 40
        '
        'columnHeader2
        '
        Me.columnHeader2.Text = "CompanyAndFullName"
        Me.columnHeader2.Width = 300
        '
        'textBox1
        '
        Me.textBox1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.textBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.textBox1.Location = New System.Drawing.Point(35, 69)
        Me.textBox1.Multiline = True
        Me.textBox1.Name = "textBox1"
        Me.textBox1.Size = New System.Drawing.Size(665, 24)
        Me.textBox1.TabIndex = 15
        Me.textBox1.Text = "this example shows you how to enumerate contacts."
        '
        'buttonStartExample
        '
        Me.buttonStartExample.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.buttonStartExample.Image = CType(resources.GetObject("buttonStartExample.Image"), System.Drawing.Image)
        Me.buttonStartExample.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.buttonStartExample.Location = New System.Drawing.Point(35, 22)
        Me.buttonStartExample.Name = "buttonStartExample"
        Me.buttonStartExample.Size = New System.Drawing.Size(665, 30)
        Me.buttonStartExample.TabIndex = 14
        Me.buttonStartExample.Text = "Start example"
        Me.buttonStartExample.UseVisualStyleBackColor = True
        '
        'Example05
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.listViewContacts)
        Me.Controls.Add(Me.textBox1)
        Me.Controls.Add(Me.buttonStartExample)
        Me.Name = "Example05"
        Me.Size = New System.Drawing.Size(739, 304)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents listViewContacts As System.Windows.Forms.ListView
    Private WithEvents columnHeader1 As System.Windows.Forms.ColumnHeader
    Private WithEvents columnHeader2 As System.Windows.Forms.ColumnHeader
    Private WithEvents textBox1 As System.Windows.Forms.TextBox
    Private WithEvents buttonStartExample As System.Windows.Forms.Button

End Class
