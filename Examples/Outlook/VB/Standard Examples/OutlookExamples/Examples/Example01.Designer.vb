<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Example01
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Example01))
        Me.labelItemsCount = New System.Windows.Forms.Label()
        Me.listViewInboxFolder = New System.Windows.Forms.ListView()
        Me.columnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.columnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.textBox1 = New System.Windows.Forms.TextBox()
        Me.buttonStartExample = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'labelItemsCount
        '
        Me.labelItemsCount.AutoSize = True
        Me.labelItemsCount.Location = New System.Drawing.Point(39, 97)
        Me.labelItemsCount.Name = "labelItemsCount"
        Me.labelItemsCount.Size = New System.Drawing.Size(10, 13)
        Me.labelItemsCount.TabIndex = 11
        Me.labelItemsCount.Text = "."
        '
        'listViewInboxFolder
        '
        Me.listViewInboxFolder.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.columnHeader1, Me.columnHeader2})
        Me.listViewInboxFolder.Location = New System.Drawing.Point(36, 114)
        Me.listViewInboxFolder.MultiSelect = False
        Me.listViewInboxFolder.Name = "listViewInboxFolder"
        Me.listViewInboxFolder.Size = New System.Drawing.Size(665, 173)
        Me.listViewInboxFolder.TabIndex = 10
        Me.listViewInboxFolder.UseCompatibleStateImageBehavior = False
        Me.listViewInboxFolder.View = System.Windows.Forms.View.Details
        '
        'columnHeader1
        '
        Me.columnHeader1.Text = "From"
        Me.columnHeader1.Width = 130
        '
        'columnHeader2
        '
        Me.columnHeader2.Text = "Subject"
        Me.columnHeader2.Width = 300
        '
        'textBox1
        '
        Me.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.textBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBox1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.textBox1.Location = New System.Drawing.Point(36, 68)
        Me.textBox1.Multiline = True
        Me.textBox1.Name = "textBox1"
        Me.textBox1.Size = New System.Drawing.Size(665, 24)
        Me.textBox1.TabIndex = 9
        Me.textBox1.Text = "this example shows you how to enumerate your inbox folder."
        '
        'buttonStartExample
        '
        Me.buttonStartExample.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.buttonStartExample.Image = CType(resources.GetObject("buttonStartExample.Image"), System.Drawing.Image)
        Me.buttonStartExample.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.buttonStartExample.Location = New System.Drawing.Point(36, 22)
        Me.buttonStartExample.Name = "buttonStartExample"
        Me.buttonStartExample.Size = New System.Drawing.Size(665, 30)
        Me.buttonStartExample.TabIndex = 8
        Me.buttonStartExample.Text = "Start example"
        Me.buttonStartExample.UseVisualStyleBackColor = True
        '
        'Example01
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.labelItemsCount)
        Me.Controls.Add(Me.listViewInboxFolder)
        Me.Controls.Add(Me.textBox1)
        Me.Controls.Add(Me.buttonStartExample)
        Me.Name = "Example01"
        Me.Size = New System.Drawing.Size(739, 304)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents labelItemsCount As System.Windows.Forms.Label
    Private WithEvents listViewInboxFolder As System.Windows.Forms.ListView
    Private WithEvents columnHeader1 As System.Windows.Forms.ColumnHeader
    Private WithEvents columnHeader2 As System.Windows.Forms.ColumnHeader
    Private WithEvents textBox1 As System.Windows.Forms.TextBox
    Private WithEvents buttonStartExample As System.Windows.Forms.Button

End Class
