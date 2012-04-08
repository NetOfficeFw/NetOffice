<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Example04
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
        Me.treeViewInfo = New System.Windows.Forms.TreeView()
        Me.buttonSelectDatabase = New System.Windows.Forms.Button()
        Me.textBoxFilePath = New System.Windows.Forms.TextBox()
        Me.label2 = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'treeViewInfo
        '
        Me.treeViewInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.treeViewInfo.BackColor = System.Drawing.Color.DarkKhaki
        Me.treeViewInfo.Location = New System.Drawing.Point(67, 54)
        Me.treeViewInfo.Name = "treeViewInfo"
        Me.treeViewInfo.Size = New System.Drawing.Size(591, 207)
        Me.treeViewInfo.TabIndex = 14
        '
        'buttonSelectDatabase
        '
        Me.buttonSelectDatabase.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.buttonSelectDatabase.Location = New System.Drawing.Point(666, 16)
        Me.buttonSelectDatabase.Name = "buttonSelectDatabase"
        Me.buttonSelectDatabase.Size = New System.Drawing.Size(40, 21)
        Me.buttonSelectDatabase.TabIndex = 13
        Me.buttonSelectDatabase.Text = "..."
        Me.buttonSelectDatabase.UseVisualStyleBackColor = True
        '
        'textBoxFilePath
        '
        Me.textBoxFilePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.textBoxFilePath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.textBoxFilePath.Location = New System.Drawing.Point(67, 17)
        Me.textBoxFilePath.Name = "textBoxFilePath"
        Me.textBoxFilePath.ReadOnly = True
        Me.textBoxFilePath.Size = New System.Drawing.Size(591, 20)
        Me.textBoxFilePath.TabIndex = 12
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.BackColor = System.Drawing.Color.Khaki
        Me.label2.Location = New System.Drawing.Point(66, 276)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(200, 13)
        Me.label2.TabIndex = 11
        Me.label2.Text = "Select a database und see details about."
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(32, 20)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(29, 13)
        Me.label1.TabIndex = 10
        Me.label1.Text = "Path"
        '
        'Example04
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.treeViewInfo)
        Me.Controls.Add(Me.buttonSelectDatabase)
        Me.Controls.Add(Me.textBoxFilePath)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.label1)
        Me.Name = "Example04"
        Me.Size = New System.Drawing.Size(739, 304)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents treeViewInfo As System.Windows.Forms.TreeView
    Private WithEvents buttonSelectDatabase As System.Windows.Forms.Button
    Private WithEvents textBoxFilePath As System.Windows.Forms.TextBox
    Private WithEvents label2 As System.Windows.Forms.Label
    Private WithEvents label1 As System.Windows.Forms.Label

End Class
