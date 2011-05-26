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
        Me.treeViewInfo = New System.Windows.Forms.TreeView
        Me.buttonSelectDatabase = New System.Windows.Forms.Button
        Me.textBoxFilePath = New System.Windows.Forms.TextBox
        Me.label2 = New System.Windows.Forms.Label
        Me.label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'treeViewInfo
        '
        Me.treeViewInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.treeViewInfo.BackColor = System.Drawing.Color.DarkKhaki
        Me.treeViewInfo.Location = New System.Drawing.Point(18, 83)
        Me.treeViewInfo.Name = "treeViewInfo"
        Me.treeViewInfo.Size = New System.Drawing.Size(297, 351)
        Me.treeViewInfo.TabIndex = 9
        '
        'buttonSelectDatabase
        '
        Me.buttonSelectDatabase.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.buttonSelectDatabase.Location = New System.Drawing.Point(275, 40)
        Me.buttonSelectDatabase.Name = "buttonSelectDatabase"
        Me.buttonSelectDatabase.Size = New System.Drawing.Size(40, 22)
        Me.buttonSelectDatabase.TabIndex = 8
        Me.buttonSelectDatabase.Text = "..."
        Me.buttonSelectDatabase.UseVisualStyleBackColor = True
        '
        'textBoxFilePath
        '
        Me.textBoxFilePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.textBoxFilePath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.textBoxFilePath.Location = New System.Drawing.Point(56, 42)
        Me.textBoxFilePath.Name = "textBoxFilePath"
        Me.textBoxFilePath.ReadOnly = True
        Me.textBoxFilePath.Size = New System.Drawing.Size(213, 20)
        Me.textBoxFilePath.TabIndex = 7
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.BackColor = System.Drawing.Color.Khaki
        Me.label2.Location = New System.Drawing.Point(21, 10)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(200, 13)
        Me.label2.TabIndex = 6
        Me.label2.Text = "Select a database und see details about."
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Location = New System.Drawing.Point(21, 45)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(29, 13)
        Me.label1.TabIndex = 5
        Me.label1.Text = "Path"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(333, 445)
        Me.Controls.Add(Me.treeViewInfo)
        Me.Controls.Add(Me.buttonSelectDatabase)
        Me.Controls.Add(Me.textBoxFilePath)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.label1)
        Me.Name = "Form1"
        Me.Text = "Example04 - Database informations"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents treeViewInfo As System.Windows.Forms.TreeView
    Private WithEvents buttonSelectDatabase As System.Windows.Forms.Button
    Private WithEvents textBoxFilePath As System.Windows.Forms.TextBox
    Private WithEvents label2 As System.Windows.Forms.Label
    Private WithEvents label1 As System.Windows.Forms.Label

End Class
