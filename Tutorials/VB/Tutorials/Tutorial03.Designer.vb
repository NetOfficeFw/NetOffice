<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Tutorial03
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
        Me.buttonAddRemoveWorkbook = New System.Windows.Forms.Button()
        Me.buttonAddins = New System.Windows.Forms.Button()
        Me.buttonWorkbook = New System.Windows.Forms.Button()
        Me.buttonExcel = New System.Windows.Forms.Button()
        Me.labelProxyCount = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'buttonAddRemoveWorkbook
        '
        Me.buttonAddRemoveWorkbook.Enabled = False
        Me.buttonAddRemoveWorkbook.Location = New System.Drawing.Point(369, 25)
        Me.buttonAddRemoveWorkbook.Name = "buttonAddRemoveWorkbook"
        Me.buttonAddRemoveWorkbook.Size = New System.Drawing.Size(176, 25)
        Me.buttonAddRemoveWorkbook.TabIndex = 26
        Me.buttonAddRemoveWorkbook.Text = "Add && Remove Workbook"
        Me.buttonAddRemoveWorkbook.UseVisualStyleBackColor = True
        '
        'buttonAddins
        '
        Me.buttonAddins.Enabled = False
        Me.buttonAddins.Location = New System.Drawing.Point(260, 25)
        Me.buttonAddins.Name = "buttonAddins"
        Me.buttonAddins.Size = New System.Drawing.Size(103, 25)
        Me.buttonAddins.TabIndex = 25
        Me.buttonAddins.Text = "Enum Addins"
        Me.buttonAddins.UseVisualStyleBackColor = True
        '
        'buttonWorkbook
        '
        Me.buttonWorkbook.Enabled = False
        Me.buttonWorkbook.Location = New System.Drawing.Point(151, 25)
        Me.buttonWorkbook.Name = "buttonWorkbook"
        Me.buttonWorkbook.Size = New System.Drawing.Size(103, 25)
        Me.buttonWorkbook.TabIndex = 24
        Me.buttonWorkbook.Text = "Add Workbook"
        Me.buttonWorkbook.UseVisualStyleBackColor = True
        '
        'buttonExcel
        '
        Me.buttonExcel.Location = New System.Drawing.Point(32, 25)
        Me.buttonExcel.Name = "buttonExcel"
        Me.buttonExcel.Size = New System.Drawing.Size(103, 25)
        Me.buttonExcel.TabIndex = 23
        Me.buttonExcel.Text = "Start Excel"
        Me.buttonExcel.UseVisualStyleBackColor = True
        '
        'labelProxyCount
        '
        Me.labelProxyCount.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.labelProxyCount.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelProxyCount.Location = New System.Drawing.Point(237, 90)
        Me.labelProxyCount.Name = "labelProxyCount"
        Me.labelProxyCount.Size = New System.Drawing.Size(47, 20)
        Me.labelProxyCount.TabIndex = 22
        Me.labelProxyCount.Text = "0"
        Me.labelProxyCount.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label1.Location = New System.Drawing.Point(34, 90)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(197, 20)
        Me.label1.TabIndex = 21
        Me.label1.Text = "Current COM Proxies open"
        '
        'Tutorial03
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.buttonAddRemoveWorkbook)
        Me.Controls.Add(Me.buttonAddins)
        Me.Controls.Add(Me.buttonWorkbook)
        Me.Controls.Add(Me.buttonExcel)
        Me.Controls.Add(Me.labelProxyCount)
        Me.Controls.Add(Me.label1)
        Me.Name = "Tutorial03"
        Me.Size = New System.Drawing.Size(686, 478)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents buttonAddRemoveWorkbook As System.Windows.Forms.Button
    Private WithEvents buttonAddins As System.Windows.Forms.Button
    Private WithEvents buttonWorkbook As System.Windows.Forms.Button
    Private WithEvents buttonExcel As System.Windows.Forms.Button
    Private WithEvents labelProxyCount As System.Windows.Forms.Label
    Private WithEvents label1 As System.Windows.Forms.Label

End Class
