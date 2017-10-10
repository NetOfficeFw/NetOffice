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
        Me.buttonDisposeChildInstances = New System.Windows.Forms.Button()
        Me.buttonAddins = New System.Windows.Forms.Button()
        Me.buttonWorkbook = New System.Windows.Forms.Button()
        Me.buttonExcel = New System.Windows.Forms.Button()
        Me.labelProxyCount = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.label2 = New System.Windows.Forms.Label()
        Me.instanceMonitor1 = New NetOffice.Contribution.Controls.InstanceMonitor()
        Me.SuspendLayout()
        '
        'buttonDisposeChildInstances
        '
        Me.buttonDisposeChildInstances.Enabled = False
        Me.buttonDisposeChildInstances.Location = New System.Drawing.Point(476, 31)
        Me.buttonDisposeChildInstances.Margin = New System.Windows.Forms.Padding(4)
        Me.buttonDisposeChildInstances.Name = "buttonDisposeChildInstances"
        Me.buttonDisposeChildInstances.Size = New System.Drawing.Size(263, 31)
        Me.buttonDisposeChildInstances.TabIndex = 26
        Me.buttonDisposeChildInstances.Text = "Dispose Application Child Instances"
        Me.buttonDisposeChildInstances.UseVisualStyleBackColor = True
        '
        'buttonAddins
        '
        Me.buttonAddins.Enabled = False
        Me.buttonAddins.Location = New System.Drawing.Point(331, 31)
        Me.buttonAddins.Margin = New System.Windows.Forms.Padding(4)
        Me.buttonAddins.Name = "buttonAddins"
        Me.buttonAddins.Size = New System.Drawing.Size(137, 31)
        Me.buttonAddins.TabIndex = 25
        Me.buttonAddins.Text = "Enum Addins"
        Me.buttonAddins.UseVisualStyleBackColor = True
        '
        'buttonWorkbook
        '
        Me.buttonWorkbook.Enabled = False
        Me.buttonWorkbook.Location = New System.Drawing.Point(187, 31)
        Me.buttonWorkbook.Margin = New System.Windows.Forms.Padding(4)
        Me.buttonWorkbook.Name = "buttonWorkbook"
        Me.buttonWorkbook.Size = New System.Drawing.Size(137, 31)
        Me.buttonWorkbook.TabIndex = 24
        Me.buttonWorkbook.Text = "Add Workbook"
        Me.buttonWorkbook.UseVisualStyleBackColor = True
        '
        'buttonExcel
        '
        Me.buttonExcel.Location = New System.Drawing.Point(43, 31)
        Me.buttonExcel.Margin = New System.Windows.Forms.Padding(4)
        Me.buttonExcel.Name = "buttonExcel"
        Me.buttonExcel.Size = New System.Drawing.Size(137, 31)
        Me.buttonExcel.TabIndex = 23
        Me.buttonExcel.Text = "Start Excel"
        Me.buttonExcel.UseVisualStyleBackColor = True
        '
        'labelProxyCount
        '
        Me.labelProxyCount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.labelProxyCount.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelProxyCount.Location = New System.Drawing.Point(247, 85)
        Me.labelProxyCount.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.labelProxyCount.Name = "labelProxyCount"
        Me.labelProxyCount.Size = New System.Drawing.Size(63, 25)
        Me.labelProxyCount.TabIndex = 22
        Me.labelProxyCount.Text = "0"
        Me.labelProxyCount.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label1.Location = New System.Drawing.Point(45, 87)
        Me.label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(197, 20)
        Me.label1.TabIndex = 21
        Me.label1.Text = "Current COM Proxies open"
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.label2.Location = New System.Drawing.Point(45, 143)
        Me.label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(83, 20)
        Me.label2.TabIndex = 27
        Me.label2.Text = "Proxy Tree"
        '
        'instanceMonitor1
        '
        Me.instanceMonitor1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.instanceMonitor1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.instanceMonitor1.Location = New System.Drawing.Point(43, 172)
        Me.instanceMonitor1.Margin = New System.Windows.Forms.Padding(5)
        Me.instanceMonitor1.Name = "instanceMonitor1"
        Me.instanceMonitor1.Size = New System.Drawing.Size(696, 378)
        Me.instanceMonitor1.TabIndex = 28
        '
        'Tutorial03
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(227, Byte), Integer), CType(CType(243, Byte), Integer))
        Me.Controls.Add(Me.instanceMonitor1)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.buttonDisposeChildInstances)
        Me.Controls.Add(Me.buttonAddins)
        Me.Controls.Add(Me.buttonWorkbook)
        Me.Controls.Add(Me.buttonExcel)
        Me.Controls.Add(Me.labelProxyCount)
        Me.Controls.Add(Me.label1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Tutorial03"
        Me.Size = New System.Drawing.Size(800, 600)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents buttonDisposeChildInstances As System.Windows.Forms.Button
    Private WithEvents buttonAddins As System.Windows.Forms.Button
    Private WithEvents buttonWorkbook As System.Windows.Forms.Button
    Private WithEvents buttonExcel As System.Windows.Forms.Button
    Private WithEvents labelProxyCount As System.Windows.Forms.Label
    Private WithEvents label1 As System.Windows.Forms.Label
    Private WithEvents label2 As Label
    Private WithEvents instanceMonitor1 As NetOffice.Contribution.Controls.InstanceMonitor
End Class
