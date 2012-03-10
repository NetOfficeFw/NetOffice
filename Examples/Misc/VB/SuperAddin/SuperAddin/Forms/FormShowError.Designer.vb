<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormShowError
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormShowError))
        Me.pictureBoxError = New System.Windows.Forms.PictureBox
        Me.labelErrorFooter = New System.Windows.Forms.Label
        Me.buttonDetails = New System.Windows.Forms.Button
        Me.buttonOk = New System.Windows.Forms.Button
        Me.listViewExceptions = New System.Windows.Forms.ListView
        Me.columnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.columnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.columnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.columnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.labelErrorHeader = New System.Windows.Forms.Label
        CType(Me.pictureBoxError, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pictureBoxError
        '
        Me.pictureBoxError.Image = CType(resources.GetObject("pictureBoxError.Image"), System.Drawing.Image)
        Me.pictureBoxError.Location = New System.Drawing.Point(38, 26)
        Me.pictureBoxError.Name = "pictureBoxError"
        Me.pictureBoxError.Size = New System.Drawing.Size(51, 47)
        Me.pictureBoxError.TabIndex = 17
        Me.pictureBoxError.TabStop = False
        '
        'labelErrorFooter
        '
        Me.labelErrorFooter.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.labelErrorFooter.BackColor = System.Drawing.SystemColors.Control
        Me.labelErrorFooter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelErrorFooter.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.labelErrorFooter.Location = New System.Drawing.Point(97, 79)
        Me.labelErrorFooter.Name = "labelErrorFooter"
        Me.labelErrorFooter.Size = New System.Drawing.Size(393, 43)
        Me.labelErrorFooter.TabIndex = 16
        Me.labelErrorFooter.Text = "labelErrorFooter"
        Me.labelErrorFooter.Visible = False
        '
        'buttonDetails
        '
        Me.buttonDetails.Location = New System.Drawing.Point(28, 125)
        Me.buttonDetails.Name = "buttonDetails"
        Me.buttonDetails.Size = New System.Drawing.Size(87, 22)
        Me.buttonDetails.TabIndex = 15
        Me.buttonDetails.Text = "<< Details"
        Me.buttonDetails.UseVisualStyleBackColor = True
        '
        'buttonOk
        '
        Me.buttonOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.buttonOk.Location = New System.Drawing.Point(403, 125)
        Me.buttonOk.Name = "buttonOk"
        Me.buttonOk.Size = New System.Drawing.Size(87, 22)
        Me.buttonOk.TabIndex = 14
        Me.buttonOk.Text = "Ok"
        Me.buttonOk.UseVisualStyleBackColor = True
        '
        'listViewExceptions
        '
        Me.listViewExceptions.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.listViewExceptions.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.columnHeader1, Me.columnHeader2, Me.columnHeader3, Me.columnHeader4})
        Me.listViewExceptions.FullRowSelect = True
        Me.listViewExceptions.GridLines = True
        Me.listViewExceptions.HideSelection = False
        Me.listViewExceptions.Location = New System.Drawing.Point(25, 173)
        Me.listViewExceptions.Name = "listViewExceptions"
        Me.listViewExceptions.Size = New System.Drawing.Size(483, 177)
        Me.listViewExceptions.TabIndex = 13
        Me.listViewExceptions.UseCompatibleStateImageBehavior = False
        Me.listViewExceptions.View = System.Windows.Forms.View.Details
        '
        'columnHeader1
        '
        Me.columnHeader1.Text = "Nr"
        Me.columnHeader1.Width = 38
        '
        'columnHeader2
        '
        Me.columnHeader2.Text = "Modul"
        Me.columnHeader2.Width = 91
        '
        'columnHeader3
        '
        Me.columnHeader3.Text = "Type"
        Me.columnHeader3.Width = 83
        '
        'columnHeader4
        '
        Me.columnHeader4.Text = "Text"
        Me.columnHeader4.Width = 218
        '
        'labelErrorHeader
        '
        Me.labelErrorHeader.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.labelErrorHeader.BackColor = System.Drawing.SystemColors.Control
        Me.labelErrorHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelErrorHeader.ForeColor = System.Drawing.SystemColors.ControlText
        Me.labelErrorHeader.Location = New System.Drawing.Point(97, 25)
        Me.labelErrorHeader.Name = "labelErrorHeader"
        Me.labelErrorHeader.Size = New System.Drawing.Size(393, 39)
        Me.labelErrorHeader.TabIndex = 12
        Me.labelErrorHeader.Text = "labelErrorHeader"
        Me.labelErrorHeader.Visible = False
        '
        'FormShowError
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(533, 375)
        Me.Controls.Add(Me.pictureBoxError)
        Me.Controls.Add(Me.labelErrorFooter)
        Me.Controls.Add(Me.buttonDetails)
        Me.Controls.Add(Me.buttonOk)
        Me.Controls.Add(Me.listViewExceptions)
        Me.Controls.Add(Me.labelErrorHeader)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormShowError"
        Me.Padding = New System.Windows.Forms.Padding(9)
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Error"
        CType(Me.pictureBoxError, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents pictureBoxError As System.Windows.Forms.PictureBox
    Private WithEvents labelErrorFooter As System.Windows.Forms.Label
    Private WithEvents buttonDetails As System.Windows.Forms.Button
    Private WithEvents buttonOk As System.Windows.Forms.Button
    Private WithEvents listViewExceptions As System.Windows.Forms.ListView
    Private WithEvents columnHeader1 As System.Windows.Forms.ColumnHeader
    Private WithEvents columnHeader2 As System.Windows.Forms.ColumnHeader
    Private WithEvents columnHeader3 As System.Windows.Forms.ColumnHeader
    Private WithEvents columnHeader4 As System.Windows.Forms.ColumnHeader
    Private WithEvents labelErrorHeader As System.Windows.Forms.Label

End Class
