<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SampleControl
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SampleControl))
        Me.pictureBox3 = New System.Windows.Forms.PictureBox()
        Me.label5 = New System.Windows.Forms.Label()
        Me.label4 = New System.Windows.Forms.Label()
        Me.propertyGridDetails = New System.Windows.Forms.PropertyGrid()
        Me.pictureBox2 = New System.Windows.Forms.PictureBox()
        Me.label3 = New System.Windows.Forms.Label()
        Me.pictureBox1 = New System.Windows.Forms.PictureBox()
        Me.label2 = New System.Windows.Forms.Label()
        Me.listViewSearchResults = New System.Windows.Forms.ListView()
        Me.columnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.columnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.columnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.imageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.textBoxSearch = New System.Windows.Forms.TextBox()
        Me.pictureBox4 = New System.Windows.Forms.PictureBox()
        Me.label1 = New System.Windows.Forms.Label()
        CType(Me.pictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pictureBox3
        '
        Me.pictureBox3.BackColor = System.Drawing.Color.Transparent
        Me.pictureBox3.Image = CType(resources.GetObject("pictureBox3.Image"), System.Drawing.Image)
        Me.pictureBox3.Location = New System.Drawing.Point(23, 17)
        Me.pictureBox3.Name = "pictureBox3"
        Me.pictureBox3.Size = New System.Drawing.Size(16, 16)
        Me.pictureBox3.TabIndex = 102
        Me.pictureBox3.TabStop = False
        '
        'label5
        '
        Me.label5.AutoSize = True
        Me.label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label5.Location = New System.Drawing.Point(42, 17)
        Me.label5.Name = "label5"
        Me.label5.Size = New System.Drawing.Size(140, 13)
        Me.label5.TabIndex = 101
        Me.label5.Text = "Customer Sample Panel"
        '
        'label4
        '
        Me.label4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.label4.Location = New System.Drawing.Point(19, 42)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(254, 47)
        Me.label4.TabIndex = 100
        Me.label4.Text = "insert a name and click in the result list to see details. double click to copy a" & _
            " result item in your current selected cell."
        '
        'propertyGridDetails
        '
        Me.propertyGridDetails.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.propertyGridDetails.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.propertyGridDetails.CommandsVisibleIfAvailable = False
        Me.propertyGridDetails.HelpVisible = False
        Me.propertyGridDetails.Location = New System.Drawing.Point(22, 422)
        Me.propertyGridDetails.Name = "propertyGridDetails"
        Me.propertyGridDetails.Size = New System.Drawing.Size(251, 163)
        Me.propertyGridDetails.TabIndex = 99
        Me.propertyGridDetails.ToolbarVisible = False
        '
        'pictureBox2
        '
        Me.pictureBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.pictureBox2.BackColor = System.Drawing.Color.Transparent
        Me.pictureBox2.Image = CType(resources.GetObject("pictureBox2.Image"), System.Drawing.Image)
        Me.pictureBox2.Location = New System.Drawing.Point(22, 403)
        Me.pictureBox2.Name = "pictureBox2"
        Me.pictureBox2.Size = New System.Drawing.Size(16, 16)
        Me.pictureBox2.TabIndex = 98
        Me.pictureBox2.TabStop = False
        '
        'label3
        '
        Me.label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.label3.Location = New System.Drawing.Point(42, 404)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(142, 15)
        Me.label3.TabIndex = 97
        Me.label3.Text = "Details:"
        '
        'pictureBox1
        '
        Me.pictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.pictureBox1.Image = CType(resources.GetObject("pictureBox1.Image"), System.Drawing.Image)
        Me.pictureBox1.Location = New System.Drawing.Point(22, 149)
        Me.pictureBox1.Name = "pictureBox1"
        Me.pictureBox1.Size = New System.Drawing.Size(16, 16)
        Me.pictureBox1.TabIndex = 96
        Me.pictureBox1.TabStop = False
        '
        'label2
        '
        Me.label2.Location = New System.Drawing.Point(42, 150)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(81, 15)
        Me.label2.TabIndex = 95
        Me.label2.Text = "Customers:"
        '
        'listViewSearchResults
        '
        Me.listViewSearchResults.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.listViewSearchResults.CausesValidation = False
        Me.listViewSearchResults.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.columnHeader3, Me.columnHeader1, Me.columnHeader2})
        Me.listViewSearchResults.FullRowSelect = True
        Me.listViewSearchResults.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.listViewSearchResults.Location = New System.Drawing.Point(22, 169)
        Me.listViewSearchResults.MultiSelect = False
        Me.listViewSearchResults.Name = "listViewSearchResults"
        Me.listViewSearchResults.ShowGroups = False
        Me.listViewSearchResults.Size = New System.Drawing.Size(251, 218)
        Me.listViewSearchResults.SmallImageList = Me.imageList1
        Me.listViewSearchResults.TabIndex = 94
        Me.listViewSearchResults.UseCompatibleStateImageBehavior = False
        Me.listViewSearchResults.View = System.Windows.Forms.View.Details
        '
        'columnHeader3
        '
        Me.columnHeader3.Text = ""
        Me.columnHeader3.Width = 20
        '
        'columnHeader1
        '
        Me.columnHeader1.Text = "ID"
        Me.columnHeader1.Width = 25
        '
        'columnHeader2
        '
        Me.columnHeader2.Text = "Name"
        Me.columnHeader2.Width = 161
        '
        'imageList1
        '
        Me.imageList1.ImageStream = CType(resources.GetObject("imageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.imageList1.Images.SetKeyName(0, "user.png")
        '
        'textBoxSearch
        '
        Me.textBoxSearch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.textBoxSearch.Location = New System.Drawing.Point(22, 113)
        Me.textBoxSearch.Name = "textBoxSearch"
        Me.textBoxSearch.Size = New System.Drawing.Size(251, 20)
        Me.textBoxSearch.TabIndex = 93
        '
        'pictureBox4
        '
        Me.pictureBox4.BackColor = System.Drawing.Color.Transparent
        Me.pictureBox4.Image = CType(resources.GetObject("pictureBox4.Image"), System.Drawing.Image)
        Me.pictureBox4.Location = New System.Drawing.Point(22, 96)
        Me.pictureBox4.Name = "pictureBox4"
        Me.pictureBox4.Size = New System.Drawing.Size(16, 16)
        Me.pictureBox4.TabIndex = 92
        Me.pictureBox4.TabStop = False
        '
        'label1
        '
        Me.label1.Location = New System.Drawing.Point(42, 97)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(142, 15)
        Me.label1.TabIndex = 91
        Me.label1.Text = "Name:"
        '
        'SampleControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.pictureBox3)
        Me.Controls.Add(Me.label5)
        Me.Controls.Add(Me.label4)
        Me.Controls.Add(Me.propertyGridDetails)
        Me.Controls.Add(Me.pictureBox2)
        Me.Controls.Add(Me.label3)
        Me.Controls.Add(Me.pictureBox1)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.listViewSearchResults)
        Me.Controls.Add(Me.textBoxSearch)
        Me.Controls.Add(Me.pictureBox4)
        Me.Controls.Add(Me.label1)
        Me.Name = "SampleControl"
        Me.Size = New System.Drawing.Size(300, 607)
        CType(Me.pictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents pictureBox3 As System.Windows.Forms.PictureBox
    Private WithEvents label5 As System.Windows.Forms.Label
    Private WithEvents label4 As System.Windows.Forms.Label
    Private WithEvents propertyGridDetails As System.Windows.Forms.PropertyGrid
    Private WithEvents pictureBox2 As System.Windows.Forms.PictureBox
    Private WithEvents label3 As System.Windows.Forms.Label
    Private WithEvents pictureBox1 As System.Windows.Forms.PictureBox
    Private WithEvents label2 As System.Windows.Forms.Label
    Private WithEvents listViewSearchResults As System.Windows.Forms.ListView
    Private WithEvents columnHeader3 As System.Windows.Forms.ColumnHeader
    Private WithEvents columnHeader1 As System.Windows.Forms.ColumnHeader
    Private WithEvents columnHeader2 As System.Windows.Forms.ColumnHeader
    Private WithEvents textBoxSearch As System.Windows.Forms.TextBox
    Private WithEvents pictureBox4 As System.Windows.Forms.PictureBox
    Private WithEvents label1 As System.Windows.Forms.Label
    Private WithEvents imageList1 As System.Windows.Forms.ImageList

End Class
