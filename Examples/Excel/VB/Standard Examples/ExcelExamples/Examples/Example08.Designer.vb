<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Example08
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Example08))
        Me.labelEventLogHeader = New System.Windows.Forms.Label()
        Me.textBoxEvents = New System.Windows.Forms.TextBox()
        Me.textBoxDescription = New System.Windows.Forms.TextBox()
        Me.buttonStartExample = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'labelEventLogHeader
        '
        Me.labelEventLogHeader.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.labelEventLogHeader.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.labelEventLogHeader.Location = New System.Drawing.Point(27, 134)
        Me.labelEventLogHeader.Name = "labelEventLogHeader"
        Me.labelEventLogHeader.Size = New System.Drawing.Size(682, 22)
        Me.labelEventLogHeader.TabIndex = 21
        Me.labelEventLogHeader.Text = "EventLog"
        '
        'textBoxEvents
        '
        Me.textBoxEvents.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.textBoxEvents.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.textBoxEvents.Location = New System.Drawing.Point(27, 159)
        Me.textBoxEvents.Multiline = True
        Me.textBoxEvents.Name = "textBoxEvents"
        Me.textBoxEvents.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.textBoxEvents.Size = New System.Drawing.Size(682, 121)
        Me.textBoxEvents.TabIndex = 20
        '
        'textBoxDescription
        '
        Me.textBoxDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.textBoxDescription.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.textBoxDescription.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textBoxDescription.Location = New System.Drawing.Point(27, 72)
        Me.textBoxDescription.Multiline = True
        Me.textBoxDescription.Name = "textBoxDescription"
        Me.textBoxDescription.Size = New System.Drawing.Size(683, 49)
        Me.textBoxDescription.TabIndex = 19
        Me.textBoxDescription.Text = "This example contains code to catch events in excel." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "example 9 shows you how to " &
    "catch click events from excel gui." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'buttonStartExample
        '
        Me.buttonStartExample.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.buttonStartExample.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.buttonStartExample.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.buttonStartExample.Image = CType(resources.GetObject("buttonStartExample.Image"), System.Drawing.Image)
        Me.buttonStartExample.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.buttonStartExample.Location = New System.Drawing.Point(28, 25)
        Me.buttonStartExample.Name = "buttonStartExample"
        Me.buttonStartExample.Size = New System.Drawing.Size(683, 28)
        Me.buttonStartExample.TabIndex = 18
        Me.buttonStartExample.Text = "Start example"
        Me.buttonStartExample.UseVisualStyleBackColor = True
        '
        'Example08
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(201, Byte), Integer), CType(CType(227, Byte), Integer), CType(CType(243, Byte), Integer))
        Me.Controls.Add(Me.labelEventLogHeader)
        Me.Controls.Add(Me.textBoxEvents)
        Me.Controls.Add(Me.textBoxDescription)
        Me.Controls.Add(Me.buttonStartExample)
        Me.Name = "Example08"
        Me.Size = New System.Drawing.Size(739, 304)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents labelEventLogHeader As System.Windows.Forms.Label
    Private WithEvents textBoxEvents As System.Windows.Forms.TextBox
    Private WithEvents textBoxDescription As System.Windows.Forms.TextBox
    Private WithEvents buttonStartExample As System.Windows.Forms.Button

End Class
