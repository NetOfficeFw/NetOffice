<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SamplePane
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SamplePane))
        Me.timerRunningTime = New System.Windows.Forms.Timer(Me.components)
        Me.imageListButtons = New System.Windows.Forms.ImageList(Me.components)
        Me.labelHint = New System.Windows.Forms.Label()
        Me.pictureBox1 = New System.Windows.Forms.PictureBox()
        Me.buttonReset = New System.Windows.Forms.Button()
        Me.label1 = New System.Windows.Forms.Label()
        Me.labelTime = New System.Windows.Forms.Label()
        Me.buttonEnabled = New System.Windows.Forms.Button()
        CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'timerRunningTime
        '
        Me.timerRunningTime.Interval = 900
        '
        'imageListButtons
        '
        Me.imageListButtons.ImageStream = CType(resources.GetObject("imageListButtons.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imageListButtons.TransparentColor = System.Drawing.Color.Transparent
        Me.imageListButtons.Images.SetKeyName(0, "alarmclock_run.png")
        Me.imageListButtons.Images.SetKeyName(1, "alarmclock_stop.png")
        Me.imageListButtons.Images.SetKeyName(2, "delete2.png")
        '
        'labelHint
        '
        Me.labelHint.AutoSize = True
        Me.labelHint.Location = New System.Drawing.Point(529, 7)
        Me.labelHint.Name = "labelHint"
        Me.labelHint.Size = New System.Drawing.Size(256, 16)
        Me.labelHint.TabIndex = 17
        Me.labelHint.Text = "NetOffice Tools - Extended Sample Addin"
        '
        'pictureBox1
        '
        Me.pictureBox1.Image = CType(resources.GetObject("pictureBox1.Image"), System.Drawing.Image)
        Me.pictureBox1.Location = New System.Drawing.Point(509, 7)
        Me.pictureBox1.Name = "pictureBox1"
        Me.pictureBox1.Size = New System.Drawing.Size(20, 17)
        Me.pictureBox1.TabIndex = 16
        Me.pictureBox1.TabStop = False
        '
        'buttonReset
        '
        Me.buttonReset.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.buttonReset.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.buttonReset.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.buttonReset.ImageKey = "delete2.png"
        Me.buttonReset.ImageList = Me.imageListButtons
        Me.buttonReset.Location = New System.Drawing.Point(127, 0)
        Me.buttonReset.Margin = New System.Windows.Forms.Padding(2)
        Me.buttonReset.Name = "buttonReset"
        Me.buttonReset.Size = New System.Drawing.Size(126, 30)
        Me.buttonReset.TabIndex = 15
        Me.buttonReset.Text = "Reset"
        Me.buttonReset.UseVisualStyleBackColor = True
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.label1.Location = New System.Drawing.Point(279, 7)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(94, 16)
        Me.label1.TabIndex = 14
        Me.label1.Text = "Running Time:"
        '
        'labelTime
        '
        Me.labelTime.AutoSize = True
        Me.labelTime.BackColor = System.Drawing.Color.Transparent
        Me.labelTime.ForeColor = System.Drawing.SystemColors.ControlText
        Me.labelTime.Location = New System.Drawing.Point(371, 7)
        Me.labelTime.Name = "labelTime"
        Me.labelTime.Size = New System.Drawing.Size(56, 16)
        Me.labelTime.TabIndex = 13
        Me.labelTime.Text = "00:00:00"
        '
        'buttonEnabled
        '
        Me.buttonEnabled.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.buttonEnabled.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.buttonEnabled.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.buttonEnabled.ImageKey = "alarmclock_run.png"
        Me.buttonEnabled.ImageList = Me.imageListButtons
        Me.buttonEnabled.Location = New System.Drawing.Point(0, 0)
        Me.buttonEnabled.Margin = New System.Windows.Forms.Padding(2)
        Me.buttonEnabled.Name = "buttonEnabled"
        Me.buttonEnabled.Size = New System.Drawing.Size(126, 30)
        Me.buttonEnabled.TabIndex = 12
        Me.buttonEnabled.Text = "Enable"
        Me.buttonEnabled.UseVisualStyleBackColor = True
        '
        'SamplePane
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.labelHint)
        Me.Controls.Add(Me.pictureBox1)
        Me.Controls.Add(Me.buttonReset)
        Me.Controls.Add(Me.label1)
        Me.Controls.Add(Me.labelTime)
        Me.Controls.Add(Me.buttonEnabled)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "SamplePane"
        Me.Size = New System.Drawing.Size(807, 30)
        CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents timerRunningTime As System.Windows.Forms.Timer
    Private WithEvents imageListButtons As System.Windows.Forms.ImageList
    Private WithEvents labelHint As System.Windows.Forms.Label
    Private WithEvents pictureBox1 As System.Windows.Forms.PictureBox
    Private WithEvents buttonReset As System.Windows.Forms.Button
    Private WithEvents label1 As System.Windows.Forms.Label
    Private WithEvents labelTime As System.Windows.Forms.Label
    Private WithEvents buttonEnabled As System.Windows.Forms.Button

End Class
