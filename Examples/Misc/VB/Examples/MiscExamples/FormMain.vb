Public Class FormMain

    Public Sub New()

        InitializeComponent()

        Me.Text = "NetOffice Misc Examples in Visual Basic"
        LoadExamples()

    End Sub

    Private Sub LoadExamples()

        LoadExample(New Example01())
        LoadExample(New Example02())
        LoadExample(New Example03())

    End Sub

End Class
