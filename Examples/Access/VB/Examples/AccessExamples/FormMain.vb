Public Class FormMain

    Public Sub New()

        InitializeComponent()

        Me.Text = "NetOffice Access Examples in Visual Basic"
        LoadExamples()

    End Sub

    Private Sub LoadExamples()

        LoadExample(New Example01())
        LoadExample(New Example02())
        LoadExample(New Example03())
        LoadExample(New Example04())
        LoadExample(New Example05())

    End Sub

End Class
