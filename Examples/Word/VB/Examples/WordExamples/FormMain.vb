Public Class FormMain

    Public Sub New()

        InitializeComponent()

        Me.Text = "NetOffice Word Examples in Visual Basic"
        LoadExamples()

    End Sub

    Private Sub LoadExamples()

        LoadExample(New Example01())
        LoadExample(New Example02())
        LoadExample(New Example03())
        LoadExample(New Example04())
        LoadExample(New Example05())
        LoadExample(New Example06())
        LoadExample(New Example07())

    End Sub

End Class
