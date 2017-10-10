Public Class FormMain

    Public Sub New()

        InitializeComponent()

        Me.Text = "NetOffice Excel Examples in Visual Basic"
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
        LoadExample(New Example08())
        LoadExample(New Example09())
        LoadExample(New Example10())
    End Sub

End Class
