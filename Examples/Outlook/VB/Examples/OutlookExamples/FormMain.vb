Public Class FormMain

    Public Sub New()

        InitializeComponent()

        Me.Text = "NetOffice Outlook Examples in VB"
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
