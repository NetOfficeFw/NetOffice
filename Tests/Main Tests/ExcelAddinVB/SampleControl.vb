Public Class SampleControl
    Implements NetOffice.ExcelApi.Tools.ITaskPane

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MsgBox("Hello")
    End Sub

    Public Sub OnConnection(ByVal application As NetOffice.ExcelApi.Application, ByVal parentPane As NetOffice.OfficeApi._CustomTaskPane, ByVal customArguments() As Object) Implements NetOffice.ExcelApi.Tools.ITaskPane.OnConnection

        Dim addin As TestAddin = customArguments(0)
        addin.TaskPaneOkay = True

    End Sub

    Public Sub OnDisconnection() Implements NetOffice.ExcelApi.Tools.ITaskPane.OnDisconnection

    End Sub

End Class
