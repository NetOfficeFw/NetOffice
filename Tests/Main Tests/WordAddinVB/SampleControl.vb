Public Class SampleControl
    Implements NetOffice.WordApi.Tools.ITaskPane

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MsgBox("Hello")
    End Sub

    Public Sub OnConnection(ByVal application As NetOffice.WordApi.Application, ByVal parentPane As NetOffice.OfficeApi._CustomTaskPane, ByVal customArguments() As Object) Implements NetOffice.WordApi.Tools.ITaskPane.OnConnection

        Dim addin As TestAddin = customArguments(0)
        addin.TaskPaneOkay = True

    End Sub

    Public Sub OnDisconnection() Implements NetOffice.WordApi.Tools.ITaskPane.OnDisconnection

    End Sub

    Public Sub OnDockPositionChanged(ByVal position As NetOffice.OfficeApi.Enums.MsoCTPDockPosition) Implements NetOffice.WordApi.Tools.ITaskPane.OnDockPositionChanged

    End Sub


    Public Sub OnVisibleStateChanged(ByVal visible As Boolean) Implements NetOffice.WordApi.Tools.ITaskPane.OnVisibleStateChanged

    End Sub

End Class

