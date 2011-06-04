Imports Outlook = NetOffice.OutlookApi

Public Class Form1

    Private Sub buttonStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStart.Click

        ' initialize api
        LateBindingApi.Core.Factory.Initialize()

        Dim application As Outlook.Application = Nothing

        Dim nativeProxy As Object = RunningObjectTable.GetRunningOutlookInstanceFromROT()
        If (Not IsNothing(nativeProxy)) Then

            application = New NetOffice.OutlookApi.Application(Nothing, nativeProxy)
            textBoxLog.Clear()
            textBoxLog.AppendText("we got running outlook instance" + vbNewLine)
            textBoxLog.AppendText("outlook version is " + application.Version)

            'instance was already running at start. we dispose references but not quit application
            application.Dispose()

        Else

            application = New NetOffice.OutlookApi.Application()

            textBoxLog.Clear()
            textBoxLog.AppendText("we create new outlook instance" + vbNewLine)
            textBoxLog.AppendText("outlook version is " + application.Version)

            'quit and dispose application
            application.Quit()
            application.Dispose()

        End If

    End Sub

End Class
