Imports Excel = NetOffice.ExcelApi

Public Class Form1

    Private Sub buttonStartExample_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonStartExample.Click

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        ' create new instance
        Dim application As New Excel.Application()

        ' check for support at runtime
        Dim enableLivePreviewSupport As Boolean = application.EntityIsAvailable("EnableLivePreview")
        Dim openDatabaseSupport As Boolean = application.Workbooks.EntityIsAvailable("OpenDatabase")

        Dim result As String = "Excel Runtime Check: " + Environment.NewLine
        result += "Support EnableLivePreview: " + enableLivePreviewSupport.ToString() + Environment.NewLine
        result += "Support OpenDatabase:      " + openDatabaseSupport.ToString() + Environment.NewLine

        richTextBoxResult.Text = result

        ' quit and dispose
        application.Quit()
        application.Dispose()

    End Sub

End Class
