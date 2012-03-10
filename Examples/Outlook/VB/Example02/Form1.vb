Imports LateBindingApi.Core
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' start outlook
        Dim outlookApplication = New Outlook.Application()

        ' Create a new TaskItem
        Dim newTask As Outlook.TaskItem = outlookApplication.CreateItem(OlItemType.olTaskItem)

        '  Configure the task at hand and save it.
        newTask.Subject = "Don't forget to check for NetOffice.Toolbox updates"
        newTask.Body = "check updates here: http://netoffice.codeplex.com/releases"
        newTask.DueDate = DateTime.Now
        newTask.Importance = OlImportance.olImportanceHigh

        newTask.Save()

        'close outlook and dispose
        outlookApplication.Quit()
        outlookApplication.Dispose()

        MessageBox.Show(Me, "Done!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

End Class
