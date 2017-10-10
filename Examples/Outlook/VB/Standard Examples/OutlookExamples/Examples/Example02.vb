Imports NetOffice
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Example02
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start outlook by trying to access running application first
        Dim outlookApplication = New Outlook.Application(True)

        ' Create a new TaskItem
        Dim newTask As Outlook.TaskItem = outlookApplication.CreateItem(OlItemType.olTaskItem)

        '  Configure the task at hand and save it.
        newTask.Subject = "Don't forget to check for NoScript updates"
        newTask.Body = "check updates here: https://addons.mozilla.org/de/firefox/addon/noscript"
        newTask.DueDate = DateTime.Now
        newTask.Importance = OlImportance.olImportanceHigh

        newTask.Save()

        'close outlook and dispose
        If Not outlookApplication.FromProxyService Then
            outlookApplication.Quit()
        End If
        outlookApplication.Dispose()

        _hostApplication.ShowFinishDialog("Done!", Nothing)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example02"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Create task item"
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As ExampleBase.IHost) Implements ExampleBase.IExample.Connect

        _hostApplication = hostApplication

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements ExampleBase.IExample.Panel
        Get
            Return Nothing
        End Get
    End Property

End Class
