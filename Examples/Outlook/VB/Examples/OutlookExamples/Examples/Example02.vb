Imports NetOffice
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums

Public Class Example02
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start outlook
        Dim outlookApplication = New Outlook.Application()

        ' Create a new TaskItem
        Dim newTask As Outlook.TaskItem = outlookApplication.CreateItem(OlItemType.olTaskItem)

        '  Configure the task at hand and save it.
        newTask.Subject = "Don't forget to check for NetOffice.DeveloperToolbox updates"
        newTask.Body = "check updates here: http://netoffice.codeplex.com/releases"
        newTask.DueDate = DateTime.Now
        newTask.Importance = OlImportance.olImportanceHigh

        newTask.Save()

        'close outlook and dispose
        outlookApplication.Quit()
        outlookApplication.Dispose()

        _hostApplication.ShowFinishDialog("Done!", Nothing)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example02", "Beispiel02")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Create task item", "Ein TaskItem erstellen")
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

#End Region

End Class
