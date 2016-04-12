Imports NetOffice
Imports Outlook = NetOffice.OutlookApi
Imports NetOffice.OutlookApi.Enums
Imports Tests.Core

Public Class Test02
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Create a task item."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test02"
        End Get
    End Property

    Public ReadOnly Property OfficeProduct As String Implements Tests.Core.ITestPackage.OfficeProduct
        Get
            Return "Outlook"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Outlook.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New NetOffice.OutlookApi.Application()
            NetOffice.OutlookSecurity.Suppress.Enabled = True

            ' Create a new TaskItem
            Dim newTask As Outlook.TaskItem = application.CreateItem(OlItemType.olTaskItem)

            '  Configure the task at hand and save it.
            newTask.Subject = "Test item"
            newTask.Body = "hello"
            newTask.DueDate = DateTime.Now
            newTask.Importance = OlImportance.olImportanceHigh
            newTask.Close(OlInspectorClose.olDiscard)

            Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            If Not IsNothing(application) Then
                application.Quit()
                application.Dispose()
            End If

        End Try

    End Function

End Class
