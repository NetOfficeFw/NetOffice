Imports Tests.Core
Imports MSProject = NetOffice.MSProjectApi

Public Class Test01
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Add a new project and add 10 tasks."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test01"
        End Get
    End Property

    Public ReadOnly Property OfficeProduct As String Implements Tests.Core.ITestPackage.OfficeProduct
        Get
            Return "Project"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As MSProject.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New MSProject.Application()

            Dim newProject As MSProject.Project = application.Projects.Add()

            newProject.Tasks.Add("Task 0")
            newProject.Tasks.Add("Task 1")
            newProject.Tasks.Add("Task 2")
            newProject.Tasks.Add("Task 4")
            newProject.Tasks.Add("Task 5")
            newProject.Tasks.Add("Task 6")
            newProject.Tasks.Add("Task 7")
            newProject.Tasks.Add("Task 8")
            newProject.Tasks.Add("Task 9")

            Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            If Not IsNothing(application) Then
                application.Quit(False)
                application.Dispose()
            End If

        End Try

    End Function

End Class
