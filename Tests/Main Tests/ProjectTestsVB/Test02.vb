Imports Tests.Core
Imports MSProject = NetOffice.MSProjectApi
Imports NetOffice.MSProjectApi.Enums

Public Class Test02
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Test events."
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
            Return "Project"
        End Get
    End Property

    Private Sub ApplicationProjectTaskNewEvent(ByVal pj As MSProject.Project, ID As Integer)

        TaskNewEventCalled = True

    End Sub

    Private Sub ApplicationProjectBeforeTaskChangeEvent(ByVal tsk As MSProject.Task, ByVal Field As PjField, ByVal NewVal As Object, ByRef Cancel As Boolean)

        TaskChangeEventCalled = True

    End Sub

    Private Sub ApplicationProjectBeforeCloseEvent(ByVal pj As MSProject.Project, ByRef Cancel As Boolean)

        BeforeCloseEventCalled = True

    End Sub
     
    Private Sub ApplicationProjectBeforeTaskDeleteEvent(ByVal tsk As MSProject.Task, ByRef Cancel As Boolean)

        TaskDeleteEventCalled = True

    End Sub

  
    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As MSProject.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            NetOffice.Settings.Default.MessageFilter.Enabled = True

            application = New MSProject.Application()

            Dim taskNewHandler As MSProject.Application_ProjectTaskNewEventHandler = AddressOf Me.ApplicationProjectTaskNewEvent
            AddHandler application.ProjectTaskNewEvent, taskNewHandler

            Dim projectBeforeCloseEventHandler As MSProject.Application_ProjectBeforeCloseEventHandler = AddressOf Me.ApplicationProjectBeforeCloseEvent
            AddHandler application.ProjectBeforeCloseEvent, projectBeforeCloseEventHandler

            Dim projectBeforeTaskChangeEventHandler As MSProject.Application_ProjectBeforeTaskChangeEventHandler = AddressOf Me.ApplicationProjectBeforeTaskChangeEvent
            AddHandler application.ProjectBeforeTaskChangeEvent, projectBeforeTaskChangeEventHandler

            Dim projectBeforeTaskDeleteHandler As MSProject.Application_ProjectBeforeTaskDeleteEventHandler = AddressOf Me.ApplicationProjectBeforeTaskDeleteEvent
            AddHandler application.ProjectBeforeTaskDeleteEvent, projectBeforeTaskDeleteHandler

            Dim newProject As MSProject.Project = application.Projects.Add()
            Dim task1 As MSProject.Task = newProject.Tasks.Add("Task 1")
            Dim task2 As MSProject.Task = newProject.Tasks.Add("Task 2")

            task2.Delete()
            application.FileCloseAll(False)

            RemoveHandler application.ProjectTaskNewEvent, taskNewHandler
            RemoveHandler application.ProjectBeforeCloseEvent, projectBeforeCloseEventHandler
            RemoveHandler application.ProjectBeforeTaskChangeEvent, projectBeforeTaskChangeEventHandler
            RemoveHandler application.ProjectBeforeTaskDeleteEvent, projectBeforeTaskDeleteHandler

            If (TaskDeleteEventCalled And TaskChangeEventCalled And BeforeCloseEventCalled And TaskNewEventCalled) Then
                Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")
            Else
                Dim errorMessage As String = ""
                If (Not TaskDeleteEventCalled) Then errorMessage += "ProjectTaskNewEvent failed "
                If (Not TaskChangeEventCalled) Then errorMessage += "ProjectBeforeCloseEvent failed "
                If (Not BeforeCloseEventCalled) Then errorMessage += "ProjectBeforeTaskChangeEvent failed "
                If (Not TaskNewEventCalled) Then errorMessage += "ProjectBeforeTaskDeleteEvent failed "
                Return New TestResult(False, DateTime.Now.Subtract(startTime), errorMessage, Nothing, "")
            End If

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            NetOffice.Settings.Default.MessageFilter.Enabled = False

            If Not IsNothing(application) Then
                application.Quit(False)
                application.Dispose()
            End If

        End Try

    End Function

    Dim TaskDeleteEventCalled As Boolean
    Dim TaskChangeEventCalled As Boolean
    Dim BeforeCloseEventCalled As Boolean
    Dim TaskNewEventCalled As Boolean

End Class
