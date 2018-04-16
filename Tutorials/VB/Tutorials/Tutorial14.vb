Imports NetOffice
Imports NetOffice.CollectionsGeneric
Imports NetOffice.Contribution.CollectionsGeneric
Imports Excel = NetOffice.ExcelApi

Public Class Tutorial14
    Implements ITutorial

    Dim _hostApplication As IHost

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' In some situations you want use NetOffice with an already running application.
        ' this tutorial shows how its possible.

        ' 1)
        '
        ' GetActiveInstance take the first instance in memory
        Dim application As Excel.Application = Excel.Application.GetActiveInstance()
        If Not IsNothing(application) Then
            application.Dispose()
        End If

        ' 2)
        '
        ' GetActiveInstances takes all instances in memory
        Dim applications As IDisposableSequence(Of Excel.Application) = Excel.Application.GetActiveInstances()
        applications.Dispose()

        ' 3)
        '
        ' Use special ctor to try access a running application first
        ' and if its failed create a new application
        application = New Excel.Application(New Core(), True)
        ' quit only if its a new application
        If Not application.FromProxyService Then
            application.Quit()
        End If
        application.Dispose()

        ' 4)
        '
        ' Creates instance from interop proxy
        Dim interopType As Type = Type.GetTypeFromProgID("Excel.Application")
        Dim proxy As Object = Activator.CreateInstance(interopType)
        application = COMObject.Create(Of Excel.Application)(proxy)
        application.Quit()
        application.Dispose()


        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial14"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return "Accessing running applications"
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As TutorialsBase.IHost) Implements TutorialsBase.ITutorial.Connect

        _hostApplication = hostApplication

    End Sub

    Public Sub Disconnect() Implements TutorialsBase.ITutorial.Disconnect

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements TutorialsBase.ITutorial.Panel
        Get
            Return Nothing
        End Get
    End Property

    Public ReadOnly Property Uri As String Implements TutorialsBase.ITutorial.Uri
        Get
            Return FormMain.DocumentationBase & "Tutorial14_EN_VB.html"
        End Get
    End Property

End Class
