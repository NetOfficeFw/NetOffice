Imports NetOffice

Public Class Tutorial07
    Implements ITutorial

    Dim _hostApplication As IHost

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' NetOffice Core supports a so-called managed C# dynamic
        ' with proxy management services. No need for additional NetOffice Api assemblies.

        ' In Visual Basic it is just an Object for late binding -
        ' but with all NetOffice proxy management services.

        ' NetOffice want convert a proxy to COMDynamicObject each time if its failed to resolve
        ' a corresponding wrapper type.

        Dim application As Object = New COMDynamicObject("Excel.Application")
        application.DisplayAlerts = False
        Dim book As Object = application.Workbooks.Add()

        For Each sheet As Object In book.Sheets
            Console.WriteLine(sheet)
        Next sheet

        'quit and dispose all open proxies
        application.Quit()
        application.Dispose()

        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial07"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return "Managed Dynamics"
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
            Return FormMain.DocumentationBase & "Tutorial07_EN_VB.html"
        End Get
    End Property

End Class