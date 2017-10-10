Imports Excel = NetOffice.ExcelApi

Public Class Tutorial13
    Implements ITutorial

    Dim _hostApplication As IHost

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' this examples shows a special method to ask at runtime for a particular method oder property
        ' morevover you can enable the option NetOffice.Settings.EnableSafeMode. 
        ' NetOffice checks(cache supported) for any method or property you call and
        ' throws a EntitiyNotSupportedException if missing

        ' create new instance
        Dim application As New Excel.Application()

        ' check for support at runtime
        Dim enableLivePreviewSupport As Boolean = application.EntityIsAvailable("EnableLivePreview")
        Dim openDatabaseSupport As Boolean = application.Workbooks.EntityIsAvailable("OpenDatabase")

        Dim result As String = "Excel Runtime Check: " + Environment.NewLine
        result += "Support EnableLivePreview: " + enableLivePreviewSupport.ToString() + Environment.NewLine
        result += "Support OpenDatabase:      " + openDatabaseSupport.ToString() + Environment.NewLine

        ' quit and dispose
        application.Quit()
        application.Dispose()

        _hostApplication.ShowMessage(result)

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial13"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return "Version-independent development"
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
            Return FormMain.DocumentationBase & "Tutorial13_EN_VB.html"
        End Get
    End Property

End Class
