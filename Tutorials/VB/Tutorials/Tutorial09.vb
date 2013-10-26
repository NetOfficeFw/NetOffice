Imports Excel = NetOffice.ExcelApi

Public Class Tutorial09
    Implements ITutorial

    Dim _hostApplication As IHost

#Region "ITutorial Member"

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' In some situations you want use NetOffice with a already running application.
        ' this examples show you how its possible.

        ' GetActiveInstance take the first instance in memory
        Dim excelApplication As Excel.Application = Excel.Application.GetActiveInstance()

        ' another method is GetActiveInstances:
        ' 
        ' GetActiveInstances takes all instances in memory. dont forget to dispose the instances.
        '            
        ' Dim excelApplications() As Excel.Application = Excel.Application.GetActiveInstance()

        excelApplication.Quit()
        excelApplication.Dispose()

        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial09"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Create a NetOffice Excel Application Object with given COM Proxy", "Eine NetOffice Excel Application Objekt basierend auf einem COM Proxy erstellen")
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As TutorialsBase.IHost) Implements TutorialsBase.ITutorial.Connect

        _hostApplication = hostApplication

    End Sub

    Public Sub ChangeLanguage(ByVal lcid As Integer) Implements TutorialsBase.ITutorial.ChangeLanguage

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
            Return IIf(_hostApplication.LCID = 1033, "http://netoffice.codeplex.com/wikipage?title=Tutorial09_EN_VB", "http://netoffice.codeplex.com/wikipage?title=Tutorial09_DE_VB")
        End Get
    End Property

#End Region

End Class
