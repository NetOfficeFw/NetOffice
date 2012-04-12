Imports Excel = NetOffice.ExcelApi

Public Class Tutorial09
    Implements ITutorial

    Dim _hostApplication As IHost

#Region "ITutorial Member"

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' In some situations you want use NetOffice with an existing proxy, its typical for COM Addins.
        ' this examples show you how its possible

        ' Initialize Netoffice
        LateBindingApi.Core.Factory.Initialize()

        ' we create a native Excel proxy
        Dim excelType As Type = Type.GetTypeFromProgID("Excel.Application")
        Dim excelProxy As Object = Activator.CreateInstance(excelType)

        ' we create an Excel Application object with the proxy as parameter,
        ' excel is now under control by NetOffice
        Dim excelApplication As Excel.Application = New Excel.Application(Nothing, excelProxy)

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
            Return IIf(_hostApplication.LCID = 1033, "Create a NetOffice Application with given COM Proxy", "Eine NetOffice Application basierend auf einem COM Proxy erstellen")
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
