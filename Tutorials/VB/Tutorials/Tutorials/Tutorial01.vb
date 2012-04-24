Imports Excel = NetOffice.ExcelApi

Public Class Tutorial01
    Implements ITutorial

    Dim _hostApplication As IHost

#Region "ITutorial Member"

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        '  NetOffice manages COM Proxies for you to avoid any kind of memory leaks
        '  and make sure your application instance removes from process list if you want.

        ' start application
        Dim application As New Excel.Application()
        application.DisplayAlerts = False

        Dim book As Excel.Workbook = application.Workbooks.Add()
        ' now we have 2 new COM Proxies created.
        ' 
        ' the first proxy was created while accessing the Workbooks collection from application
        ' the second proxy was created by the Add() method from Workbooks and stored now in book
        ' with the application object we have 3 created proxies now. the workbooks proxy was created
        ' about application and the book proxy was created about the workbooks.
        ' NetOffice holds the proxies now in a list as follows:
        ' 
        ' Application
        '   + Workbooks
        '     + Workbook  
        ' 
        ' any object in NetOffice implements the IDisposible Interface.
        ' use the Dispose() Method to release an object. the method release all created child proxies too.

        application.Quit()
        application.Dispose()
        ' the application object is ouer root object
        ' dispose them release himself and any childs of application, in this case workbooks and workbook
        ' the excel instance are now removed from process list

        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial01"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Understand COM Proxy Management", "COM Proxy Management verstehen")
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
            Return IIf(_hostApplication.LCID = 1033, "http://netoffice.codeplex.com/wikipage?title=Tutorial01_EN_VB", "http://netoffice.codeplex.com/wikipage?title=Tutorial01_DE_VB")
        End Get
    End Property

#End Region

End Class
