Imports Excel = NetOffice.ExcelApi

Public Class Tutorial04
    Implements ITutorial

    Dim _hostApplication As IHost

#Region "ITutorial Member"

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' this example shows you how i still can recieve events from an disposed proxy.
        ' you have to use th Dispose oder DisposeChildInstances method with a parameter.

        ' start application
        Dim application As New Excel.Application()
        application.DisplayAlerts = False

        ' create new Workbook & attach close event trigger
        Dim book As Excel.Workbook = application.Workbooks.Add()

        Dim closeHandler As Excel.Workbook_BeforeCloseEventHandler = AddressOf Me.book_BeforeCloseEvent
        AddHandler book.BeforeCloseEvent, closeHandler

        ' we dispose the instance. the parameter false signals to api dont release the event listener
        ' set parameter to true and the event listener will stopped and you dont get events for the instance
        ' the DisposeChildInstances() method has the same method overload
        book.Close()
        book.Dispose(False)

        application.Quit()
        application.Dispose()
        ' 
        ' the application object is ouer root object
        ' dispose them release himself and any childs of application, in this case workbooks and workbook
        ' the excel instance are now removed from process list
        ' 

        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial04"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Using Dispose with event exporting Objects", "Verwenden von Dispose mit Objekten die Ereignisse auslösen")
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
            Return IIf(_hostApplication.LCID = 1033, "http://netoffice.codeplex.com/wikipage?title=Tutorial04_EN_VB", "http://netoffice.codeplex.com/wikipage?title=Tutorial04_DE_VB")
        End Get
    End Property

#End Region

    Public Sub book_BeforeCloseEvent(ByRef Cancel As Boolean)

    End Sub

End Class
