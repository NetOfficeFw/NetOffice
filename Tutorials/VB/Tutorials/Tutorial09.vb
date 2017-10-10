Imports NetOffice
Imports Excel = NetOffice.ExcelApi

Public Class Tutorial09
    Implements ITutorial

    Dim _hostApplication As IHost

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' NetOffice instances implements the IClonable interface
        ' and deal with underlying proxies as well

        ' start application
        Dim application As New Excel.Application()
        application.DisplayAlerts = False
        Dim book As Excel.Workbook = application.Workbooks.Add()

        ' clone the book
        Dim cloneBook As Excel.Workbook = book.Clone()

        ' dispose the origin book keep the underlying proxy alive
        ' until the clone Is disposed
        book.Dispose()

        ' alive and works even the origin book is disposed
        For Each sheet As Object In cloneBook.Sheets
            Console.WriteLine(sheet)
        Next sheet

        application.Quit()
        application.Dispose()

        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial09"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return "Cloning Instances"
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
            Return FormMain.DocumentationBase & "Tutorial09_EN_VB.html"
        End Get
    End Property

End Class
