Imports Excel = NetOffice.ExcelApi

Public Class Tutorial02
    Implements ITutorial

    Dim _hostApplication As IHost

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' this example shows you another dispose method: DisposeChildInstances
        ' this means all child proxies from an instance

        ' start application
        Dim application As New Excel.ApplicationClass()
        application.DisplayAlerts = False

        Dim book As Excel.Workbook = application.Workbooks.Add()
        Dim sheet As Excel.Worksheet = book.Worksheets.Add()
        ' we have 5 created proxies now in proxy table as follows
        ' 
        ' Application
        '  + Workbooks
        '     + Workbook  
        '        + Worksheets  
        '           + Worksheet  
        '

        ' we dispose the child instances of book
        book.DisposeChildInstances()

        ' we have 3 created proxies now, the childs from book are disposed
        ' 
        ' Application
        '   + Workbooks
        '    + Workbook  
        '

        application.Quit()
        application.Dispose()
        '
        'the Dispose() call for application release the instance and created childs Workbooks and Workbook
        '

        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial02"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return "Using Dispose & DisposeChildInstances"
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
            Return FormMain.DocumentationBase & "Tutorial02_EN_VB.html"
        End Get
    End Property

End Class
