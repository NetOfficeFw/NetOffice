Imports NetOffice
Imports Excel = NetOffice.ExcelApi

Public Class Tutorial18
    Implements ITutorial

    Dim _hostApplication As IHost

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        '/*
        ' *  NetOffice provides features to compare 2 proxies directly on server.
        ' *
        ' *  2 proxies may different instances but pointing to the same instance on the com server(the office application)
        ' *
        ' *  This is a showstopper to demonstrate a deep comparison.
        ' *
        ' *  -------------------------------------------------------
        ' *  Former NetOffice versions spend operator overloads here.
        ' *  This Is impossible in NetOffice 2.0 And above because
        ' *  NetOffice 2.0 use interfaces instead of classes.
        ' *
        ' */

        ' start application
        Dim application As New Excel.ApplicationClass()
        application.DisplayAlerts = False

        Dim book As Excel.Workbook = application.Workbooks.Add()

        Dim isEqual As Boolean = False

        'determine active workbook is the same as book1 on the server
        isEqual = application.ActiveWorkbook.EqualsOnServer(book)

        ' another static version to do the same
        isEqual = COMObject.EqualsOnServer(application.ActiveWorkbook, book)


        application.Quit()
        application.Dispose()

        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial18"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return "Compare Instances"
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
            Return FormMain.DocumentationBase & "Tutorial18_EN_VB.html"
        End Get
    End Property

End Class
