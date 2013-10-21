Imports NetOffice
Imports Excel = NetOffice.ExcelApi

Public Class Tutorial06
    Implements ITutorial

    Dim _hostApplication As IHost

#Region "ITutorial Member"

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' this examples shows how i can use variant types(object in NetOffice) at runtime
        ' the reason for the most variant definitions in office is a more flexible value set.(95%)
        ' here is the code to demonstrate this

        ' start application
        Dim application As New Excel.Application()
        application.DisplayAlerts = False

        ' create new Workbook and a new named style
        Dim book As Excel.Workbook = application.Workbooks.Add()
        Dim sheet As Excel.Worksheet = book.Worksheets(1)
        Dim range As Excel.Range = sheet.Cells(1, 1)
        Dim myStyle As Excel.Style = book.Styles.Add("myUniqueStyle")

        ' Range.Style is defined as Variant in Excel and represents as object in NetOffice
        ' You got always an Excel.Style instance if you ask for
        Dim style As Excel.Style = range.Style

        'and here comes the magic. both sets are valid because the variants was very flexible in the setter
        range.Style = "myUniqueStyle"
        range.Style = myStyle

        ' Name, Bold, Size are string, bool and double but defined as Variant 
        style.Font.Name = "Arial"
        style.Font.Bold = True ' you can also set "True" and it works. variants makes it possible
        style.Font.Size = 14

        ' quit & dipose
        application.Quit()
        application.Dispose()

        _hostApplication.ShowFinishDialog()

    End Sub

    Public ReadOnly Property Caption As String Implements TutorialsBase.ITutorial.Caption
        Get
            Return "Tutorial06"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements TutorialsBase.ITutorial.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Understanding Variant", "Verstehen und verwenden von Variant Typen")
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
            Return IIf(_hostApplication.LCID = 1033, "http://netoffice.codeplex.com/wikipage?title=Tutorial06_EN_VB", "http://netoffice.codeplex.com/wikipage?title=Tutorial06_DE_VB")
        End Get
    End Property

#End Region

End Class
