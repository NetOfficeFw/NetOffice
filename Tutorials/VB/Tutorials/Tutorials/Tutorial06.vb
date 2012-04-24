Imports NetOffice
Imports Excel = NetOffice.ExcelApi

Public Class Tutorial06
    Implements ITutorial

    Dim _hostApplication As IHost

#Region "ITutorial Member"

    Public Sub Run() Implements TutorialsBase.ITutorial.Run

        ' start application
        Dim application As New Excel.Application()
        application.DisplayAlerts = False

        Dim book As Excel.Workbook = application.Workbooks.Add()
        Dim sheet As Excel.Worksheet = book.Worksheets(1)
        Dim range As Excel.Range = sheet.Cells(1, 1)

        ' Style is defined as Variant in Excel and represents as object in NetOffice
        '  You can cast them at runtime without problems
        Dim style As Excel.Style = range.Style

        'variant types can be a scalar type at runtime
        'another example way to use is 
        If (TypeName(range.Style) = "String") Then

            Dim myStyle As String = range.Style

        ElseIf (TypeName(range.Style) = "Style") Then

            Dim myStyle As Excel.Style = range.Style
        End If

        ' Name, Bold, Size are bool but defined as Variant and also converted to object
        style.Font.Name = "Arial"
        style.Font.Bold = True
        style.Font.Size = 14


        ' Please note: the reason for the most variant definition is a more flexible value set.
        ' the Style property from Range returns always a Style object
        ' but if you have a new named style created with the name "myStyle" you can set range.Style = myNewStyleObject; or range.Style = "myStyle"
        ' this kind of flexibility is the primary reason for Variants in Office
        ' in any case, you dont lost the COM Proxy management from NetOffice for Variants. 

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
