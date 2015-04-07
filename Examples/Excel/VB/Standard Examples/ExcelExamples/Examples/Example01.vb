Imports ExampleBase
Imports NetOffice
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Example01
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start excel and turn off msg boxes
        Dim excelApplication As New Excel.Application()
        excelApplication.DisplayAlerts = False

        ' create a utils instance, not need for but helpful to keep the lines of code low
        Dim utils As Excel.Tools.CommonUtils = New Excel.Tools.CommonUtils(excelApplication)

        ' add a new workbook
        Dim workBook As Excel.Workbook = excelApplication.Workbooks.Add()
        Dim workSheet As Excel.Worksheet = workBook.Worksheets(1)

        ' draw back color and perform the BorderAround method
        workSheet.Range("$B2:$B5").Interior.Color = utils.Color.ToDouble(Color.DarkGreen)
        workSheet.Range("$B2:$B5").BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic)

        ' draw back color and border the range explicitly
        workSheet.Range("$D2:$D5").Interior.Color = utils.Color.ToDouble(Color.DarkGreen)
        workSheet.Range("$D2:$D5").Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlDouble
        workSheet.Range("$D2:$D5").Borders(XlBordersIndex.xlInsideHorizontal).Weight = 4
        workSheet.Range("$D2:$D5").Borders(XlBordersIndex.xlInsideHorizontal).Color = utils.Color.ToDouble(Color.Black)

        workSheet.Cells(1, 1).Value = "We have 2 simple shapes created."

        'save document
        Dim workbookFile As String = utils.File.Combine(_hostApplication.RootDirectory, "Example01", Excel.Tools.DocumentFormat.Normal)
        workBook.SaveAs(workbookFile)

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

        ' show dialog for the user(you!)
        _hostApplication.ShowFinishDialog(Nothing, workbookFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example01", "Beispiel01")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Background Colors and Borders for Cells", "Hintergrundfarben und Rahmen in Zellen")
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As ExampleBase.IHost) Implements ExampleBase.IExample.Connect

        _hostApplication = hostApplication

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements ExampleBase.IExample.Panel
        Get
            Return Nothing
        End Get
    End Property

#End Region

End Class
