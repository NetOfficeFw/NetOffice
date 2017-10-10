Imports ExampleBase
Imports NetOffice
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports NetOffice.ExcelApi.Tools.Contribution

''' <summary>
''' Example 10 - Create PDF Document (Microsoft PDF printer must be run)
''' </summary>
Public Class Example10
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost
    Dim _excelApplication As Excel.Application

    Public Sub RunExample() Implements IExample.RunExample

        ' start excel and turn off msg boxes
        Dim excelApplication As New Excel.Application()
        excelApplication.DisplayAlerts = False

        ' create a utils instance, no need for but helpful to keep the lines of code low
        Dim utils As CommonUtils = New CommonUtils(excelApplication)

        ' add a new workbook
        Dim workBook As Excel.Workbook = excelApplication.Workbooks.Add()
        Dim workSheet As Excel.Worksheet = workBook.Worksheets(1)

        ' draw back color and perform the BorderAround method
        workSheet.Range("$B2:$B5").Interior.Color = Utils.Color.ToDouble(Color.DarkGreen)
        workSheet.Range("$B2:$B5").BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic)

        ' draw back color and border the range explicitly
        workSheet.Range("$D2:$D5").Interior.Color = Utils.Color.ToDouble(Color.DarkGreen)
        workSheet.Range("$D2:$D5").Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlDouble
        workSheet.Range("$D2:$D5").Borders(XlBordersIndex.xlInsideHorizontal).Weight = 4
        workSheet.Range("$D2:$D5").Borders(XlBordersIndex.xlInsideHorizontal).Color = utils.Color.ToDouble(Color.Black)

        Dim workbookFile As String = Nothing
        If (workSheet.EntityIsAvailable("ExportAsFixedFormat")) Then

            workbookFile = System.IO.Path.Combine(_hostApplication.RootDirectory, "Example10.pdf")
            workSheet.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, workbookFile, XlFixedFormatQuality.xlQualityStandard)
        Else

            ' we are sorry - pdf export is not supported in Excel 2003 or below
            workbookFile = utils.File.Combine(_hostApplication.RootDirectory, "Example10", DocumentFormat.Normal)
            workBook.SaveAs(workbookFile)

        End If

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

        ' show end dialog
        _hostApplication.ShowFinishDialog(Nothing, workbookFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example10"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Create a PDF Document"
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As ExampleBase.IHost) Implements ExampleBase.IExample.Connect

        _hostApplication = hostApplication

    End Sub

    Public ReadOnly Property Panel As UserControl Implements IExample.Panel
        Get
            Return Nothing
        End Get
    End Property

End Class
