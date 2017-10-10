Imports ExampleBase
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Tools.Contribution

''' <summary>
''' Example 5 - Working with Charts
''' </summary>
Public Class Example05
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start excel and turn off msg boxes
        Dim excelApplication As New Excel.Application()
        excelApplication.DisplayAlerts = False

        ' create a utils instance, no need for but helpful to keep the lines of code low
        Dim utils As CommonUtils = New CommonUtils(excelApplication)

        ' add a new workbook
        Dim workBook As Excel.Workbook = excelApplication.Workbooks.Add()
        Dim workSheet As Excel.Worksheet = workBook.Worksheets(1)

        ' we need some data to display
        Dim dataRange As Excel.Range = PutSampleData(workSheet)

        ' create a nice diagram
        Dim chartObjects As Excel.ChartObjects = workSheet.ChartObjects()
        Dim chart As Excel.ChartObject = chartObjects.Add(70, 100, 375, 225)
        chart.Chart.SetSourceData(dataRange)

        ' save the book 
        Dim workbookFile As String = utils.File.Combine(_hostApplication.RootDirectory, "Example05", DocumentFormat.Normal)
        workBook.SaveAs(workbookFile)

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

        ' show dialog for the user(you!)
        _hostApplication.ShowFinishDialog(Nothing, workbookFile)

    End Sub

    Private Function PutSampleData(ByVal workSheet As Excel.Worksheet) As Excel.Range

        workSheet.Cells(2, 2).Value = "Datum"
        workSheet.Cells(3, 2).Value = DateTime.Now.ToShortDateString()
        workSheet.Cells(4, 2).Value = DateTime.Now.ToShortDateString()
        workSheet.Cells(5, 2).Value = DateTime.Now.ToShortDateString()
        workSheet.Cells(6, 2).Value = DateTime.Now.ToShortDateString()

        workSheet.Cells(2, 3).Value = "Columns1"
        workSheet.Cells(3, 3).Value = 25
        workSheet.Cells(4, 3).Value = 33
        workSheet.Cells(5, 3).Value = 30
        workSheet.Cells(6, 3).Value = 22

        workSheet.Cells(2, 4).Value = "Column2"
        workSheet.Cells(3, 4).Value = 25
        workSheet.Cells(4, 4).Value = 33
        workSheet.Cells(5, 4).Value = 30
        workSheet.Cells(6, 4).Value = 22

        workSheet.Cells(2, 5).Value = "Column3"
        workSheet.Cells(3, 5).Value = 25
        workSheet.Cells(4, 5).Value = 33
        workSheet.Cells(5, 5).Value = 30
        workSheet.Cells(6, 5).Value = 22

        Return workSheet.Range("$B2:$E6")

    End Function

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example05"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Working with Charts"
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

End Class
