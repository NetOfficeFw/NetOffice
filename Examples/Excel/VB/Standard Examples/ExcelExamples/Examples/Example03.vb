Imports ExampleBase
Imports NetOffice
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Tools.Contribution

''' <summary>
''' Example 3 - Using Numberformats
''' </summary>
Public Class Example03
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

        '  the given thread culture in all latebinding calls are stored in NetOffice.Settings.
        '  you can change the culture. default is en-us.
        Dim cultureInfo As CultureInfo = NetOffice.Settings.Default.ThreadCulture
        Dim Pattern1 As String = String.Format("0{0}00", cultureInfo.NumberFormat.CurrencyDecimalSeparator)
        Dim Pattern2 As String = String.Format("#{1}##0{0}00", cultureInfo.NumberFormat.CurrencyDecimalSeparator, cultureInfo.NumberFormat.CurrencyGroupSeparator)

        workSheet.Range("A1").Value = "Type"
        workSheet.Range("B1").Value = "Value"
        workSheet.Range("C1").Value = "Formatted " + Pattern1
        workSheet.Range("D1").Value = "Formatted " + Pattern2

        Dim integerValue As Integer = 532234
        workSheet.Range("A3").Value = "Integer"
        workSheet.Range("B3").Value = integerValue
        workSheet.Range("C3").Value = integerValue
        workSheet.Range("C3").NumberFormat = Pattern1
        workSheet.Range("D3").Value = integerValue
        workSheet.Range("D3").NumberFormat = Pattern2

        Dim doubleValue As Double = 23172.64
        workSheet.Range("A4").Value = "double"
        workSheet.Range("B4").Value = doubleValue
        workSheet.Range("C4").Value = doubleValue
        workSheet.Range("C4").NumberFormat = Pattern1
        workSheet.Range("D4").Value = doubleValue
        workSheet.Range("D4").NumberFormat = Pattern2

        Dim floatValue As Single = 84345.9141F
        workSheet.Range("A5").Value = "float"
        workSheet.Range("B5").Value = floatValue
        workSheet.Range("C5").Value = floatValue
        workSheet.Range("C5").NumberFormat = Pattern1
        workSheet.Range("D5").Value = floatValue
        workSheet.Range("D5").NumberFormat = Pattern2

        Dim decimalValue As Decimal = 7251231.313367
        workSheet.Range("A6").Value = "Decimal"
        workSheet.Range("B6").Value = decimalValue
        workSheet.Range("C6").Value = decimalValue
        workSheet.Range("C6").NumberFormat = Pattern1
        workSheet.Range("D6").Value = decimalValue
        workSheet.Range("D6").NumberFormat = Pattern2

        workSheet.Range("A9").Value = "DateTime"
        workSheet.Range("B10").Value = Settings.Default.ThreadCulture.DateTimeFormat.FullDateTimePattern
        workSheet.Range("C10").Value = Settings.Default.ThreadCulture.DateTimeFormat.LongDatePattern
        workSheet.Range("D10").Value = Settings.Default.ThreadCulture.DateTimeFormat.ShortDatePattern
        workSheet.Range("E10").Value = Settings.Default.ThreadCulture.DateTimeFormat.LongTimePattern
        workSheet.Range("F10").Value = Settings.Default.ThreadCulture.DateTimeFormat.ShortTimePattern

        ' DateTime
        Dim dateTimeValue As DateTime = DateTime.Now
        workSheet.Range("B11").Value = dateTimeValue
        workSheet.Range("B11").NumberFormat = Settings.Default.ThreadCulture.DateTimeFormat.FullDateTimePattern

        workSheet.Range("C11").Value = dateTimeValue
        workSheet.Range("C11").NumberFormat = Settings.Default.ThreadCulture.DateTimeFormat.LongDatePattern

        workSheet.Range("D11").Value = dateTimeValue
        workSheet.Range("D11").NumberFormat = Settings.Default.ThreadCulture.DateTimeFormat.ShortDatePattern

        workSheet.Range("E11").Value = dateTimeValue
        workSheet.Range("E11").NumberFormat = Settings.Default.ThreadCulture.DateTimeFormat.LongTimePattern

        workSheet.Range("F11").Value = dateTimeValue
        workSheet.Range("F11").NumberFormat = Settings.Default.ThreadCulture.DateTimeFormat.ShortTimePattern

        ' string
        workSheet.Range("A14").Value = "String"
        workSheet.Range("B14").Value = "This is a sample String"
        workSheet.Range("B14").NumberFormat = "@"

        ' number as string
        workSheet.Range("B15").Value = "513"
        workSheet.Range("B15").NumberFormat = "@"

        ' set colums
        workSheet.Columns(1).AutoFit()
        workSheet.Columns(2).AutoFit()
        workSheet.Columns(3).AutoFit()
        workSheet.Columns(4).AutoFit()

        ' save the book 
        Dim workbookFile As String = utils.File.Combine(_hostApplication.RootDirectory, "Example03", DocumentFormat.Normal)
        workBook.SaveAs(workbookFile)

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

        ' show dialog for the user(you!)
        _hostApplication.ShowFinishDialog(Nothing, workbookFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example03"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Using Numberformats"
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
