Imports System.Globalization
Imports Excel = NetOffice.ExcelApi
Imports Tests.Core

Public Class Test03
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Set numberformat in cells."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test03"
        End Get
    End Property

    Public ReadOnly Property OfficeProduct As String Implements Tests.Core.ITestPackage.OfficeProduct
        Get
            Return "Excel"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Excel.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New NetOffice.ExcelApi.Application()
            application.DisplayAlerts = False
            application.Workbooks.Add()

            Dim workSheet As Excel.Worksheet = application.Workbooks(1).Sheets(1)

            ' some kind of numerics

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
            workSheet.get_Range("D4").NumberFormat = Pattern2

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
            workSheet.Range("B10").Value = cultureInfo.DateTimeFormat.FullDateTimePattern
            workSheet.Range("C10").Value = cultureInfo.DateTimeFormat.LongDatePattern
            workSheet.Range("D10").Value = cultureInfo.DateTimeFormat.ShortDatePattern
            workSheet.Range("E10").Value = cultureInfo.DateTimeFormat.LongTimePattern
            workSheet.Range("F10").Value = cultureInfo.DateTimeFormat.ShortTimePattern

            ' DateTime
            Dim dateTimeValue As DateTime = DateTime.Now
            workSheet.Range("B11").Value = dateTimeValue
            workSheet.Range("B11").NumberFormat = cultureInfo.DateTimeFormat.FullDateTimePattern

            workSheet.Range("C11").Value = dateTimeValue
            workSheet.Range("C11").NumberFormat = cultureInfo.DateTimeFormat.LongDatePattern

            workSheet.Range("D11").Value = dateTimeValue
            workSheet.Range("D11").NumberFormat = cultureInfo.DateTimeFormat.ShortDatePattern

            workSheet.Range("E11").Value = dateTimeValue
            workSheet.Range("E11").NumberFormat = cultureInfo.DateTimeFormat.LongTimePattern

            workSheet.Range("F11").Value = dateTimeValue
            workSheet.Range("F11").NumberFormat = cultureInfo.DateTimeFormat.ShortTimePattern

            ' string
            workSheet.Range("A14").Value = "String"
            workSheet.Range("B14").Value = "This is a sample String"
            workSheet.Range("B14").NumberFormat = "@"

            'number as string
            workSheet.Range("B15").Value = "513"
            workSheet.Range("B15").NumberFormat = "@"

            ' set colums
            workSheet.Columns(1).AutoFit()
            workSheet.Columns(2).AutoFit()
            workSheet.Columns(3).AutoFit()
            workSheet.Columns(4).AutoFit()

            Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            If Not IsNothing(application) Then
                application.Quit()
                application.Dispose()
            End If

        End Try

    End Function
End Class
