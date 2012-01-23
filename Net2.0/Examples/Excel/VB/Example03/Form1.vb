Imports System.Reflection
Imports System.Globalization

Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        ' start excel and turn off msg boxes
        Dim excelApplication As New Excel.Application()
        excelApplication.DisplayAlerts = False

        ' add a new workbook
        Dim workBook As Excel.Workbook = excelApplication.Workbooks.Add()
        Dim workSheet As Excel.Worksheet = workBook.Worksheets(1)

        ' /* some kind of numerics */

        '  the given thread culture in all latebinding calls are stored in LateBindingApi.Core.Settings.
        '  you can change the culture. default is en-us.
        Dim cultureInfo As CultureInfo = LateBindingApi.Core.Settings.ThreadCulture
        Dim Pattern1 As String = String.Format("0{0}00", cultureInfo.NumberFormat.CurrencyDecimalSeparator)
        Dim Pattern2 As String = String.Format("#{1}##0{0}00", cultureInfo.NumberFormat.CurrencyDecimalSeparator, cultureInfo.NumberFormat.CurrencyGroupSeparator)

        workSheet.get_Range("A1").Value = "Type"
        workSheet.get_Range("B1").Value = "Value"
        workSheet.get_Range("C1").Value = "Formatted " + Pattern1
        workSheet.get_Range("D1").Value = "Formatted " + Pattern2

        Dim integerValue As Integer = 532234
        workSheet.get_Range("A3").Value = "Integer"
        workSheet.get_Range("B3").Value = integerValue
        workSheet.get_Range("C3").Value = integerValue
        workSheet.get_Range("C3").NumberFormat = Pattern1
        workSheet.get_Range("D3").Value = integerValue
        workSheet.get_Range("D3").NumberFormat = Pattern2

        Dim doubleValue As Double = 23172.64
        workSheet.get_Range("A4").Value = "double"
        workSheet.get_Range("B4").Value = doubleValue
        workSheet.get_Range("C4").Value = doubleValue
        workSheet.get_Range("C4").NumberFormat = Pattern1
        workSheet.get_Range("D4").Value = doubleValue
        workSheet.get_Range("D4").NumberFormat = Pattern2

        Dim floatValue As Single = 84345.9141F
        workSheet.get_Range("A5").Value = "float"
        workSheet.get_Range("B5").Value = floatValue
        workSheet.get_Range("C5").Value = floatValue
        workSheet.get_Range("C5").NumberFormat = Pattern1
        workSheet.get_Range("D5").Value = floatValue
        workSheet.get_Range("D5").NumberFormat = Pattern2

        Dim decimalValue As Decimal = 7251231.313367
        workSheet.get_Range("A6").Value = "Decimal"
        workSheet.get_Range("B6").Value = decimalValue
        workSheet.get_Range("C6").Value = decimalValue
        workSheet.get_Range("C6").NumberFormat = Pattern1
        workSheet.get_Range("D6").Value = decimalValue
        workSheet.get_Range("D6").NumberFormat = Pattern2

        workSheet.get_Range("A9").Value = "DateTime"
        workSheet.get_Range("B10").Value = cultureInfo.DateTimeFormat.FullDateTimePattern
        workSheet.get_Range("C10").Value = cultureInfo.DateTimeFormat.LongDatePattern
        workSheet.get_Range("D10").Value = cultureInfo.DateTimeFormat.ShortDatePattern
        workSheet.get_Range("E10").Value = cultureInfo.DateTimeFormat.LongTimePattern
        workSheet.get_Range("F10").Value = cultureInfo.DateTimeFormat.ShortTimePattern

        ' DateTime
        Dim dateTimeValue As DateTime = DateTime.Now
        workSheet.get_Range("B11").Value = dateTimeValue
        workSheet.get_Range("B11").NumberFormat = cultureInfo.DateTimeFormat.FullDateTimePattern

        workSheet.get_Range("C11").Value = dateTimeValue
        workSheet.get_Range("C11").NumberFormat = cultureInfo.DateTimeFormat.LongDatePattern

        workSheet.get_Range("D11").Value = dateTimeValue
        workSheet.get_Range("D11").NumberFormat = cultureInfo.DateTimeFormat.ShortDatePattern

        workSheet.get_Range("E11").Value = dateTimeValue
        workSheet.get_Range("E11").NumberFormat = cultureInfo.DateTimeFormat.LongTimePattern

        workSheet.get_Range("F11").Value = dateTimeValue
        workSheet.get_Range("F11").NumberFormat = cultureInfo.DateTimeFormat.ShortTimePattern

        ' string
        workSheet.get_Range("A14").Value = "String"
        workSheet.get_Range("B14").Value = "This is a sample String"
        workSheet.get_Range("B14").NumberFormat = "@"

        ' number as string
        workSheet.get_Range("B15").Value = "513"
        workSheet.get_Range("B15").NumberFormat = "@"

        ' set colums
        workSheet.Columns(1).AutoFit()
        workSheet.Columns(2).AutoFit()
        workSheet.Columns(3).AutoFit()
        workSheet.Columns(4).AutoFit()

        ' save the book 
        Dim fileExtension As String = GetDefaultExtension(excelApplication)
        Dim workbookFile As String = String.Format("{0}\Example03{1}", Application.StartupPath, fileExtension)
        workBook.SaveAs(workbookFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlExclusive)

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

        Dim fDialog As New FinishDialog("Workbook saved.", workbookFile)
        fDialog.ShowDialog(Me)

    End Sub

#Region "Helper"

    ''' <summary>
    ''' returns the valid file extension for the instance. for example ".xls" or ".xlsx"
    ''' </summary>
    ''' <param name="application">the instance</param>
    ''' <returns>the extension</returns>
    ''' <remarks></remarks>
    Private Function GetDefaultExtension(ByVal application As Excel.Application) As String

        Dim version As Double = application.Version
        If (version >= 120.0) Then
            Return ".xlsx"
        Else
            Return ".xls"
        End If

    End Function

#End Region

End Class
