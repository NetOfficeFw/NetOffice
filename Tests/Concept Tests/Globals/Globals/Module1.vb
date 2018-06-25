Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.ModulesLegacy.ApplicationModule

Module Module1

    Sub Main()

        Dim excelApplication As New Excel.ApplicationClass()
        excelApplication.DisplayAlerts = False
        excelApplication.Workbooks.Add()

        '' GlobalModule ''

        'active workbook
        Dim workBookName As String = ActiveWorkbook.Name
        Console.WriteLine("ActiveWorkbook.Name: {0}", workBookName)

        'write test value in active sheet
        ActiveSheet.Range("A1").Value = "myValue"

        excelApplication.Quit()
        excelApplication.Dispose()

        Console.WriteLine("Press any key...")
        Console.ReadKey()

    End Sub

End Module
