Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.GlobalHelperModules.GlobalModule

Module Module1

    Sub Main()

        Dim excelApplication As New Excel.Application()
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
