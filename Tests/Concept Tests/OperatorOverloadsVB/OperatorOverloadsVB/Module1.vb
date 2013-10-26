Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.GlobalHelperModules.GlobalModule

Module Module1
    '
    ' Test: NetOffice.Settings.EnableOperatorOverlads (true by default)
    '
    Sub Main()

        ' start excel
        Dim application As New Excel.Application()
        application.DisplayAlerts = False

        ' add 3 workbooks
        Dim book1 As Excel.Workbook = application.Workbooks.Add()
        Dim book2 As Excel.Workbook = application.Workbooks.Add()
        Dim book3 As Excel.Workbook = application.Workbooks.Add()

        ' check "==" operator
        If ActiveWorkbook = book1 Then
            Console.WriteLine("Book 1 is ActiveWorkbook")
        ElseIf ActiveWorkbook = book2 Then
            Console.WriteLine("Book 2 is ActiveWorkbook")
        ElseIf ActiveWorkbook = book3 Then
            Console.WriteLine("Book 3 is ActiveWorkbook")
        Else
            Console.WriteLine("Operator Overload failed")
        End If

        ' check "!=" operator
        If Not ActiveWorkbook = book1 Then Console.WriteLine("Book 1 is not ActiveWorkbook")
        If Not ActiveWorkbook = book2 Then Console.WriteLine("Book 2 is not ActiveWorkbook")
        If Not ActiveWorkbook = book3 Then Console.WriteLine("Book 3 is not ActiveWorkbook")


        ' with vb latebining
        If ActiveSheet.Parent = book1 Then
            Console.WriteLine("Parent from ActiveSheet is Book 1")
        ElseIf ActiveSheet.Parent = book2 Then
            Console.WriteLine("Parent from ActiveSheet is Book 2")
        ElseIf ActiveSheet.Parent = book3 Then
            Console.WriteLine("Parent from ActiveSheet is Book 3")
        Else
            Console.WriteLine("Latebining Operator Overload failed")
        End If


        ' close and dispose
        application.Quit()
        application.Dispose()

        Console.WriteLine("Press any key...")
        Console.Read()

    End Sub

End Module
