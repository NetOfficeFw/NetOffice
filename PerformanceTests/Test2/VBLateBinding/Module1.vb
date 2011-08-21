Imports System.Runtime.InteropServices

Module Module1

    Sub Main()

        Console.WriteLine("Visual Basic LateBinding Performance Test - 10000 Cells.")
        Console.WriteLine("Write simple text, change Font, NumberFormat and do a BorderArround.")

        '  start excel, and get a new sheet reference
        Dim excelApplication As Object = CreateExcelApplication()
        Dim books As Object = excelApplication.Workbooks
        Dim book As Object = books.Add()
        Dim sheets As Object = book.Worksheets
        Dim sheet As Object = sheets.Add()

        ' do test 10 times
        Dim comReferencesList As New List(Of MarshalByRefObject)
        Dim timeElapsedList As New List(Of TimeSpan)

        For i = 1 To 10

            Dim timeStart As DateTime = DateTime.Now
            For y = 1 To 10000

                Dim rangeAdress As String = "$A" + y.ToString()
                Dim cellRange As Object = sheet.Range(rangeAdress)
                cellRange.Value = "value"
                Dim font As Object = cellRange.Font
                font.Name = "Verdana"
                cellRange.NumberFormat = "@"
                cellRange.BorderAround(-4119, -4138, -4105, 0)
                comReferencesList.Add(font)
                comReferencesList.Add(cellRange)

            Next
            Dim timeElapsed As TimeSpan = DateTime.Now - timeStart

            ' display info and dispose references
            Console.WriteLine("Time Elapsed: {0}", timeElapsed)
            timeElapsedList.Add(timeElapsed)
            For Each item As Object In comReferencesList
                Marshal.ReleaseComObject(item)
            Next
            comReferencesList.Clear()

        Next

        ' display info & log to file
        Dim timeAverage As TimeSpan = AppendResultToLogFile(timeElapsedList, "Test2-VBLateBinding.log")
        Console.WriteLine("Time Average: {0}{1}Press any key...", timeAverage, Environment.NewLine)
        Console.Read()

        ' release & quit
        Marshal.ReleaseComObject(sheet)
        Marshal.ReleaseComObject(sheets)
        Marshal.ReleaseComObject(book)
        Marshal.ReleaseComObject(books)

        excelApplication.Quit()
        Marshal.ReleaseComObject(excelApplication)

    End Sub

    ''' <summary>
    ''' creates a new excel application
    ''' </summary>
    ''' <remarks></remarks>
    Function CreateExcelApplication() As Object

        Dim excelApplication As Object = CreateObject("Excel.Application")
        excelApplication.DisplayAlerts = False
        excelApplication.Interactive = False
        excelApplication.ScreenUpdating = False
        Return excelApplication

    End Function

    ''' <summary>
    ''' writes list items to a logile and append average of items at the end
    ''' </summary>
    ''' <param name="timeElapsedList">a list with log results</param>
    ''' <param name="fileName">name of logfile in current assembly folder</param>
    ''' <returns>average of timeElapsedList</returns>
    Function AppendResultToLogFile(ByVal timeElapsedList As List(Of TimeSpan), ByVal fileName As String) As TimeSpan

        Dim timeSummary As TimeSpan = TimeSpan.Zero
        Dim logFile As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location), fileName)

        If (System.IO.File.Exists(logFile)) Then
            System.IO.File.Delete(logFile)
        End If

        For Each item As TimeSpan In timeElapsedList

            timeSummary += item
            Dim logFileAppend As String = item.ToString() + Environment.NewLine
            System.IO.File.AppendAllText(logFile, logFileAppend, System.Text.Encoding.UTF8)

        Next

        Dim timeAverage As TimeSpan = TimeSpan.FromTicks(timeSummary.Ticks / timeElapsedList.Count)
        System.IO.File.AppendAllText(logFile, "Time Average: " + timeAverage.ToString(), System.Text.Encoding.UTF8)
        Return timeAverage

    End Function

End Module
