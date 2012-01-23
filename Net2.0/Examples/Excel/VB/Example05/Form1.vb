Imports System.Reflection

Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports NetOffice.OfficeApi.Enums

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

        ' we need some data to display
        Dim dataRange As Excel.Range = PutSampleData(workSheet)

        ' create a nice diagram
        Dim chartObjects As Excel.ChartObjects = workSheet.ChartObjects()
        Dim chart As Excel.ChartObject = chartObjects.Add(70, 100, 375, 225)
        chart.Chart.SetSourceData(dataRange)

        ' save the book 
        Dim fileExtension As String = GetDefaultExtension(excelApplication)
        Dim workbookFile As String = String.Format("{0}\Example05{1}", Application.StartupPath, fileExtension)
        workBook.SaveAs(workbookFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlExclusive)

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

        Dim fDialog As New FinishDialog("Workbook saved.", workbookFile)
        fDialog.ShowDialog(Me)

    End Sub

#Region "Helper"

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

        Return workSheet.get_Range("$B2:$E6")

    End Function

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
