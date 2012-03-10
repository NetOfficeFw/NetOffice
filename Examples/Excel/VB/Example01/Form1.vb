Imports System.Reflection

Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums


Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize Api COMObject Support
        LateBindingApi.Core.Factory.Initialize()

        ' start excel and turn off msg boxes
        Dim excelApplication As New Excel.Application()
        excelApplication.DisplayAlerts = False

        ' add a new workbook
        Dim workBook As Excel.Workbook = excelApplication.Workbooks.Add()
        Dim workSheet As Excel.Worksheet = workBook.Worksheets(1)


        ' draw back color and perform the BorderAround method
        workSheet.Range("$B2:$B5").Interior.Color = ToDouble(Color.DarkGreen)
        workSheet.Range("$B2:$B5").BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic)

        ' draw back color and border the range explicitly
        workSheet.Range("$D2:$D5").Interior.Color = ToDouble(Color.DarkGreen)
        workSheet.Range("$D2:$D5").Borders(XlBordersIndex.xlInsideHorizontal).LineStyle = XlLineStyle.xlDouble
        workSheet.Range("$D2:$D5").Borders(XlBordersIndex.xlInsideHorizontal).Weight = 4
        workSheet.Range("$D2:$D5").Borders(XlBordersIndex.xlInsideHorizontal).Color = ToDouble(Color.Black)


        ' save the book
        Dim fileExtension As String = GetDefaultExtension(excelApplication)
        Dim workbookFile As String = String.Format("{0}\Example01{1}", Environment.CurrentDirectory, fileExtension)
        workBook.SaveAs(workbookFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlExclusive)

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

        Dim fDialog As New FinishDialog("Workbook saved.", workbookFile)
        fDialog.ShowDialog(Me)

    End Sub


#Region "Helper"

    ''' <summary>
    ''' Translate a color to double
    ''' </summary>
    ''' <param name="color">expression to convert</param>
    ''' <returns>color</returns>
    ''' <remarks></remarks>
    Private Function ToDouble(ByVal color As System.Drawing.Color) As Double

        Dim returnValue As UInteger = color.B
        returnValue = returnValue << 8
        returnValue += color.G
        returnValue = returnValue << 8
        returnValue += color.R
        Return returnValue

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
