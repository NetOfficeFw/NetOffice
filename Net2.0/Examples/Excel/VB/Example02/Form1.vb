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

        ' font action
        workSheet.get_Range("A1").Value = "Arial Size:8 Bold Italic Underline"
        workSheet.get_Range("A1").Font.Name = "Arial"
        workSheet.get_Range("A1").Font.Size = 8
        workSheet.get_Range("A1").Font.Bold = True
        workSheet.get_Range("A1").Font.Italic = True
        workSheet.get_Range("A1").Font.Underline = True
        workSheet.get_Range("A1").Font.Color = Color.Violet.ToArgb()

        workSheet.get_Range("A3").Value = "Times New Roman Size:10"
        workSheet.get_Range("A3").Font.Name = "Times New Roman"
        workSheet.get_Range("A3").Font.Size = 10
        workSheet.get_Range("A3").Font.Color = Color.Orange.ToArgb()

        workSheet.get_Range("A5").Value = "Comic Sans MS Size:12 WrapText"
        workSheet.get_Range("A5").Font.Name = "Comic Sans MS"
        workSheet.get_Range("A5").Font.Size = 12
        workSheet.get_Range("A5").WrapText = True
        workSheet.get_Range("A5").Font.Color = Color.Navy.ToArgb()

        ' HorizontalAlignment
        workSheet.get_Range("A7").Value = "xlHAlignLeft"
        workSheet.get_Range("A7").HorizontalAlignment = XlHAlign.xlHAlignLeft

        workSheet.get_Range("B7").Value = "xlHAlignCenter"
        workSheet.get_Range("B7").HorizontalAlignment = XlHAlign.xlHAlignCenter

        workSheet.get_Range("C7").Value = "xlHAlignRight"
        workSheet.get_Range("C7").HorizontalAlignment = XlHAlign.xlHAlignRight

        workSheet.get_Range("D7").Value = "xlHAlignJustify"
        workSheet.get_Range("D7").HorizontalAlignment = XlHAlign.xlHAlignJustify

        workSheet.get_Range("E7").Value = "xlHAlignDistributed"
        workSheet.get_Range("E7").HorizontalAlignment = XlHAlign.xlHAlignDistributed

        ' VerticalAlignment
        workSheet.get_Range("A9").Value = "xlVAlignTop"
        workSheet.get_Range("A9").VerticalAlignment = XlVAlign.xlVAlignTop

        workSheet.get_Range("B9").Value = "xlVAlignCenter"
        workSheet.get_Range("B9").VerticalAlignment = XlVAlign.xlVAlignCenter

        workSheet.get_Range("C9").Value = "xlVAlignBottom"
        workSheet.get_Range("C9").VerticalAlignment = XlVAlign.xlVAlignBottom

        workSheet.get_Range("D9").Value = "xlVAlignDistributed"
        workSheet.get_Range("D9").VerticalAlignment = XlVAlign.xlVAlignDistributed

        workSheet.get_Range("E9").Value = "xlVAlignJustify"
        workSheet.get_Range("E9").VerticalAlignment = XlVAlign.xlVAlignJustify

        ' setup rows and columns
        workSheet.Columns(1, Missing.Value).AutoFit()
        workSheet.Columns(2, Missing.Value).ColumnWidth = 25
        workSheet.Columns(3, Missing.Value).ColumnWidth = 25
        workSheet.Columns(4, Missing.Value).ColumnWidth = 25
        workSheet.Columns(5, Missing.Value).ColumnWidth = 25
        workSheet.Rows(9, Missing.Value).RowHeight = 25

        ' save the book 
        Dim fileExtension As String = GetDefaultExtension(excelApplication)
        Dim workbookFile As String = String.Format("{0}\Example02{1}", Environment.CurrentDirectory, fileExtension)
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
