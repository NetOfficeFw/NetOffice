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

        ' do background color for cells
        Dim listSeperator As String = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator

        ' draw a smile face
        Dim rangeAdressFace As String = String.Format("$C10:$M10{0}$C30:$M30{0}$C11:$C30{0}$M11:$M30", listSeperator)
        workSheet.get_Range(rangeAdressFace).Interior.Color = ToDouble(Color.DarkGreen)

        Dim rangeAdressEyes As String = String.Format("$F14{0}$J14", listSeperator)
        workSheet.get_Range(rangeAdressEyes).Interior.Color = ToDouble(Color.Black)

        Dim rangeAdressNoise As String = String.Format("$G18:$I19", listSeperator)
        workSheet.get_Range(rangeAdressNoise).Interior.Color = ToDouble(Color.DarkGreen)

        Dim rangeAdressMouth As String = String.Format("$F26{0}$J26{0}$G27:$I27", listSeperator)
        workSheet.get_Range(rangeAdressMouth).Interior.Color = ToDouble(Color.DarkGreen)

        ' border the face with the border arround method
        workSheet.get_Range(rangeAdressFace).BorderAround(XlLineStyle.xlDashDot, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic)
        workSheet.get_Range(rangeAdressEyes).BorderAround(XlLineStyle.xlDashDot, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic)
        workSheet.get_Range(rangeAdressNoise).BorderAround(XlLineStyle.xlDouble, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic)

        ' border explicitly
        workSheet.get_Range(rangeAdressMouth).Borders(XlBordersIndex.xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
        workSheet.get_Range(rangeAdressMouth).Borders(XlBordersIndex.xlEdgeBottom).Weight = 4
        workSheet.get_Range(rangeAdressMouth).Borders(XlBordersIndex.xlEdgeBottom).Color = ToDouble(Color.Black)

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
