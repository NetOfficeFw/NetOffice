Imports ExampleBase
Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums

Public Class Example02
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        ' start excel and turn off msg boxes
        Dim excelApplication As New Excel.Application()
        excelApplication.DisplayAlerts = False

        ' add a new workbook
        Dim workBook As Excel.Workbook = excelApplication.Workbooks.Add()
        Dim workSheet As Excel.Worksheet = workBook.Worksheets(1)

        ' font action
        workSheet.Range("A1").Value = "Arial Size:8 Bold Italic Underline"
        workSheet.Range("A1").Font.Name = "Arial"
        workSheet.Range("A1").Font.Size = 8
        workSheet.Range("A1").Font.Bold = True
        workSheet.Range("A1").Font.Italic = True
        workSheet.Range("A1").Font.Underline = True
        workSheet.Range("A1").Font.Color = Color.Violet.ToArgb()

        workSheet.Range("A3").Value = "Times New Roman Size:10"
        workSheet.Range("A3").Font.Name = "Times New Roman"
        workSheet.Range("A3").Font.Size = 10
        workSheet.Range("A3").Font.Color = Color.Orange.ToArgb()

        workSheet.Range("A5").Value = "Comic Sans MS Size:12 WrapText"
        workSheet.Range("A5").Font.Name = "Comic Sans MS"
        workSheet.Range("A5").Font.Size = 12
        workSheet.Range("A5").WrapText = True
        workSheet.Range("A5").Font.Color = Color.Navy.ToArgb()

        ' HorizontalAlignment
        workSheet.Range("A7").Value = "xlHAlignLeft"
        workSheet.Range("A7").HorizontalAlignment = XlHAlign.xlHAlignLeft

        workSheet.Range("B7").Value = "xlHAlignCenter"
        workSheet.Range("B7").HorizontalAlignment = XlHAlign.xlHAlignCenter

        workSheet.Range("C7").Value = "xlHAlignRight"
        workSheet.Range("C7").HorizontalAlignment = XlHAlign.xlHAlignRight

        workSheet.Range("D7").Value = "xlHAlignJustify"
        workSheet.Range("D7").HorizontalAlignment = XlHAlign.xlHAlignJustify

        workSheet.Range("E7").Value = "xlHAlignDistributed"
        workSheet.Range("E7").HorizontalAlignment = XlHAlign.xlHAlignDistributed

        ' VerticalAlignment
        workSheet.Range("A9").Value = "xlVAlignTop"
        workSheet.Range("A9").VerticalAlignment = XlVAlign.xlVAlignTop

        workSheet.Range("B9").Value = "xlVAlignCenter"
        workSheet.Range("B9").VerticalAlignment = XlVAlign.xlVAlignCenter

        workSheet.Range("C9").Value = "xlVAlignBottom"
        workSheet.Range("C9").VerticalAlignment = XlVAlign.xlVAlignBottom

        workSheet.Range("D9").Value = "xlVAlignDistributed"
        workSheet.Range("D9").VerticalAlignment = XlVAlign.xlVAlignDistributed

        workSheet.Range("E9").Value = "xlVAlignJustify"
        workSheet.Range("E9").VerticalAlignment = XlVAlign.xlVAlignJustify

        ' setup rows and columns
        workSheet.Columns(1).AutoFit()
        workSheet.Columns(2).ColumnWidth = 25
        workSheet.Columns(3).ColumnWidth = 25
        workSheet.Columns(4).ColumnWidth = 25
        workSheet.Columns(5).ColumnWidth = 25
        workSheet.Rows(9).RowHeight = 25

        ' save the book 
        Dim fileExtension As String = GetDefaultExtension(excelApplication)
        Dim workbookFile As String = String.Format("{0}\Example02{1}", _hostApplication.RootDirectory, fileExtension)
        workBook.SaveAs(workbookFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlExclusive)

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

        ' show dialog for the user(you!)
        _hostApplication.ShowFinishDialog(Nothing, workbookFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example02", "Beispiel02")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Font Attributes and Alignment for Cells", "Schrifteigenschaften und Ausrichtung in Zellen")
        End Get
    End Property

    Public Sub Connect(ByVal hostApplication As ExampleBase.IHost) Implements ExampleBase.IExample.Connect

        _hostApplication = hostApplication

    End Sub

    Public ReadOnly Property Panel As System.Windows.Forms.UserControl Implements ExampleBase.IExample.Panel
        Get
            Return Nothing
        End Get
    End Property

#End Region

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
