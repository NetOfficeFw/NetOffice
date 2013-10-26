Imports System.Drawing
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports Tests.Core

Public Class Test02
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Set alignment and font style."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test02"
        End Get
    End Property

    Public ReadOnly Property OfficeProduct As String Implements Tests.Core.ITestPackage.OfficeProduct
        Get
            Return "Excel"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As Excel.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New NetOffice.ExcelApi.Application()
            application.DisplayAlerts = False
            application.Workbooks.Add()

            Dim workSheet As Excel.Worksheet = application.Workbooks(1).Sheets(1)

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

            Return New TestResult(True, DateTime.Now.Subtract(startTime), "", Nothing, "")

        Catch ex As Exception

            Return New TestResult(False, DateTime.Now.Subtract(startTime), ex.Message, ex, "")

        Finally

            If Not IsNothing(application) Then
                application.Quit()
                application.Dispose()
            End If

        End Try

    End Function
End Class
