Imports System.Reflection

Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports NetOffice.OfficeApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        LateBindingApi.Core.Factory.Initialize()

        ' start excel and turn off msg boxes
        Dim excelApplication As New Excel.Application()
        excelApplication.DisplayAlerts = False

        ' add a new workbook
        Dim workBook As Excel.Workbook = excelApplication.Workbooks.Add()
        Dim workSheet As Excel.Worksheet = workBook.Worksheets(1)

        workSheet.Cells(1, 1).Value = "these sample shapes was dynamicly created by code."

        ' create a star
        Dim starShape As Excel.Shape = workSheet.Shapes.AddShape(MsoAutoShapeType.msoShape32pointStar, 10, 50, 200, 20)

        'create a simple textbox
        Dim textBox As Excel.Shape = workSheet.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 10, 150, 200, 50)
        TextBox.TextFrame.Characters().Text = "text"
        TextBox.TextFrame.Characters().Font.Size = 14

        ' create a wordart
        Dim textEffect As Excel.Shape = workSheet.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect14, "WordArt", "Arial", 12, _
                                                                                MsoTriState.msoTrue, MsoTriState.msoFalse, 10, 250)

        ' create text effect
        Dim textDiagram As Excel.Shape = workSheet.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect11, "Effect", "Arial", 14, _
                                                                     MsoTriState.msoFalse, MsoTriState.msoFalse, 10, 350)


        ' save the book 
        Dim fileExtension As String = GetDefaultExtension(excelApplication)
        Dim workbookFile As String = String.Format("{0}\Example04{1}", Application.StartupPath, fileExtension)
        workBook.SaveAs(workbookFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlExclusive)

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

        Dim fDialog As New FinishDialog("Workbook saved.", workbookFile)
        fDialog.ShowDialog(Me)

    End Sub

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
