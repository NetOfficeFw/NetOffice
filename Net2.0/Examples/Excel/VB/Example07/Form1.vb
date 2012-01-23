Imports System.Reflection

Imports LateBindingApi.Core
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports VBE = NetOffice.VBIDEApi
Imports NetOffice.VBIDEApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        Dim workbookFile As String = ""
        Dim excelApplication As Excel.Application = Nothing

        Try

            ' start excel and turn off msg boxes
            excelApplication = New Excel.Application()
            excelApplication.DisplayAlerts = False
            excelApplication.Visible = False

            ' add a new workbook
            Dim workBook As Excel.Workbook = excelApplication.Workbooks.Add()

            ' add new global Code Module
            Dim globalModule As VBE.VBComponent = workBook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
            globalModule.Name = "MyNewCodeModule"

            ' add a new procedure to the modul
            globalModule.CodeModule.InsertLines(1, "Public Sub HelloWorld(Param as string)" & vbNewLine & " MsgBox ""Hello World!"" & vbnewline & Param" & vbNewLine & "End Sub")

            ' create a click event trigger for the first worksheet
            Dim linePosition As Integer = workBook.VBProject.VBComponents.Item(2).CodeModule.CreateEventProc("BeforeDoubleClick", "Worksheet")
            workBook.VBProject.VBComponents.Item(2).CodeModule.InsertLines(linePosition + 1, "HelloWorld ""BeforeDoubleClick""")

            ' display info in the worksheet
            Dim sheet As Excel.Worksheet = workBook.Worksheets(1)

            sheet.Cells(2, 2).Value = "This workbook contains dynamic created VBA Moduls and Event Code"
            sheet.Cells(5, 2).Value = "Open the VBA Editor to see the code"
            sheet.Cells(8, 2).Value = "Do a double click to catch the BeforeDoubleClick Event from this Worksheet."

            ' save the book 
            Dim fileExtension As String = GetDefaultExtension(excelApplication)
            workbookFile = String.Format("{0}\Example07{1}", Application.StartupPath, fileExtension)
            workBook.SaveAs(workbookFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlExclusive)

        Catch throwedException As System.Runtime.InteropServices.COMException

            Dim message As String = String.Format("An error is occured.{0}ExceptionTrace:{0}", Environment.NewLine)

            Dim exception As Exception = throwedException
            While (Not IsNothing(exception))
                message += String.Format("{0}{1}", exception.Message, Environment.NewLine)
                exception = exception.InnerException
            End While

            MessageBox.Show(message)

        Finally

            ' close excel and dispose reference
            excelApplication.Quit()
            excelApplication.Dispose()

            If (Not IsNothing(workbookFile)) Then
                Dim fDialog As FinishDialog = New FinishDialog("Workbook saved.", workbookFile)
                fDialog.ShowDialog(Me)
            End If


        End Try

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
