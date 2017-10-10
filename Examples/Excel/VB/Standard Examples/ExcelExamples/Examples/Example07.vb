Imports ExampleBase
Imports NetOffice
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports VB = NetOffice.VBIDEApi
Imports NetOffice.VBIDEApi.Enums
Imports NetOffice.ExcelApi.Tools.Contribution

''' <summary>
''' Example 7 - Attach VBA Code to a workbook
''' </summary>
Public Class Example07
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        Dim isFailed = False
        Dim workbookFile As String = ""
        Dim excelApplication As Excel.Application = Nothing

        Try

            ' start excel and turn off msg boxes
            excelApplication = New Excel.Application()
            excelApplication.DisplayAlerts = False
            excelApplication.Visible = False

            ' create a utils instance, no need for but helpful to keep the lines of code low
            Dim utils As CommonUtils = New CommonUtils(excelApplication)

            ' add a new workbook
            Dim workBook As Excel.Workbook = excelApplication.Workbooks.Add()

            ' add new global Code Module
            Dim globalModule As VB.VBComponent = workBook.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule)
            globalModule.Name = "MyNewCodeModule"

            ' add a new procedure to the modul
            globalModule.CodeModule.InsertLines(1, "Public Sub HelloWorld(Param as string)" & vbNewLine & " MsgBox ""Hello from NetOffice!"" & vbnewline & Param" & vbNewLine & "End Sub")

            ' create a click event trigger for the first worksheet
            Dim linePosition As Integer = workBook.VBProject.VBComponents.Item(2).CodeModule.CreateEventProc("BeforeDoubleClick", "Worksheet")
            workBook.VBProject.VBComponents.Item(2).CodeModule.InsertLines(linePosition + 1, "HelloWorld ""BeforeDoubleClick""")

            ' display info in the worksheet
            Dim sheet As Excel.Worksheet = workBook.Worksheets(1)

            sheet.Cells(2, 2).Value = "This workbook contains dynamic created VBA Moduls and Event Code"
            sheet.Cells(5, 2).Value = "Open the VBA Editor to see the code"
            sheet.Cells(8, 2).Value = "Do a double click to catch the BeforeDoubleClick Event from this Worksheet."

            ' save the book 
            Dim fileFormat As XlFileFormat = GetFileFormat(excelApplication)
            workbookFile = utils.File.Combine(_hostApplication.RootDirectory, "Example07", DocumentFormat.Macros)
            workBook.SaveAs(workbookFile, fileFormat)

        Catch throwedException As System.Runtime.InteropServices.COMException

            isFailed = True
            _hostApplication.ShowErrorDialog("VBA Error", throwedException)

        Finally

            ' close excel and dispose reference
            excelApplication.Quit()
            excelApplication.Dispose()

            If (Not IsNothing(workbookFile) And Not isFailed) Then
                ' show dialog for the user(you!)
                _hostApplication.ShowFinishDialog(Nothing, workbookFile)
            End If

        End Try

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example07"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Attach VBA Code to a workbook. The option 'Trust Visual Basic projects' must be set."
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

    Private Function GetFileFormat(ByVal application As Excel.Application) As XlFileFormat

        Dim version As Double = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture)
        If (version >= 12.0) Then
            Return XlFileFormat.xlOpenXMLWorkbookMacroEnabled
        Else
            Return XlFileFormat.xlExcel7
        End If

    End Function

End Class
