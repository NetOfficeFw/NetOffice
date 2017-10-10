Imports ExampleBase
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.OfficeApi.Enums
Imports NetOffice.ExcelApi.Tools.Contribution

''' <summary>
''' Example 4 - Shapes, WordArts, Pictures, 3D-Effects
''' </summary>
Public Class Example04
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start excel and turn off msg boxes
        Dim excelApplication As New Excel.Application()
        excelApplication.DisplayAlerts = False

        ' create a utils instance, no need for but helpful to keep the lines of code low
        Dim utils As CommonUtils = New CommonUtils(excelApplication)

        ' add a new workbook
        Dim workBook As Excel.Workbook = excelApplication.Workbooks.Add()
        Dim workSheet As Excel.Worksheet = workBook.Worksheets(1)

        workSheet.Cells(1, 1).Value = "these sample shapes was dynamicly created by code."

        ' create a star
        Dim starShape As Excel.Shape = workSheet.Shapes.AddShape(MsoAutoShapeType.msoShape32pointStar, 10, 50, 200, 20)

        'create a simple textbox
        Dim textBox As Excel.Shape = workSheet.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 10, 150, 200, 50)
        textBox.TextFrame.Characters().Text = "text"
        textBox.TextFrame.Characters().Font.Size = 14

        ' create a wordart
        Dim textEffect As Excel.Shape = workSheet.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect14, "WordArt", "Arial", 12,
                                                                                MsoTriState.msoTrue, MsoTriState.msoFalse, 10, 250)

        ' create text effect
        Dim textDiagram As Excel.Shape = workSheet.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect11, "Effect", "Arial", 14,
                                                                     MsoTriState.msoFalse, MsoTriState.msoFalse, 10, 350)


        ' save the book 
        Dim workbookFile As String = utils.File.Combine(_hostApplication.RootDirectory, "Example04", DocumentFormat.Normal)
        workBook.SaveAs(workbookFile)

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

        ' show dialog for the user(you!)
        _hostApplication.ShowFinishDialog(Nothing, workbookFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example04"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Shapes, WordArts, Pictures, 3D-Effects"
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

End Class
