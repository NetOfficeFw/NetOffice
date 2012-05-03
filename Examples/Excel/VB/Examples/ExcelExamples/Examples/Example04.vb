Imports ExampleBase
Imports NetOffice
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports NetOffice.OfficeApi.Enums

Public Class Example04
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

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
        textBox.TextFrame.Characters().Text = "text"
        textBox.TextFrame.Characters().Font.Size = 14

        ' create a wordart
        Dim textEffect As Excel.Shape = workSheet.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect14, "WordArt", "Arial", 12, _
                                                                                MsoTriState.msoTrue, MsoTriState.msoFalse, 10, 250)

        ' create text effect
        Dim textDiagram As Excel.Shape = workSheet.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect11, "Effect", "Arial", 14, _
                                                                     MsoTriState.msoFalse, MsoTriState.msoFalse, 10, 350)


        ' save the book 
        Dim fileExtension As String = GetDefaultExtension(excelApplication)
        Dim workbookFile As String = String.Format("{0}\Example04{1}", _hostApplication.RootDirectory, fileExtension)
        workBook.SaveAs(workbookFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlExclusive)

        ' close excel and dispose reference
        excelApplication.Quit()
        excelApplication.Dispose()

        ' show dialog for the user(you!)
        _hostApplication.ShowFinishDialog(Nothing, workbookFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example04", "Beispiel04")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Shapes, WordArts, Pictures, 3D-Effects", "Shapes, WordArts, Pictures, 3D-Effects")
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

        Dim version As Double = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture)
        If (version >= 12.0) Then
            Return ".xlsx"
        Else
            Return ".xls"
        End If

    End Function

#End Region

End Class
