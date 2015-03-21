Imports ExampleBase
Imports NetOffice
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Enums
Imports NetOffice.OfficeApi.Enums

Public Class Example02
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

#Region "IExample Member"

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start powerpoint
        Dim powerApplication As New PowerPoint.Application()

        ' create a utils instance, not need for but helpful to keep the lines of code low
        Dim utils As PowerPoint.Tools.CommonUtils = New PowerPoint.Tools.CommonUtils(powerApplication)

        ' add a new presentation with one new slide
        Dim presentation As PowerPoint.Presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue)
        Dim slide As PowerPoint.Slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank)

        ' add a label
        Dim label As PowerPoint.Shape = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, 10, 10, 600, 20)
        label.TextFrame.TextRange.Text = "This slide and created Shapes are created by NetOffice example."

        ' add a line
        slide.Shapes.AddLine(10, 80, 700, 80)

        ' add a wordart
        slide.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect9, "This a WordArt", "Arial", 20, _
                                           MsoTriState.msoTrue, MsoTriState.msoFalse, 10, 150)

        ' add a star
        slide.Shapes.AddShape(MsoAutoShapeType.msoShape24pointStar, 200, 200, 250, 250)

        ' save the document 
        Dim documentFile As String = utils.File.Combine(_hostApplication.RootDirectory, "Example02", PowerPoint.Tools.DocumentFormat.Normal)
        presentation.SaveAs(documentFile)

        ' close power point and dispose reference
        powerApplication.Quit()
        powerApplication.Dispose()

        ' show dialog for the user(you!)
        _hostApplication.ShowFinishDialog(Nothing, documentFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return IIf(_hostApplication.LCID = 1033, "Example02", "Beispiel02")
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return IIf(_hostApplication.LCID = 1033, "Create some kind of shapes", "Verschiede Shapes erstellen")
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

End Class