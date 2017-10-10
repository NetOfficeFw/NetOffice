Imports ExampleBase
Imports NetOffice
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Enums
Imports NetOffice.OfficeApi.Enums
Imports NetOffice.PowerPointApi.Tools.Contribution

''' <summary>
''' Example 4 - Create blend animation 
''' </summary>
Public Class Example04
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start powerpoint
        Dim powerApplication As New PowerPoint.Application()

        ' create a utils instance, no need for but helpful to keep the lines of code low
        Dim utils As CommonUtils = New CommonUtils(powerApplication)

        ' add a new presentation with two new slides
        Dim presentation As PowerPoint.Presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue)
        Dim slide1 As PowerPoint.Slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank)
        Dim slide2 As PowerPoint.Slide = presentation.Slides.Add(2, PpSlideLayout.ppLayoutBlank)

        ' add shapes
        slide1.Shapes.AddShape(MsoAutoShapeType.msoShape4pointStar, 100, 100, 200, 200)
        slide2.Shapes.AddShape(MsoAutoShapeType.msoShapeDoubleWave, 200, 200, 200, 200)

        ' change blend animation
        slide1.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectCoverDown
        slide1.SlideShowTransition.Speed = PpTransitionSpeed.ppTransitionSpeedFast

        slide2.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectCoverLeftDown
        slide2.SlideShowTransition.Speed = PpTransitionSpeed.ppTransitionSpeedFast

        ' save the document 
        Dim documentFile As String = utils.File.Combine(_hostApplication.RootDirectory, "Example04", DocumentFormat.Normal)
        presentation.SaveAs(documentFile)

        ' close power point and dispose reference
        powerApplication.Quit()
        powerApplication.Dispose()

        ' show end dialog
        _hostApplication.ShowFinishDialog(Nothing, documentFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example04"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Create blend animation"
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