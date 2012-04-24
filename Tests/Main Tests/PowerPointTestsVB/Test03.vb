Imports NetOffice
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Enums
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums
Imports Tests.Core

Public Class Test03
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Create blend animation."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test03"
        End Get
    End Property

    Public ReadOnly Property OfficeProduct As String Implements Tests.Core.ITestPackage.OfficeProduct
        Get
            Return "PowerPoint"
        End Get
    End Property

    Public Function DoTest() As Tests.Core.TestResult Implements Tests.Core.ITestPackage.DoTest

        Dim application As PowerPoint.Application = Nothing
        Dim startTime As DateTime = DateTime.Now
        Try
            application = New PowerPoint.Application()

            ' add a new presentation with two new slides
            Dim presentation As PowerPoint.Presentation = application.Presentations.Add(MsoTriState.msoTrue)
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
