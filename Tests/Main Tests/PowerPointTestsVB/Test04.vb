Imports NetOffice
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Enums
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums
Imports Tests.Core

Public Class Test04
    Implements ITestPackage

    Public ReadOnly Property Description As String Implements Tests.Core.ITestPackage.Description
        Get
            Return "Create a chart."
        End Get
    End Property

    Public ReadOnly Property Language As String Implements Tests.Core.ITestPackage.Language
        Get
            Return "VB"
        End Get
    End Property

    Public ReadOnly Property Name As String Implements Tests.Core.ITestPackage.Name
        Get
            Return "Test04"
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

            ' add a new presentation with one new slide
            Dim presentation As PowerPoint.Presentation = application.Presentations.Add(MsoTriState.msoTrue)
            Dim slide As PowerPoint.Slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank)

            ' add a chart. MSGraph.Chart is may not installed
            slide.Shapes.AddOLEObject(120, 111, 480, 320, "MSGraph.Chart", "", MsoTriState.msoFalse, "", 0, "", MsoTriState.msoFalse)

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
