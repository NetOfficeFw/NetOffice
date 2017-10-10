Imports ExampleBase
Imports NetOffice
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Enums
Imports NetOffice.OfficeApi.Enums
Imports NetOffice.PowerPointApi.Tools.Contribution

''' <summary>
''' Example 1 - Create a presentation with 1 empty slide
''' </summary>
Public Class Example01
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start powerpoint
        Dim powerApplication As PowerPoint.Application = New PowerPoint.Application()

        ' create a utils instance, no need for but helpful to keep the lines of code low
        Dim utils As CommonUtils = New CommonUtils(powerApplication)

        ' add a new presentation with one new slide
        Dim presentation As PowerPoint.Presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue)
        presentation.Slides.Add(1, PpSlideLayout.ppLayoutClipArtAndVerticalText)

        ' save the document 
        Dim documentFile As String = utils.File.Combine(_hostApplication.RootDirectory, "Example01", DocumentFormat.Normal)
        presentation.SaveAs(documentFile)

        ' close power point and dispose reference
        powerApplication.Quit()
        powerApplication.Dispose()

        ' show end dialog
        _hostApplication.ShowFinishDialog(Nothing, documentFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example01"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Create a presentation with 1 empty slide"
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