Imports ExampleBase
Imports NetOffice
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Enums
Imports NetOffice.OfficeApi.Enums
Imports NetOffice.PowerPointApi.Tools.Contribution

''' <summary>
''' Example 5 - Create ole chart object
''' </summary>
Public Class Example05
    Implements IExample

    Dim _hostApplication As ExampleBase.IHost

    Public Sub RunExample() Implements ExampleBase.IExample.RunExample

        ' start powerpoint
        Dim powerApplication As New PowerPoint.Application()

        ' create a utils instance, no need for but helpful to keep the lines of code low
        Dim utils As CommonUtils = New CommonUtils(powerApplication)

        ' add a new presentation with one new slide
        Dim presentation As PowerPoint.Presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue)
        Dim slide As PowerPoint.Slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank)

        ' add a chart
        slide.Shapes.AddOLEObject(120, 111, 480, 320, "MSGraph.Chart", "", MsoTriState.msoFalse, "", 0, "", MsoTriState.msoFalse)

        ' save the document 
        Dim documentFile As String = utils.File.Combine(_hostApplication.RootDirectory, "Example05", DocumentFormat.Normal)
        presentation.SaveAs(documentFile)

        ' close power point and dispose reference
        powerApplication.Quit()
        powerApplication.Dispose()

        ' show end dialog
        _hostApplication.ShowFinishDialog(Nothing, documentFile)

    End Sub

    Public ReadOnly Property Caption As String Implements ExampleBase.IExample.Caption
        Get
            Return "Example05"
        End Get
    End Property

    Public ReadOnly Property Description As String Implements ExampleBase.IExample.Description
        Get
            Return "Create OLE chart object"
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