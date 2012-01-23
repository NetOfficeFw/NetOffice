Imports LateBindingApi.Core
Imports PowerPoint = NetOffice.PowerPointApi
Imports NetOffice.PowerPointApi.Enums
Imports NetOffice.OfficeApi.Enums

Public Class Form1

    Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click

        ' Initialize NetOffice
        LateBindingApi.Core.Factory.Initialize()

        ' start powerpoint
        Dim powerApplication As New PowerPoint.Application()

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
        Dim fileExtension As String = GetDefaultExtension(powerApplication)
        Dim documentFile As String = String.Format("{0}\\Example02{1}", Application.StartupPath, fileExtension)
        presentation.SaveAs(documentFile, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue)

        ' close power point and dispose reference
        powerApplication.Quit()
        powerApplication.Dispose()

        Dim fDialog As New FinishDialog("Presentation saved.", documentFile)
        fDialog.ShowDialog(Me)

    End Sub

#Region "Helper"

    ''' <summary>
    ''' returns the valid file extension for the instance. for example ".ppt" or ".pptx"
    ''' </summary>
    ''' <param name="application">the instance</param>
    ''' <returns>the extension</returns>
    ''' <remarks></remarks>
    Private Function GetDefaultExtension(ByVal application As PowerPoint.Application) As String

        Dim version As Double = application.Version
        If (version >= 120.0) Then
            Return ".pptx"
        Else
            Return ".ppt"
        End If

    End Function

#End Region

End Class
